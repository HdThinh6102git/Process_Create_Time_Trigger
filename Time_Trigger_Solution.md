# UiPath Time Trigger Management Solution

This document outlines the end-to-end design, workflow steps, variables, and VB.NET code snippets to manage Time Triggers in UiPath Orchestrator using the API.

---

## 🏗️ 1. UiPath Workflow Design

### Workflow Structure (Sequence Layout)

1. **Initialization Sequence**
   - **Excel Read**: Use *Read Range Workbook* to load the config file into `dt_TriggersInput`.
   - **Initialize Output Report**: *Build Data Table* with columns `TriggerName`, `Status`, `Message` → `dt_Report`.
   - **Authenticate**: Use *HTTP Request* to acquire the OAuth2 `access_token` from UiPath Identity.

2. **Orchestrator Caching Sequence**
   - **GET Releases**: *HTTP Request* to `/odata/Releases` → `str_ReleasesJson`.
   - **GET Triggers**: *HTTP Request* to `/odata/ProcessSchedules` → `str_TriggersJson`.
   - **Cache Dictionary Generation**: Use *Invoke Code* with VB.NET to build:
     - `dict_Releases` (`ProcessName` → `ReleaseId`)
     - `dict_Triggers` (`TriggerName` → `TriggerId`)

3. **Loop & Processing Sequence**
   - **For Each Row in `dt_TriggersInput`**:
     - **Validation / Build Cron**: Use *Invoke Code* to validate the ProcessName existence and generate the Cron Expression.
     - **Decide Action**: Check if `dict_Triggers.ContainsKey(Row("TriggerName").ToString)`.
     - **If True (Update)**: *HTTP Request* (PUT) to `/odata/ProcessSchedules({Id})`.
     - **If False (Create)**: *HTTP Request* (POST) to `/odata/ProcessSchedules`.
     - **Exception Handling**: Try Catch around row actions to record Status (`Success`/`Failed`) and Message.
     - **Add to Report**: Append row status to `dt_Report`.

4. **Finalization Sequence**
   - **Write Range Workbook**: Export `dt_Report` to the Excel report file.

---

## 📦 2. Variables & Arguments

| Variable Name | Type | Scope | Description |
|---|---|---|---|
| `dt_TriggersInput` | `DataTable` | Main Sequence | Input trigger rows read from Excel |
| `dt_Report` | `DataTable` | Main Sequence | Output results table for logging |
| `str_AccessToken` | `String` | Main Sequence | Bearer OAuth token |
| `str_ReleasesJson` | `String` | Main Sequence | API response from GET Releases |
| `str_TriggersJson` | `String` | Main Sequence | API response from GET Triggers |
| `dict_Releases` | `Dictionary(Of String, Long)` | Main Sequence | ProcessName to ReleaseId mapping |
| `dict_Triggers` | `Dictionary(Of String, Long)` | Main Sequence | TriggerName to TriggerId mapping |
| `str_CronExpression` | `String` | For Each Row | Output Cron string computed by VB.NET |
| `lng_ReleaseId` | `Long` | For Each Row | Mapped ID for the trigger’s process |
| `lng_TriggerId` | `Long` | For Each Row | Existing trigger ID (if updating) |

---

## 💻 3. VB.NET Code Snippets (Invoke Code)

### A. Convert JSON to Dictionary

This activity maps current **Process Releases** and **Time Triggers** into memory dictionaries to prevent repeated API calls.

- **Arguments (`In/Out`)**:
  - `in_ReleasesJson` (In, String)
  - `in_TriggersJson` (In, String)
  - `out_DictReleases` (Out, `Dictionary(Of String, Long)`)
  - `out_DictTriggers` (Out, `Dictionary(Of String, Long)`)

```vbnet
out_DictReleases = New Dictionary(Of String, Long)(StringComparer.OrdinalIgnoreCase)
out_DictTriggers = New Dictionary(Of String, Long)(StringComparer.OrdinalIgnoreCase)

' 1. Parse Releases
If Not String.IsNullOrEmpty(in_ReleasesJson) Then
    Dim releasesObj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(in_ReleasesJson)
    Dim valueArray As Newtonsoft.Json.Linq.JArray = CType(releasesObj("value"), Newtonsoft.Json.Linq.JArray)
    If valueArray IsNot Nothing Then
        For Each item As Newtonsoft.Json.Linq.JObject In valueArray
            ' Mapping Name or ProcessKey to Release ID
            Dim pName As String = item("ProcessKey").ToString()
            Dim pId As Long = CType(item("Id"), Long)
            If Not out_DictReleases.ContainsKey(pName) Then
                out_DictReleases.Add(pName, pId)
            End If
        Next
    End If
End If

' 2. Parse Triggers
If Not String.IsNullOrEmpty(in_TriggersJson) Then
    Dim triggersObj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(in_TriggersJson)
    Dim valueArray As Newtonsoft.Json.Linq.JArray = CType(triggersObj("value"), Newtonsoft.Json.Linq.JArray)
    If valueArray IsNot Nothing Then
        For Each item As Newtonsoft.Json.Linq.JObject In valueArray
            Dim tName As String = item("Name").ToString()
            Dim tId As Long = CType(item("Id"), Long)
            If Not out_DictTriggers.ContainsKey(tName) Then
                out_DictTriggers.Add(tName, tId)
            End If
        Next
    End If
End If
```

---

### B. Validate Process & Generate Cron Expression (BuildCron)

Runs for each row. Ensures the Process exists in Orchestrator and compiles a standard Quartz Cron string.

- **Arguments (`In/Out`)**:
  - `in_ProcessName` (In, String)
  - `in_DictReleases` (In, `Dictionary(Of String, Long)`)
  - `in_Time` (In, String)
  - `in_Weekdays` (In, String)
  - `in_Frequency` (In, String)
  - `out_ReleaseId` (Out, Long)
  - `out_CronExpression` (Out, String)

```vbnet
' 1. Validate Process Existence
If Not in_DictReleases.ContainsKey(in_ProcessName.Trim()) Then
    Throw New Exception(String.Format("Process not found: {0}", in_ProcessName))
Else
    out_ReleaseId = in_DictReleases(in_ProcessName.Trim())
End If

' 2. Parse Time (supports both HH:mm and HH,mm formats)
Dim timeParts As String() = in_Time.Split(New Char() {":"c, ","c})
If timeParts.Length <> 2 Then
    Throw New ArgumentException("Time must be in 'HH:mm' or 'HH,mm' format. Received: " & in_Time)
End If

Dim hours As String = CInt(timeParts(0).Trim()).ToString()
Dim minutes As String = CInt(timeParts(1).Trim()).ToString()

' 3. Derive Weekday pattern
Dim daysOfWeek As String = If(String.IsNullOrWhiteSpace(in_Weekdays), "*", in_Weekdays.Trim().ToUpper())

If in_Frequency.Trim().Equals("Daily", StringComparison.OrdinalIgnoreCase) OrElse daysOfWeek = "*" OrElse daysOfWeek = "?" Then
    out_CronExpression = String.Format("0 {0} {1} ? * *", minutes, hours)
Else
    out_CronExpression = String.Format("0 {0} {1} ? * {2}", minutes, hours, daysOfWeek)
End If
```

---

## 📡 4. HTTP Request Configuration

### 🔑 Authentication Endpoint (POST)
- **Endpoint**: `https://cloud.uipath.com/identity_/connect/token`
- **Method**: `POST`
- **Headers**:
  - `Content-Type`: `application/x-www-form-urlencoded`
- **Body Content**:
  ```text
  grant_type=client_credentials&client_id=YOUR_CLIENT_ID&client_secret=YOUR_CLIENT_SECRET&scope=OR.Jobs OR.Schedules OR.Execution
  ```

---

### 📥 1. Releases API (GET)
- **Endpoint**: `https://cloud.uipath.com/{OrganizationName}/{TenantName}/orchestrator_/odata/Releases`
- **Method**: `GET`
- **Headers**:
  - `Authorization`: `"Bearer " + str_AccessToken`
  - `X-UIPATH-OrganizationUnitId`: `YOUR_FOLDER_ID` (If modern folders)

---

### 📥 2. Fetch Existing Triggers (GET)
- **Endpoint**: `https://cloud.uipath.com/{OrganizationName}/{TenantName}/orchestrator_/odata/ProcessSchedules`
- **Method**: `GET`
- **Headers**: Same as above

---

### 📤 3. Create Trigger API (POST)
- **Endpoint**: `https://cloud.uipath.com/{OrganizationName}/{TenantName}/orchestrator_/odata/ProcessSchedules`
- **Method**: `POST`
- **Headers**: Same as above + `Content-Type`: `application/json`
- **JSON Body**:
  ```json
  {
    "Name": "TriggerName",
    "ReleaseId": 123,
    "CronExpression": "0 20 14 ? * SUN,SAT",
    "TimeZoneId": "Korea Standard Time",
    "Enabled": true
  }
  ```

---

### 📤 4. Update Trigger API (PUT)
- **Endpoint**: `https://cloud.uipath.com/{OrganizationName}/{TenantName}/orchestrator_/odata/ProcessSchedules({Id})`
- **Method**: `PUT`
- **Headers**: Same as above + `Content-Type`: `application/json`
- **JSON Body**:
  ```json
  {
    "Id": 456,
    "Name": "TriggerName",
    "ReleaseId": 123,
    "CronExpression": "0 20 14 ? * SUN,SAT",
    "TimeZoneId": "Korea Standard Time",
    "Enabled": true
  }
  ```

---

## ⚠️ 5. Error Handling & Exception Management

To guarantee execution integrity:

1. **Global Process Level Try Catch**: Catches network/auth errors before starting rows.
2. **Per Row Try Catch**:
   - Place a **Try Catch** around individual row logic (Validations, Trigger Check, and API invocation).
   - If a specific row throws an error (e.g. `Process not found` or `Invalid Time format`), the **Catch** block assigns the `Status` as `Failed` and logs the `Exception.Message`.
   - Normal completions add a `Success` status.
   - At the end of the workflow, we overwrite/append the logs into `Report.xlsx` using *Write Range Workbook*.

---

## 🎁 6. Bonus Features & Design Suggestions

### 📊 Excel Input Validation Strategy
To prevent malformed inputs, apply these constraints in the config template. All data types in the Excel file are `Text`:
- **Frequency**: Add Excel Data Validation Dropdown: `Daily`, `Weekly`.
- **Weekdays**: Format as text with uppercase three-letter abbreviations separated by commas (e.g., `SUN,MON,TUE`).
- **Time**: Restrict string format as text using Custom Validation with formula `=AND(ISNUMBER(VALUE(LEFT(A1,2))), VALUE(LEFT(A1,2))<24, ISNUMBER(VALUE(RIGHT(A1,2))), VALUE(RIGHT(A1,2))<60)`.
- **Note on Enabled**: The `Enabled` status defaults to `true` by default in the API body and is not required as a column in the Excel file.

### 🚀 Performance Optimization (for 100+ triggers)
- **Cache Local Results**: Do not fetch releases or current triggers inside the loop. Retrieve them once at startup.
- **Batched API Checks**: Use OData filters `$filter` instead of bringing down massive JSON dumps if orchestrator releases scale to thousands. For example: `/odata/Releases?$select=Id,ProcessKey,Name`.
- **Parallel Processing**: For cloud environments or large numbers of rows, use a *Parallel For Each* to issue HTTP requests concurrently, but beware of API rate limits.
