# How to use the Swagger Specification File

The file `Time_Trigger_API_Swagger.json` provides an OpenAPI 3.0-compliant schema for testing and interacting with the Orchestrator APIs. Follow the steps below to use it:

## 1. Viewing with Swagger UI / Editor (Web)
1. Navigate to the online [Swagger Editor](https://editor.swagger.io/).
2. From the top menu, click **File** -> **Import file**.
3. Select `Time_Trigger_API_Swagger.json` from your project directory.
4. The web view will render an interactive visual UI on the right side of the screen.

## 2. Testing directly in Swagger Editor
- Once the file is loaded, you can click on any API endpoint (e.g., `POST /identity_/connect/token`).
- Click the **Try it out** button.
- Modify the JSON body or fields if necessary.
- Click **Execute** to send requests to the server (you will need to provide valid credentials and access tokens).

## 3. Importing into Postman
Postman supports reading the OpenAPI specification file directly:
1. Open the Postman application.
2. Click the **Import** button at the top left of the screen.
3. Select `Time_Trigger_API_Swagger.json` and click open.
4. Postman will automatically convert the paths and schemas into a ready-to-test Collection with environment variables.
