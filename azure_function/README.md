# Conversation API proxy

This Node.js Azure Function exposes `/api/conversation` for the SPFx chatbot.

The route is anonymous at the Function host layer so browser CORS preflight can complete, but `POST` requests are rejected unless they include a valid Entra bearer token for the chatbot API scope.

The Function App should use the Node.js 24 stack. This project uses the Azure Functions Node.js v4 programming model with `@azure/functions`.

Required Function App settings:

```text
AUTH_CLIENT_ID=47cbcbfe-6efd-4113-b089-0dcb7c7b33bc
AUTH_ISSUER_URI=https://login.microsoftonline.com/91d14ddc-e451-4a63-83a5-7c9dc6335632/v2.0
AZURE_TENANT_ID=91d14ddc-e451-4a63-83a5-7c9dc6335632
BACKEND_CONVERSATION_URL=https://intranettchatbot.orkland.kommune.no/api/conversation
BACKEND_REQUEST_TIMEOUT_SECONDS=230
CHATBOT_ALLOWED_ORIGINS=https://orkland.sharepoint.com
CHATBOT_AUTH_AUDIENCES=api://47cbcbfe-6efd-4113-b089-0dcb7c7b33bc,47cbcbfe-6efd-4113-b089-0dcb7c7b33bc
CHATBOT_REQUIRED_SCOPE=access_as_user
```

The SPFx package currently calls:

```text
https://intranettchatbot-conversation-gkeuczhbgvczddfp.norwayeast-01.azurewebsites.net/api/conversation
```

Deploy this folder through the `api-proxy` service in `azure.yaml`, or deploy `azure_function` directly to the Function App.
