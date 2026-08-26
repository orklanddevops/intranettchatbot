import argparse
import uuid

from azure.identity import AzureDeveloperCliCredential
import urllib3


def update_redirect_uris(credential, app_id, uri):
    urllib3.request(
        "PATCH",
        f"https://graph.microsoft.com/v1.0/applications/{app_id}",
        headers={
            "Authorization": "Bearer "
            + credential.get_token("https://graph.microsoft.com/.default").token,
        },
        json={
            "web": {
                "redirectUris": [
                    "http://localhost:5000/.auth/login/aad/callback",
                    f"{uri}/.auth/login/aad/callback",
                ]
            }
        },
    )


def get_auth_headers(credential):
    return {
        "Authorization": "Bearer "
        + credential.get_token("https://graph.microsoft.com/.default").token,
    }


def ensure_application_api_scope(credential, app_id):
    app_response = urllib3.request(
        "GET",
        f"https://graph.microsoft.com/v1.0/applications/{app_id}",
        headers=get_auth_headers(credential),
    )
    app = app_response.json()
    client_id = app["appId"]
    scopes = app.get("api", {}).get("oauth2PermissionScopes", [])
    identifier_uris = app.get("identifierUris", [])
    default_identifier_uri = f"api://{client_id}"
    if default_identifier_uri not in identifier_uris:
        identifier_uris.append(default_identifier_uri)

    if not any(scope.get("value") == "access_as_user" for scope in scopes):
        scopes.append(
            {
                "id": str(uuid.uuid4()),
                "adminConsentDescription": "Allow SharePoint to call the intranet chatbot as the signed-in user.",
                "adminConsentDisplayName": "Access intranet chatbot",
                "isEnabled": True,
                "type": "User",
                "userConsentDescription": "Allow SharePoint to call the intranet chatbot as you.",
                "userConsentDisplayName": "Access intranet chatbot",
                "value": "access_as_user",
            }
        )

    urllib3.request(
        "PATCH",
        f"https://graph.microsoft.com/v1.0/applications/{app_id}",
        headers=get_auth_headers(credential),
        json={
            "identifierUris": identifier_uris,
            "api": {
                "oauth2PermissionScopes": scopes,
            },
        },
    )

    return client_id


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Add a redirect URI to a registered application",
        epilog="Example: auth_update.py --appid 123 --uri https://abc.azureservices.net",
    )
    parser.add_argument(
        "--appid",
        required=False,
        help="Required. ID of the application to update.",
    )
    parser.add_argument(
        "--uri",
        required=False,
        help="Required. URI of the deployed application.",
    )
    args = parser.parse_args()

    credential = AzureDeveloperCliCredential()

    print(
        f"Updating application registration {args.appid} with redirect URI for {args.uri}"
    )
    update_redirect_uris(credential, args.appid, args.uri)
    client_id = ensure_application_api_scope(credential, args.appid)
    print(f"Configured API scope api://{client_id}/access_as_user")
