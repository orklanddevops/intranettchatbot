import os
from functools import lru_cache
from typing import Any

import jwt
from jwt import PyJWKClient


class AuthenticationError(Exception):
    pass


def _get_header(request_headers: Any, name: str) -> str | None:
    if hasattr(request_headers, "get"):
        return request_headers.get(name)

    return None


def _split_csv(value: str | None) -> list[str]:
    if not value:
        return []

    return [item.strip() for item in value.split(",") if item.strip()]


def _tenant_id_from_issuer(issuer: str | None) -> str | None:
    if not issuer:
        return None

    issuer = issuer.rstrip("/")
    parts = issuer.split("/")
    if len(parts) >= 4 and parts[2].lower() == "login.microsoftonline.com":
        return parts[3]

    return None


def _accepted_audiences() -> list[str]:
    audiences = _split_csv(os.environ.get("CHATBOT_AUTH_AUDIENCES"))
    auth_client_id = os.environ.get("AUTH_CLIENT_ID")

    if auth_client_id:
        audiences.extend([auth_client_id, f"api://{auth_client_id}"])

    return sorted(set(audiences))


def _accepted_issuers(tenant_id: str) -> set[str]:
    issuers = set(_split_csv(os.environ.get("CHATBOT_AUTH_ISSUERS")))
    auth_issuer_uri = os.environ.get("AUTH_ISSUER_URI")

    if auth_issuer_uri:
        issuers.add(auth_issuer_uri)

    issuers.add(f"https://login.microsoftonline.com/{tenant_id}/v2.0")
    issuers.add(f"https://sts.windows.net/{tenant_id}/")

    return {issuer.rstrip("/") for issuer in issuers}


@lru_cache(maxsize=4)
def _get_jwk_client(tenant_id: str) -> PyJWKClient:
    return PyJWKClient(f"https://login.microsoftonline.com/{tenant_id}/discovery/v2.0/keys")


def _extract_bearer_token(request_headers: Any) -> str | None:
    authorization = _get_header(request_headers, "Authorization")
    if not authorization:
        return None

    scheme, _, token = authorization.partition(" ")
    if scheme.lower() != "bearer" or not token:
        return None

    return token


def _validate_bearer_token(token: str) -> dict:
    try:
        unverified_claims = jwt.decode(token, options={"verify_signature": False})
    except jwt.PyJWTError as exc:
        raise AuthenticationError("Invalid bearer token.") from exc

    configured_tenant_id = os.environ.get("AZURE_TENANT_ID") or _tenant_id_from_issuer(
        os.environ.get("AUTH_ISSUER_URI")
    )
    token_tenant_id = unverified_claims.get("tid")
    tenant_id = configured_tenant_id or token_tenant_id

    if not tenant_id:
        raise AuthenticationError("Token tenant could not be determined.")

    if configured_tenant_id and token_tenant_id and token_tenant_id != configured_tenant_id:
        raise AuthenticationError("Token tenant is not allowed.")

    audiences = _accepted_audiences()
    if not audiences:
        raise AuthenticationError("No chatbot token audience is configured.")

    try:
        signing_key = _get_jwk_client(tenant_id).get_signing_key_from_jwt(token)
        claims = jwt.decode(
            token,
            signing_key.key,
            algorithms=["RS256"],
            audience=audiences,
            options={"verify_iss": False},
        )
    except jwt.PyJWTError as exc:
        raise AuthenticationError("Bearer token validation failed.") from exc

    issuer = str(claims.get("iss", "")).rstrip("/")
    if issuer not in _accepted_issuers(tenant_id):
        raise AuthenticationError("Token issuer is not allowed.")

    return claims


def _user_from_bearer_claims(claims: dict, token: str) -> dict:
    user_id = claims.get("oid") or claims.get("sub")
    user_name = claims.get("preferred_username") or claims.get("upn") or claims.get("email") or claims.get("name")

    if not user_id:
        raise AuthenticationError("Token does not contain a usable user id.")

    return {
        "user_principal_id": user_id,
        "user_name": user_name,
        "auth_provider": "aad",
        "auth_token": token,
        "client_principal_b64": None,
        "aad_id_token": None,
        "claims": claims,
    }


def _user_from_easy_auth_headers(request_headers: Any) -> dict | None:
    if _get_header(request_headers, "X-Ms-Client-Principal-Id") is None:
        return None

    return {
        "user_principal_id": _get_header(request_headers, "X-Ms-Client-Principal-Id"),
        "user_name": _get_header(request_headers, "X-Ms-Client-Principal-Name"),
        "auth_provider": _get_header(request_headers, "X-Ms-Client-Principal-Idp"),
        "auth_token": _get_header(request_headers, "X-Ms-Token-Aad-Id-Token"),
        "client_principal_b64": _get_header(request_headers, "X-Ms-Client-Principal"),
        "aad_id_token": _get_header(request_headers, "X-Ms-Token-Aad-Id-Token"),
    }


def get_authenticated_user_details(request_headers, require_auth: bool = False):
    bearer_token = _extract_bearer_token(request_headers)
    if bearer_token:
        return _user_from_bearer_claims(_validate_bearer_token(bearer_token), bearer_token)

    easy_auth_user = _user_from_easy_auth_headers(request_headers)
    if easy_auth_user:
        return easy_auth_user

    if require_auth:
        raise AuthenticationError("Authentication is required.")

    from . import sample_user

    raw_user_object = sample_user.sample_user
    return {
        "user_principal_id": raw_user_object.get("X-Ms-Client-Principal-Id"),
        "user_name": raw_user_object.get("X-Ms-Client-Principal-Name"),
        "auth_provider": raw_user_object.get("X-Ms-Client-Principal-Idp"),
        "auth_token": raw_user_object.get("X-Ms-Token-Aad-Id-Token"),
        "client_principal_b64": raw_user_object.get("X-Ms-Client-Principal"),
        "aad_id_token": raw_user_object.get("X-Ms-Token-Aad-Id-Token"),
    }
