"""NETSCOUT HTTP 인증 helper.

Basic Auth와 NSSESSIONID cookie 인증을 같은 호출부에서 사용할 수 있게
requests 인자 dict로 변환한다. password/token/cookie 값은 로그에 남기지 않는다.
"""

from __future__ import annotations

from typing import Any

from requests.auth import HTTPBasicAuth


TOKEN_AUTH_TYPES = {"nssessionid", "session_cookie", "cookie", "token"}


def normalized_auth_type(auth_type: str) -> str:
    """설정의 auth_type 값을 비교하기 쉬운 소문자 형태로 정규화한다."""
    return str(auth_type or "basic").strip().lower()


def is_token_auth(auth_type: str) -> bool:
    """NSSESSIONID cookie 기반 인증인지 확인한다."""
    return normalized_auth_type(auth_type) in TOKEN_AUTH_TYPES


def request_auth_kwargs(settings: Any) -> dict[str, Any]:
    """requests 호출에 넘길 인증 관련 keyword argument를 만든다."""
    return auth_kwargs_for_type(settings, settings.auth_type)


def data_request_auth_kwargs(settings: Any) -> dict[str, Any]:
    """Data API requests 호출에 넘길 인증 관련 keyword argument를 만든다."""
    return auth_kwargs_for_type(settings, getattr(settings, "data_auth_type", settings.auth_type))


def auth_kwargs_for_type(settings: Any, selected_auth_type: str) -> dict[str, Any]:
    """지정된 인증 방식에 맞는 requests keyword argument를 만든다."""
    auth_type = normalized_auth_type(selected_auth_type)
    if auth_type == "basic":
        return {"auth": HTTPBasicAuth(settings.username, settings.password)}
    if auth_type in TOKEN_AUTH_TYPES:
        return {"cookies": {"NSSESSIONID": settings.session_token}}
    raise ValueError(f"Unsupported auth_type: {settings.auth_type}")


def data_auth_log_label(settings: Any) -> str:
    """Data API 로그에 남길 인증 방식 이름을 반환한다."""
    return auth_log_label(getattr(settings, "data_auth_type", settings.auth_type))


def auth_log_label(auth_type: str) -> str:
    """로그에 남겨도 되는 인증 방식 이름만 반환한다."""
    auth_type = normalized_auth_type(auth_type)
    if auth_type in TOKEN_AUTH_TYPES:
        return "NSSESSIONID"
    return auth_type
