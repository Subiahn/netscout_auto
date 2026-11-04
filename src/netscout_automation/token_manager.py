"""NSSESSIONID token 발급과 로컬 env 파일 갱신 helper."""

from __future__ import annotations

import logging
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
import urllib3
import yaml
from requests.auth import HTTPBasicAuth

from .config import ConfigError, NetscoutSettings, load_dotenv_file, load_settings

logger = logging.getLogger(__name__)

TOKEN_ENV_KEY = "NETSCOUT_NSSESSIONID"


class TokenRefreshError(RuntimeError):
    """NSSESSIONID 발급 또는 저장에 실패했을 때 발생한다."""


def load_settings_with_auto_token(
    config_path: str | Path,
    env_path: str | Path,
    refresh_token: bool = False,
) -> NetscoutSettings:
    """V9 token 인증용 설정을 로드하고 필요하면 NSSESSIONID 파일을 갱신한다.

    token auth 설정에서 token 값이 없거나 `refresh_token=True`이면 Basic 계정으로
    NSSESSIONID를 다시 발급받아 env 파일에 저장한 뒤 설정을 다시 로드한다.
    """
    try:
        settings = load_settings(config_path, env_path)
    except ConfigError as exc:
        if "NETSCOUT_NSSESSIONID" not in str(exc):
            raise
        settings = load_refresh_settings(config_path, env_path)
        token_path = token_file_path(settings, env_path)
        token = issue_nssessionid(settings)
        update_token_env_file(token_path, token)
        logger.info("NSSESSIONID token file updated path=%s", token_path)
        return load_settings(config_path, token_path)

    if refresh_token and settings.auth_type in {"nssessionid", "session_cookie", "cookie", "token"}:
        refresh_settings = load_refresh_settings(config_path, env_path)
        token_path = token_file_path(refresh_settings, env_path)
        token = issue_nssessionid(refresh_settings)
        update_token_env_file(token_path, token)
        logger.info("NSSESSIONID token file refreshed path=%s", token_path)
        return load_settings(config_path, token_path)

    return settings


def load_refresh_settings(config_path: str | Path, env_path: str | Path) -> NetscoutSettings:
    """token 발급용 Basic Auth 설정을 만든다.

    원래 config가 `auth_type: nssessionid`여도 발급 요청에는 username/password가
    필요하므로, 같은 YAML을 Basic Auth로 임시 해석한다.
    """
    load_dotenv_file(env_path)
    path = Path(config_path)
    if not path.exists():
        raise ConfigError(f"Config file not found: {path}")
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    data.setdefault("netscout", {})["auth_type"] = "basic"
    temp_config = path.with_suffix(path.suffix + ".token-refresh.tmp")
    try:
        temp_config.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")
        return load_settings(temp_config, env_path)
    finally:
        try:
            temp_config.unlink()
        except OSError:
            pass


def issue_nssessionid(settings: NetscoutSettings) -> str:
    """설정된 token 발급 API를 Basic Auth로 호출하고 NSSESSIONID 값을 가져온다."""
    endpoint = settings.session_token_endpoint
    if not endpoint:
        raise TokenRefreshError(
            "session_token_endpoint is required for automatic NSSESSIONID issue. "
            "Check the NETSCOUT token issue API and set api.session_token_endpoint."
        )
    url = urljoin(settings.base_url.rstrip("/") + "/", endpoint.lstrip("/"))
    host = urlparse(url).hostname or "unknown-host"
    method = str(settings.session_token_method or "GET").upper()
    if method not in {"GET", "POST"}:
        raise TokenRefreshError(f"Unsupported session_token_method: {method}")
    logger.info(
        "Requesting NSSESSIONID token host=%s endpoint=%s method=%s verify_ssl=%s",
        host,
        endpoint,
        method,
        settings.verify_ssl,
    )
    if not settings.verify_ssl:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    try:
        response = requests.request(
            method,
            url,
            headers=token_request_headers(settings),
            auth=HTTPBasicAuth(settings.username, settings.password),
            data=settings.session_token_body.encode("utf-8") if settings.session_token_body else None,
            verify=settings.verify_ssl,
            timeout=settings.timeout,
        )
    except requests.exceptions.Timeout as exc:
        raise TokenRefreshError("NSSESSIONID token request timed out.") from exc
    except requests.exceptions.SSLError as exc:
        raise TokenRefreshError("NSSESSIONID token request failed with SSL error.") from exc
    except requests.exceptions.ConnectionError as exc:
        raise TokenRefreshError("NSSESSIONID token request connection failed.") from exc
    except requests.exceptions.RequestException as exc:
        raise TokenRefreshError("NSSESSIONID token request failed.") from exc

    if response.status_code in {401, 403}:
        raise TokenRefreshError(f"NSSESSIONID token request failed with HTTP {response.status_code}. Check credentials.")
    if response.status_code >= 400:
        raise TokenRefreshError(f"NSSESSIONID token request failed with HTTP {response.status_code}.")

    cookie_name = settings.session_token_cookie_name or "NSSESSIONID"
    token = response.cookies.get(cookie_name)
    if not token:
        token = extract_cookie_from_header(response.headers.get("Set-Cookie", ""), cookie_name)
    if not token:
        raise TokenRefreshError(f"{cookie_name} was not found in response cookies. Check session_token_endpoint.")
    return token


def token_request_headers(settings: NetscoutSettings) -> dict[str, str]:
    """token 발급 요청에 사용할 header를 만든다."""
    content_type = settings.session_token_content_type or "application/xml"
    return {"Content-Type": content_type, "Accept": "application/json, application/xml, */*"}


def extract_cookie_from_header(header: str, cookie_name: str) -> str:
    """Set-Cookie header에서 지정한 cookie 값을 추출한다."""
    prefix = f"{cookie_name}="
    for part in header.split(";"):
        text = part.strip()
        if text.startswith(prefix):
            return text[len(prefix) :]
    return ""


def token_file_path(settings: NetscoutSettings, env_path: str | Path) -> Path:
    """token 저장 파일 경로를 결정한다."""
    if settings.session_token_file:
        return Path(settings.session_token_file)
    return Path(env_path)


def update_token_env_file(path: str | Path, token: str) -> None:
    """env 파일의 NETSCOUT_NSSESSIONID 값을 갱신한다.

    token 값은 반환하거나 로그에 출력하지 않는다.
    """
    env_path = Path(path)
    env_path.parent.mkdir(parents=True, exist_ok=True)
    lines = env_path.read_text(encoding="utf-8").splitlines() if env_path.exists() else []
    updated = False
    new_lines = []
    for line in lines:
        if line.strip().startswith(f"{TOKEN_ENV_KEY}="):
            new_lines.append(f"{TOKEN_ENV_KEY}={token}")
            updated = True
        else:
            new_lines.append(line)
    if not updated:
        if new_lines and new_lines[-1].strip():
            new_lines.append("")
        new_lines.append(f"{TOKEN_ENV_KEY}={token}")
    env_path.write_text("\n".join(new_lines) + "\n", encoding="utf-8")
