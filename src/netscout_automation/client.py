"""NETSCOUT Config API 호출용 HTTP client."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from urllib.parse import quote, urljoin, urlparse

import requests
import urllib3

from .auth import auth_log_label, is_token_auth, request_auth_kwargs
from .config import NetscoutSettings

logger = logging.getLogger(__name__)


class NetscoutClientError(RuntimeError):
    """NETSCOUT Config API 요청 실패 시 발생한다."""


@dataclass
class NetscoutClient:
    """Basic Auth 또는 NSSESSIONID cookie 인증을 포함한 Config API client."""

    settings: NetscoutSettings

    def __post_init__(self) -> None:
        if self.settings.auth_type.lower() != "basic" and not is_token_auth(self.settings.auth_type):
            raise NetscoutClientError(f"Unsupported auth_type for Config API: {self.settings.auth_type}")

    def get_applications_raw(self) -> str:
        """Application catalog 원문 응답을 가져온다."""
        return self._get_raw(self.settings.config_applications_endpoint, "Config API Application List")

    def get_devices(self) -> str:
        """Device List 원문 응답을 가져온다."""
        return self._get_raw(self.settings.devices_endpoint, "Config API Device List")

    def get_device_interfaces(self, device_name: str) -> str:
        """Device name을 URL encoding하여 해당 Device의 Interface List 원문 응답을 가져온다."""
        encoded_name = quote(device_name, safe="")
        endpoint = self.settings.device_interfaces_endpoint_template.format(device_name=encoded_name)
        return self._get_raw(endpoint, "Config API Device Interface List")

    def _get_raw(self, endpoint: str, action: str) -> str:
        url = self._url(endpoint)
        host = urlparse(url).hostname or "unknown-host"
        # 인증정보와 SNMP community string은 로그에 남기지 않고 host/endpoint/SSL 설정만 기록한다.
        logger.info(
            "Calling %s on host=%s endpoint=%s auth=%s verify_ssl=%s",
            action,
            host,
            endpoint,
            auth_log_label(self.settings.auth_type),
            self.settings.verify_ssl,
        )

        try:
            if not self.settings.verify_ssl:
                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            response = self._get(url)
        except requests.exceptions.Timeout as exc:
            raise NetscoutClientError(f"{action} request timed out. Check network and timeout settings.") from exc
        except requests.exceptions.SSLError as exc:
            raise NetscoutClientError(f"{action} failed with an SSL error. Check verify_ssl and server certificate.") from exc
        except requests.exceptions.ConnectionError as exc:
            raise NetscoutClientError(f"{action} connection failed. Check base_url and network access.") from exc
        except requests.exceptions.RequestException as exc:
            raise NetscoutClientError(f"{action} request failed.") from exc

        if response.status_code == 500:
            raise NetscoutClientError(f"{action} failed with HTTP 500 internal server error.")
        if response.status_code != 200:
            preview = sanitize_response_preview(response.text)
            raise NetscoutClientError(
                f"{action} failed with HTTP {response.status_code}. Check endpoint, protocol, and credentials. "
                f"Response preview='{preview}'"
            )

        if not response.text.strip():
            content_type = response.headers.get("Content-Type", "unknown")
            raise NetscoutClientError(f"{action} returned HTTP 200 with an empty body. Content-Type={content_type}.")
        return response.text

    def _get(self, url: str) -> requests.Response:
        """설정된 인증 방식이 포함된 Config API GET 요청을 보낸다."""
        return requests.get(
            url,
            headers={"Content-Type": "application/xml", "Accept": "application/json, application/xml"},
            verify=self.settings.verify_ssl,
            timeout=self.settings.timeout,
            **request_auth_kwargs(self.settings),
        )

    def _url(self, endpoint: str) -> str:
        return urljoin(self.settings.base_url.rstrip("/") + "/", endpoint.lstrip("/"))


def sanitize_response_preview(text: str, limit: int = 300) -> str:
    """HTTP error 응답 일부를 로그용으로 줄인다.

    인증정보가 들어갈 수 있는 긴 본문을 그대로 출력하지 않기 위해 한 줄 preview만 남긴다.
    """
    return " ".join(str(text or "").split())[:limit]

