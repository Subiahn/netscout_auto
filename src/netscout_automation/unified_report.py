"""Application 중심 Unified XLSX report 생성 로직."""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from time import perf_counter
from typing import Any
from urllib.parse import urljoin, urlparse

import requests
import urllib3
from openpyxl import Workbook
from openpyxl.styles import Font
from requests.auth import HTTPBasicAuth

from .client import NetscoutClient, NetscoutClientError
from .config import NetscoutSettings
from .parsers import APPLICATION_COLUMNS, ParserError, parse_applications_response
from .tcp_other import SearchTarget, TcpOtherError, build_tcp_other_topnconv_query_xml, search_application

logger = logging.getLogger(__name__)

APPLICATION_SUMMARY_COLUMNS = [
    "Application",
    "App ID",
    "Total bps",
    "Total Byte",
    "Total Byte In",
    "Total Byte Out",
    "Total Packets",
    "Total Packets In",
    "Total Packets Out",
    "Flow Count",
    "Detail Link",
]

HOST_SUMMARY_COLUMNS = [
    "Application",
    "App ID",
    "Source IP",
    "Source Port",
    "Destination IP",
    "Destination Port",
    "Total bps",
    "Total Byte",
    "Total Byte In",
    "Total Byte Out",
    "Total Packets",
    "Total Packets In",
    "Total Packets Out",
    "Observed Interfaces",
    "Observed Interface Count",
]

RAW_DETAIL_COLUMNS = [
    "deviceName",
    "deviceIPAddress",
    "interfaceName",
    "interfaceNumber",
    "interface_key",
    "appId",
    "appId_String",
    "srcAddress",
    "srcPort",
    "destAddress",
    "destPort",
    "srcToDestOctets",
    "destToSrcOctets",
    "octets",
    "srcToDestPackets",
    "destToSrcPackets",
    "packets",
    "bps",
]


class UnifiedReportError(RuntimeError):
    """Unified report를 생성할 수 없을 때 발생한다."""


@dataclass(frozen=True)
class UnifiedReportResult:
    applications_count: int
    interface_count: int
    detail_count: int
    application_summary_count: int
    host_summary_count: int
    detail_success_count: int
    detail_failure_count: int
    output_path: Path


def build_interface_key(monitor_ip: object, interface_no: object) -> str:
    """Monitor IP와 Interface No. 조합으로 관측 지점 key를 만든다."""
    return f"{monitor_ip}|{interface_no}"


def resolve_unified_output_path(output: str | Path | None, output_directory: str, timestamp: bool = False) -> Path:
    """--output 값과 --timestamp 옵션을 반영해 Unified Report 저장 경로를 결정한다."""
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    if output in (None, ""):
        return Path(output_directory) / f"netscout_unified_report_{now}.xlsx"

    path = Path(output)
    if timestamp:
        suffix = path.suffix or ".xlsx"
        return path.with_name(f"{path.stem}_{now}{suffix}")
    return path


def fetch_topnconv_all_for_target(
    settings: NetscoutSettings,
    target: SearchTarget,
    duration: str,
) -> list[dict[str, object]]:
    """특정 관측 지점의 TOPNCONV 전체 Application row를 조회한다."""
    url = _url(settings.base_url, settings.data_api_endpoint)
    host = urlparse(url).hostname or "unknown-host"
    query_xml = build_tcp_other_topnconv_query_xml(target.ip_address, target.ifn, duration)
    logger.info(
        "Calling unified TOPNCONV host=%s endpoint=%s monitor_ip=%s ifn=%s verify_ssl=%s",
        host,
        settings.data_api_endpoint,
        target.ip_address,
        target.ifn,
        settings.verify_ssl,
    )

    try:
        if not settings.verify_ssl:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        start = perf_counter()
        response = requests.post(
            url,
            auth=HTTPBasicAuth(settings.username, settings.password),
            headers={"Content-Type": "application/xml", "Accept": "application/json"},
            data=query_xml.encode("utf-8"),
            verify=settings.verify_ssl,
            timeout=settings.timeout,
        )
        elapsed = perf_counter() - start
    except requests.exceptions.Timeout as exc:
        raise UnifiedReportError(f"Unified TOPNCONV timed out for ip={target.ip_address} ifn={target.ifn}.") from exc
    except requests.exceptions.SSLError as exc:
        raise UnifiedReportError(f"Unified TOPNCONV SSL error for ip={target.ip_address} ifn={target.ifn}.") from exc
    except requests.exceptions.ConnectionError as exc:
        raise UnifiedReportError(f"Unified TOPNCONV connection failed for ip={target.ip_address} ifn={target.ifn}.") from exc
    except requests.exceptions.RequestException as exc:
        raise UnifiedReportError(f"Unified TOPNCONV request failed for ip={target.ip_address} ifn={target.ifn}.") from exc

    logger.info(
        "Unified TOPNCONV response status=%s elapsed=%.3fs monitor_ip=%s ifn=%s",
        response.status_code,
        elapsed,
        target.ip_address,
        target.ifn,
    )
    if response.status_code != 200:
        raise UnifiedReportError(
            f"Unified TOPNCONV failed with HTTP {response.status_code}. Response preview={_sanitized_preview(response.text)!r}"
        )

    try:
        payload = response.json()
    except ValueError as exc:
        raise UnifiedReportError(f"Unified TOPNCONV returned invalid JSON for ip={target.ip_address} ifn={target.ifn}.") from exc

    rows = payload.get("response", [])
    if not isinstance(rows, list):
        raise UnifiedReportError("Unified TOPNCONV JSON response field is not a list.")
    return [row for row in rows if isinstance(row, dict)]


def normalize_raw_detail_rows(rows: list[dict[str, object]], target: SearchTarget, duration: str) -> list[dict[str, object]]:
    """TOPNCONV row에 관측 지점과 bps 계산값을 붙여 Raw_Detail row로 정규화한다."""
    duration_seconds = _duration_to_seconds(duration)
    normalized = []
    for row in rows:
        normalized.append(
            {
                "deviceName": None,
                "deviceIPAddress": target.ip_address,
                "interfaceName": None,
                "interfaceNumber": target.ifn,
                "interface_key": build_interface_key(target.ip_address, target.ifn),
                "appId": row.get("appId"),
                "appId_String": row.get("appId_String"),
                "srcAddress": row.get("srcAddress"),
                "srcPort": row.get("srcPort"),
                "destAddress": row.get("destAddress"),
                "destPort": row.get("destPort"),
                "srcToDestOctets": row.get("srcToDestOctets"),
                "destToSrcOctets": row.get("destToSrcOctets"),
                "octets": row.get("octets"),
                "srcToDestPackets": row.get("srcToDestPackets"),
                "destToSrcPackets": row.get("destToSrcPackets"),
                "packets": row.get("packets"),
                "bps": _calculate_bps(row.get("octets"), duration_seconds),
            }
        )
    return normalized


def build_application_summary(raw_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    """Application/App ID 기준으로 전체 사용량을 집계한다."""
    grouped: dict[tuple[object, object], dict[str, object]] = {}
    for row in raw_rows:
        key = (row.get("appId_String"), row.get("appId"))
        if key not in grouped:
            grouped[key] = {
                "Application": row.get("appId_String"),
                "App ID": row.get("appId"),
                "Total bps": 0,
                "Total Byte": 0,
                "Total Byte In": 0,
                "Total Byte Out": 0,
                "Total Packets": 0,
                "Total Packets In": 0,
                "Total Packets Out": 0,
                "Flow Count": 0,
                "Detail Link": "Go to Host_Summary",
            }
        summary = grouped[key]
        summary["Total bps"] = _sum_number(summary["Total bps"], row.get("bps"))
        summary["Total Byte"] = _sum_number(summary["Total Byte"], row.get("octets"))
        summary["Total Byte In"] = _sum_number(summary["Total Byte In"], row.get("srcToDestOctets"))
        summary["Total Byte Out"] = _sum_number(summary["Total Byte Out"], row.get("destToSrcOctets"))
        summary["Total Packets"] = _sum_number(summary["Total Packets"], row.get("packets"))
        summary["Total Packets In"] = _sum_number(summary["Total Packets In"], row.get("srcToDestPackets"))
        summary["Total Packets Out"] = _sum_number(summary["Total Packets Out"], row.get("destToSrcPackets"))
        summary["Flow Count"] = int(summary["Flow Count"]) + 1
    return sorted(grouped.values(), key=lambda item: (_numeric(item.get("Total bps")), _numeric(item.get("Total Byte"))), reverse=True)


def build_host_summary(raw_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    """Application 내부의 Source/Destination Host 통신을 요약한다."""
    grouped: dict[tuple[object, object, object, object, object, object], dict[str, object]] = {}
    observed: dict[tuple[object, object, object, object, object, object], set[str]] = {}
    for row in raw_rows:
        key = (
            row.get("appId_String"),
            row.get("appId"),
            row.get("srcAddress"),
            row.get("srcPort"),
            row.get("destAddress"),
            row.get("destPort"),
        )
        if key not in grouped:
            grouped[key] = {
                "Application": row.get("appId_String"),
                "App ID": row.get("appId"),
                "Source IP": row.get("srcAddress"),
                "Source Port": row.get("srcPort"),
                "Destination IP": row.get("destAddress"),
                "Destination Port": row.get("destPort"),
                "Total bps": 0,
                "Total Byte": 0,
                "Total Byte In": 0,
                "Total Byte Out": 0,
                "Total Packets": 0,
                "Total Packets In": 0,
                "Total Packets Out": 0,
                "Observed Interfaces": "",
                "Observed Interface Count": 0,
            }
            observed[key] = set()
        summary = grouped[key]
        summary["Total bps"] = _sum_number(summary["Total bps"], row.get("bps"))
        summary["Total Byte"] = _sum_number(summary["Total Byte"], row.get("octets"))
        summary["Total Byte In"] = _sum_number(summary["Total Byte In"], row.get("srcToDestOctets"))
        summary["Total Byte Out"] = _sum_number(summary["Total Byte Out"], row.get("destToSrcOctets"))
        summary["Total Packets"] = _sum_number(summary["Total Packets"], row.get("packets"))
        summary["Total Packets In"] = _sum_number(summary["Total Packets In"], row.get("srcToDestPackets"))
        summary["Total Packets Out"] = _sum_number(summary["Total Packets Out"], row.get("destToSrcPackets"))
        if row.get("interface_key") not in (None, ""):
            observed[key].add(str(row.get("interface_key")))

    summaries = []
    for key, summary in grouped.items():
        interfaces = sorted(observed[key])
        summary["Observed Interfaces"] = ", ".join(interfaces)
        summary["Observed Interface Count"] = len(interfaces)
        summaries.append(summary)
    return sorted(summaries, key=lambda item: (str(item.get("Application") or ""), -_numeric(item.get("Total bps")), -_numeric(item.get("Total Byte"))))


def export_unified_report_xlsx(
    application_summary: list[dict[str, object]],
    host_summary: list[dict[str, object]],
    raw_detail: list[dict[str, object]],
    applications: list[dict[str, Any]],
    output_path: str | Path,
) -> Path:
    """Application summary 중심의 Unified workbook을 저장한다."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    _write_sheet(workbook.active, "Application_Summary", APPLICATION_SUMMARY_COLUMNS, application_summary)
    _write_sheet(workbook.create_sheet("Host_Summary"), "Host_Summary", HOST_SUMMARY_COLUMNS, host_summary)
    _write_sheet(workbook.create_sheet("Raw_Detail"), "Raw_Detail", RAW_DETAIL_COLUMNS, raw_detail)
    _write_sheet(workbook.create_sheet("Application_List"), "Application_List", APPLICATION_COLUMNS, applications)
    _add_summary_links(workbook["Application_Summary"])

    workbook.save(path)
    return path


def generate_unified_report(
    settings: NetscoutSettings,
    application: str,
    duration: str,
    output_path: str | Path,
) -> UnifiedReportResult:
    """Search API로 관측 지점을 찾고 TOPNCONV 전체 Application을 Application/Host 기준으로 집계한다."""
    try:
        applications = parse_applications_response(NetscoutClient(settings).get_applications_raw())
    except (NetscoutClientError, ParserError) as exc:
        raise UnifiedReportError("Failed to collect Config API Application List.") from exc

    try:
        targets = search_application(settings, application)
    except TcpOtherError as exc:
        raise UnifiedReportError("Failed to discover monitor interfaces with Search API.") from exc

    raw_detail: list[dict[str, object]] = []
    success_count = 0
    failure_count = 0
    for target in targets:
        try:
            raw_rows = fetch_topnconv_all_for_target(settings, target, duration)
            raw_detail.extend(normalize_raw_detail_rows(raw_rows, target, duration))
            success_count += 1
        except UnifiedReportError as exc:
            failure_count += 1
            logger.warning("Skipping failed unified TOPNCONV target ip=%s ifn=%s: %s", target.ip_address, target.ifn, exc)

    application_summary = build_application_summary(raw_detail)
    host_summary = build_host_summary(raw_detail)
    path = export_unified_report_xlsx(application_summary, host_summary, raw_detail, applications, output_path)

    logger.info(
        "Unified report completed applications=%s interfaces=%s raw=%s application_summary=%s host_summary=%s succeeded=%s failed=%s output=%s",
        len(applications),
        len(targets),
        len(raw_detail),
        len(application_summary),
        len(host_summary),
        success_count,
        failure_count,
        path,
    )
    return UnifiedReportResult(
        applications_count=len(applications),
        interface_count=len(targets),
        detail_count=len(raw_detail),
        application_summary_count=len(application_summary),
        host_summary_count=len(host_summary),
        detail_success_count=success_count,
        detail_failure_count=failure_count,
        output_path=path,
    )


def _write_sheet(worksheet, title: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    worksheet.title = title
    worksheet.append(columns)
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        worksheet.append([row.get(column) for column in columns])
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column_cells in worksheet.columns:
        width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 12), 60)


def _add_summary_links(worksheet) -> None:
    if worksheet.max_row < 2:
        return
    detail_link_column = APPLICATION_SUMMARY_COLUMNS.index("Detail Link") + 1
    for row in range(2, worksheet.max_row + 1):
        cell = worksheet.cell(row=row, column=detail_link_column)
        cell.hyperlink = "#'Host_Summary'!A1"
        cell.style = "Hyperlink"


def _duration_to_seconds(duration: str) -> int | None:
    return {
        "LAST_HOUR": 3600,
        "LAST_6_HOURS": 21600,
        "LAST_12_HOURS": 43200,
        "LAST_24_HOURS": 86400,
        "LAST_DAY": 86400,
        "LAST_7_DAYS": 604800,
        "LAST_WEEK": 604800,
    }.get(duration.upper())


def _calculate_bps(octets: object, duration_seconds: int | None) -> float | None:
    if not duration_seconds:
        return None
    try:
        return float(octets) * 8 / duration_seconds
    except (TypeError, ValueError):
        return None


def _sum_number(left: object, right: object) -> float:
    return _numeric(left) + _numeric(right)


def _numeric(value: object) -> float:
    if value in (None, ""):
        return 0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0


def _url(base_url: str, endpoint: str) -> str:
    return urljoin(base_url.rstrip("/") + "/", endpoint.lstrip("/"))


def _sanitized_preview(text: str, limit: int = 300) -> str:
    compact = " ".join((text or "").split())
    try:
        parsed = json.loads(compact)
    except json.JSONDecodeError:
        return compact[:limit]
    if isinstance(parsed, dict):
        for key in list(parsed):
            if key.lower() in {"password", "token", "cookie", "authorization", "apikey", "api_key"}:
                parsed[key] = "***"
    return json.dumps(parsed, ensure_ascii=True)[:limit]

