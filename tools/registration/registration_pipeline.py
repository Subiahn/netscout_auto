"""v0.9.1 multi-device registration preparation pipeline.

This tool never changes a NETSCOUT appliance. It creates an editable collection
scope, consolidates Detail evidence, and writes a registration review workbook.
Use netscout_sync.py separately for plan/apply/verify.
"""
from __future__ import annotations

import argparse
import fnmatch
import importlib.util
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
REGISTRATION_DIR = Path(__file__).resolve().parent
for path in (ROOT, ROOT / "vendor", REGISTRATION_DIR):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from registration_rules import HEADERS, text


REVIEW_EXTRA = [
    "SourceApplication", "DeviceCount", "InterfaceCount", "EvidenceCount",
    "Devices", "Interfaces", "Review", "PolicyRule",
]


def _rows(ws) -> list[dict[str, Any]]:
    values = list(ws.iter_rows(values_only=True))
    if not values:
        return []
    headers = [text(value) for value in values[0]]
    return [dict(zip(headers, row)) for row in values[1:] if any(value not in (None, "") for value in row)]


def generate_scope(inventory: Path, output: Path, duration: str = "LAST_24_HOURS", top_n: int = 50) -> dict[str, int]:
    """Create editable device/interface scope YAML from an Inventory workbook."""
    workbook = load_workbook(inventory, read_only=True, data_only=True)
    try:
        sheet_name = "Interface_List" if "Interface_List" in workbook.sheetnames else "Interfaces"
        if sheet_name not in workbook.sheetnames:
            raise ValueError("Inventory workbook requires Interface_List or Interfaces sheet")
        source = _rows(workbook[sheet_name])
    finally:
        workbook.close()
    targets = []
    seen = set()
    for row in source:
        device = text(row.get("deviceName") or row.get("Device Name"))
        interface = text(row.get("interfaceName") or row.get("Interface Name"))
        number = text(row.get("interfaceNumber") or row.get("Interface Number"))
        if not device or not (interface or number):
            continue
        key = (device.casefold(), f"{interface.casefold()}|{number.casefold()}")
        if key in seen:
            continue
        seen.add(key)
        targets.append({
            "enabled": True,
            "device": device,
            "device_ip": text(row.get("deviceIPAddress") or row.get("Device IP")),
            "interface": interface,
            "interface_number": number,
            "interface_alias": text(row.get("alias") or row.get("Interface Alias")),
        })
    payload = {
        "description": "DUMMY example structure. Review enabled before collection.",
        "defaults": {"duration": duration, "top_n": top_n, "run_detail": False, "detail_top_n": 20},
        "targets": targets,
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(yaml.safe_dump(payload, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return {"targets": len(targets), "devices": len({row[0] for row in seen})}


def _extract_module():
    path = ROOT / "tools" / "extract_ip_port_candidates.py"
    spec = importlib.util.spec_from_file_location("netscout_v091_extract", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def consolidate(detail_root: Path, output: Path, yaml_output: Path | None = None) -> dict[str, Any]:
    """Consolidate all device/interface Detail workbooks into one evidence set."""
    module = _extract_module()
    return_path, yaml_path, stats = module.extract_ip_port_report(
        detail_root, output, yaml_output=yaml_output
    )
    return {"output": str(return_path), "yaml_output": str(yaml_path or ""), **stats}


def build_detail_scope(top_root: Path, output: Path, detail_top_n: int = 20, limit: int = 1000) -> dict[str, int]:
    """Merge every Top-N workbook into one run-detail-targets compatible YAML."""
    targets: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str]] = set()
    files = sorted(path for path in top_root.rglob("*.xlsx") if not path.name.startswith("~$"))
    for path in files:
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            if "Application_BPS_TopN" not in book.sheetnames:
                continue
            rows = _rows(book["Application_BPS_TopN"])
        finally:
            book.close()
        rows.sort(key=lambda row: float(row.get("Rank") or 999999))
        if detail_top_n > 0:
            rows = rows[:detail_top_n]
        for row in rows:
            application = text(row.get("Application Type Code") or row.get("Application Name"))
            device = text(row.get("Device Name"))
            interface = text(row.get("Interface Name") or row.get("Interface Number"))
            if not application or not device or not interface:
                continue
            key = (device.casefold(), interface.casefold(), application.casefold())
            if key in seen:
                continue
            seen.add(key)
            targets.append({
                "enabled": True,
                "rank": row.get("Rank"),
                "application": application,
                "app_id": row.get("App ID"),
                "bps": row.get("Bit Rate") or row.get("BPS"),
                "device": device,
                "device_ip": row.get("Device IP"),
                "interface": row.get("Interface Name"),
                "interface_number": row.get("Interface Number"),
                "duration": row.get("Duration") or "LAST_24_HOURS",
                "limit": limit,
                "dedup_ip": False,
                "source_top_xlsx": str(path),
            })
    payload = {
        "description": "Generated v0.9.1 multi-device Detail scope. Review enabled before execution.",
        "defaults": {"duration": "LAST_24_HOURS", "limit": limit, "dedup_ip": False},
        "targets": targets,
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(yaml.safe_dump(payload, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return {"top_files": len(files), "targets": len(targets)}


def _load_policy(path: Path) -> dict[str, Any]:
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ValueError("Policy root must be a mapping")
    return data


def _rule_for(application: str, policy: dict[str, Any]) -> tuple[dict[str, Any], str]:
    for index, rule in enumerate(policy.get("application_rules", []), 1):
        pattern = text(rule.get("match"))
        if pattern and fnmatch.fnmatchcase(application.upper(), pattern.upper()):
            return {**policy.get("defaults", {}), **rule}, f"rule#{index}:{pattern}"
    return dict(policy.get("defaults", {})), "defaults"


def build_review(candidates: Path, policy_path: Path, output: Path) -> dict[str, int]:
    """Create a netscout_sync-compatible review workbook; Register defaults to N."""
    if output.exists():
        raise ValueError("Output already exists; preserve manual review by using a new path")
    policy = _load_policy(policy_path)
    workbook = load_workbook(candidates, read_only=True, data_only=True)
    try:
        if "IP_Port_Candidates" not in workbook.sheetnames:
            raise ValueError("Candidate workbook requires IP_Port_Candidates sheet")
        rows = _rows(workbook["IP_Port_Candidates"])
    finally:
        workbook.close()

    grouped: dict[tuple[str, str, int], dict[str, Any]] = {}
    for row in rows:
        application, ip = text(row.get("Application")), text(row.get("IP"))
        try:
            port = int(row.get("Port"))
        except (TypeError, ValueError):
            continue
        if not application or not ip or not 1 <= port <= 65535:
            continue
        rule, rule_name = _rule_for(application, policy)
        name = text(rule.get("application_name"))
        if name:
            name = name.format(application=application, port=port)
        key = (application, name, port)
        item = grouped.setdefault(key, {
            "application": application, "name": name, "port": port, "ips": set(),
            "devices": set(), "interfaces": set(), "evidence": 0,
            "rule": rule, "rule_name": rule_name,
        })
        item["ips"].add(ip)
        item["devices"].update(part.strip() for part in text(row.get("Devices")).split(",") if part.strip())
        item["interfaces"].update(part.strip() for part in text(row.get("Interfaces")).split(",") if part.strip())
        item["evidence"] += int(row.get("Row Count") or 0)

    result = Workbook()
    ws = result.active
    ws.title = "Registration"
    columns = HEADERS + REVIEW_EXTRA
    ws.append(columns)
    needs_review = 0
    for item in sorted(grouped.values(), key=lambda value: (value["application"], value["port"])):
        rule = item["rule"]
        registration_type = text(rule.get("registration_type")).upper()
        parent = text(rule.get("parent_application")).upper()
        protocol = text(rule.get("protocol")).upper()
        reasons = []
        if not item["name"]:
            reasons.append("ApplicationName policy mapping required")
        if registration_type not in {"SBA", "WKA", "CHILD", "URL"}:
            reasons.append("RegistrationType review required")
        if not parent or protocol not in {"TCP", "UDP", "SCTP"}:
            reasons.append("ParentApplication/Protocol review required")
        if reasons:
            needs_review += 1
        values = {
            "Register": "N", "ApplicationName": item["name"],
            "RegistrationType": registration_type, "ParentApplication": parent,
            "Protocol": protocol, "IP": ",".join(sorted(item["ips"])), "Port": item["port"],
            "Description": text(rule.get("description")),
            "ApplicationGroup": text(rule.get("application_group")),
            "ApplicationTag": text(rule.get("application_tag")),
            "SourceApplication": item["application"], "DeviceCount": len(item["devices"]),
            "InterfaceCount": len(item["interfaces"]), "EvidenceCount": item["evidence"],
            "Devices": ", ".join(sorted(item["devices"])),
            "Interfaces": ", ".join(sorted(item["interfaces"])),
            "Review": "; ".join(reasons), "PolicyRule": item["rule_name"],
        }
        ws.append([values.get(column, "") for column in columns])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="16324F")
    for column, formula in (("A", '"Y,N"'), ("C", '"SBA,WKA,CHILD,URL"'), ("E", '"TCP,UDP,SCTP"')):
        validation = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(validation)
        validation.add(f"{column}2:{column}{max(2, ws.max_row)}")
    for cells in ws.columns:
        ws.column_dimensions[cells[0].column_letter].width = min(max(max(len(text(c.value)) for c in cells) + 2, 12), 60)
    output.parent.mkdir(parents=True, exist_ok=True)
    result.save(output)
    result.close()
    return {"rows": len(grouped), "needs_review": needs_review}


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__)
    commands = root.add_subparsers(dest="command", required=True)
    scope = commands.add_parser("scope-generate")
    scope.add_argument("--inventory", required=True, type=Path)
    scope.add_argument("--output", default=Path("config/collection_scope.yaml"), type=Path)
    scope.add_argument("--duration", default="LAST_24_HOURS")
    scope.add_argument("--top-n", type=int, default=50)
    candidates = commands.add_parser("candidates-build")
    candidates.add_argument("--detail-root", required=True, type=Path)
    candidates.add_argument("--output", required=True, type=Path)
    candidates.add_argument("--yaml-output", type=Path)
    detail = commands.add_parser("detail-scope-build")
    detail.add_argument("--top-root", required=True, type=Path)
    detail.add_argument("--output", default=Path("config/detail_scope.yaml"), type=Path)
    detail.add_argument("--detail-top-n", default=20, type=int)
    detail.add_argument("--limit", default=1000, type=int)
    review = commands.add_parser("review-build")
    review.add_argument("--candidates", required=True, type=Path)
    review.add_argument("--policy", required=True, type=Path)
    review.add_argument("--output", required=True, type=Path)
    return root


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    if args.command == "scope-generate":
        print(generate_scope(args.inventory, args.output, args.duration, args.top_n))
    elif args.command == "detail-scope-build":
        print(build_detail_scope(args.top_root, args.output, args.detail_top_n, args.limit))
    elif args.command == "candidates-build":
        print(consolidate(args.detail_root, args.output, args.yaml_output))
    else:
        print(build_review(args.candidates, args.policy, args.output))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
