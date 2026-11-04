import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
REGISTRATION = ROOT / "tools" / "registration"
for path in (ROOT, ROOT / "vendor", REGISTRATION):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

import yaml
from openpyxl import Workbook, load_workbook

from registration_pipeline import build_detail_scope, build_review, generate_scope
from registration_rules import read_excel


def test_generate_scope_keeps_same_name_different_interface_numbers(tmp_path):
    source = tmp_path / "inventory.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Interface_List"
    sheet.append(["deviceName", "deviceIPAddress", "interfaceName", "interfaceNumber", "alias"])
    sheet.append(["DEVICE_A", "192.0.2.10", "eth", "1", "A"])
    sheet.append(["DEVICE_A", "192.0.2.10", "eth", "2", "B"])
    book.save(source)
    output = tmp_path / "scope.yaml"

    stats = generate_scope(source, output)

    data = yaml.safe_load(output.read_text(encoding="utf-8"))
    assert stats == {"targets": 2, "devices": 1}
    assert [row["interface_number"] for row in data["targets"]] == ["1", "2"]


def _candidate_book(path: Path):
    book = Workbook()
    sheet = book.active
    sheet.title = "IP_Port_Candidates"
    sheet.append(["Application", "IP", "Port", "Row Count", "Devices", "Interfaces"])
    sheet.append(["HTTPS", "192.0.2.20", 8443, 5, "DEVICE_A, DEVICE_B", "eth1|1, eth2|2"])
    sheet.append(["HTTPS", "192.0.2.21", 8443, 3, "DEVICE_A", "eth1|1"])
    sheet.append(["TCP_OTHER", "192.0.2.30", 9000, 2, "DEVICE_C", "eth3|3"])
    book.save(path)


def test_build_review_aggregates_evidence_and_defaults_register_to_n(tmp_path):
    candidates = tmp_path / "candidates.xlsx"
    _candidate_book(candidates)
    policy = tmp_path / "policy.yaml"
    policy.write_text(
        "defaults:\n  registration_type: SBA\n  parent_application: TCP\n  protocol: TCP\n  application_name: ''\n"
        "application_rules:\n  - match: HTTPS\n    application_name: WEB_{port}\n",
        encoding="utf-8",
    )
    output = tmp_path / "review.xlsx"

    stats = build_review(candidates, policy, output)

    book = load_workbook(output, data_only=True)
    sheet = book["Registration"]
    headers = [cell.value for cell in sheet[1]]
    rows = [dict(zip(headers, values)) for values in sheet.iter_rows(min_row=2, values_only=True)]
    book.close()
    https = next(row for row in rows if row["SourceApplication"] == "HTTPS")
    other = next(row for row in rows if row["SourceApplication"] == "TCP_OTHER")
    assert stats == {"rows": 2, "needs_review": 1}
    assert https["Register"] == "N"
    assert https["ApplicationName"] == "WEB_8443"
    assert set(https["IP"].split(",")) == {"192.0.2.20", "192.0.2.21"}
    assert https["DeviceCount"] == 2
    assert https["InterfaceCount"] == 2
    assert https["EvidenceCount"] == 8
    assert "ApplicationName" in other["Review"]

    editable = load_workbook(output)
    editable["Registration"]["A2"] = "Y"
    editable.save(output)
    editable.close()
    groups = read_excel(output)
    assert groups[0]["name"] == "WEB_8443"
    assert groups[0]["fields"]["ServerAddresses"] in {
        "192.0.2.20/32,192.0.2.21/32", "192.0.2.21/32,192.0.2.20/32"
    }


def test_build_review_refuses_overwrite(tmp_path):
    candidates = tmp_path / "candidates.xlsx"
    _candidate_book(candidates)
    policy = tmp_path / "policy.yaml"
    policy.write_text("defaults: {}\n", encoding="utf-8")
    output = tmp_path / "review.xlsx"
    output.write_text("keep", encoding="utf-8")
    try:
        build_review(candidates, policy, output)
    except ValueError as exc:
        assert "already exists" in str(exc)
    else:
        raise AssertionError("overwrite must be rejected")


def test_build_detail_scope_merges_multiple_top_workbooks(tmp_path):
    top_root = tmp_path / "01_top"
    for folder, device, interface in (("a", "DEVICE_A", "eth1"), ("b", "DEVICE_B", "eth2")):
        target = top_root / folder / "top.xlsx"
        target.parent.mkdir(parents=True, exist_ok=True)
        book = Workbook()
        sheet = book.active
        sheet.title = "Application_BPS_TopN"
        sheet.append(["Rank", "Application Name", "Application Type Code", "Device Name", "Interface Name", "Bit Rate"])
        sheet.append([1, "HTTPS", "HTTPS", device, interface, 1000])
        sheet.append([2, "HTTP", "HTTP", device, interface, 500])
        book.save(target)
    output = tmp_path / "detail_scope.yaml"

    stats = build_detail_scope(top_root, output, detail_top_n=1)

    data = yaml.safe_load(output.read_text(encoding="utf-8"))
    assert stats == {"top_files": 2, "targets": 2}
    assert {(row["device"], row["application"]) for row in data["targets"]} == {
        ("DEVICE_A", "HTTPS"), ("DEVICE_B", "HTTPS")
    }
