from pathlib import Path

import pytest

import oracle_query_hostname as oracle


def test_build_and_parse_sqlplus_output() -> None:
    config = {
        "username": "NETSCOUT_LAB",
        "connect_identifier": "//localhost:1521/XE",
        "sql": "SELECT HOSTNAME FROM HOSTNAME_LOOKUP WHERE IP_ADDRESS = '{{IP}}' AND ROWNUM = 1",
    }
    script = oracle.build_sql_script(["192.168.0.10", "203.0.113.99"], config, "secret")
    assert "192.168.0.10" in script
    assert "CONNECT NETSCOUT_LAB/\"secret\"@//localhost:1521/XE" in script

    output = "\n".join(
        (
            "__NETSCOUT_BEGIN__192.168.0.10",
            "md-feed-gw-01",
            "__NETSCOUT_END__192.168.0.10",
            "__NETSCOUT_BEGIN__203.0.113.99",
            "__NETSCOUT_END__203.0.113.99",
        )
    )
    parsed = oracle.parse_sqlplus_output(output, ["192.168.0.10", "203.0.113.99"])
    assert parsed == {"192.168.0.10": "md-feed-gw-01", "203.0.113.99": ""}


def test_load_config_requires_ip_placeholder(tmp_path: Path) -> None:
    config = tmp_path / "oracle.yaml"
    config.write_text(
        "oracle:\n  username: TEST\n  connect_identifier: XE\n"
        "query:\n  sql: SELECT HOSTNAME FROM HOSTS\n",
        encoding="utf-8",
    )
    with pytest.raises(oracle.OracleHostnameError, match="placeholder"):
        oracle.load_config(config)


def test_invalid_ip_is_rejected() -> None:
    config = {
        "username": "TEST",
        "connect_identifier": "XE",
        "sql": "SELECT HOSTNAME FROM HOSTS WHERE IP = '{{IP}}'",
    }
    with pytest.raises(ValueError):
        oracle.build_sql_script(["1.2.3.4' OR '1'='1"], config, "secret")


def test_yaml_candidate_rows_preserve_port(tmp_path: Path) -> None:
    source = tmp_path / "candidates.yaml"
    source.write_text(
        "unique_ip_ports:\n"
        "  - ip: 192.168.0.10\n"
        "    port: 443\n"
        "    applications: [HTTPS]\n",
        encoding="utf-8",
    )
    assert oracle.load_yaml_candidate_rows(source) == [
        {"IP": "192.168.0.10", "Port": "443", "Applications": ["HTTPS"]}
    ]
