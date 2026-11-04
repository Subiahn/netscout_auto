# 더미 데이터 적용 가이드

이 저장소의 주소, 사용자, 비밀번호, 장비, Application, Oracle 정보는 모두 사용 예시를 위한 더미 값이다. 실제 운영값은 Git에 커밋하지 않는다.

## 더미 값 규칙

| 항목 | 저장소 예시 | 실제 적용 위치 |
|---|---|---|
| NETSCOUT URL | `https://netscout.example.com:8443` | `config/config.test.yaml` |
| NETSCOUT 사용자 | `DUMMY_ADMIN` | `config/.env.local` |
| NETSCOUT 비밀번호 | `CHANGE_ME` | `config/.env.local` |
| Device | `DUMMY_DEVICE` | CLI 인수 또는 대상 YAML |
| Interface | `DUMMY_INTERFACE` | CLI 인수 또는 대상 YAML |
| Monitor IP | `192.0.2.10` | 테스트 설정 |
| Oracle TNS Alias | `HOSTNAME_DB` | `config/oracle_hostname.yaml` |
| Oracle 사용자 | `NETSCOUT_READONLY` | `config/oracle_hostname.yaml` |
| Oracle 비밀번호 | `CHANGE_ME` | `config/oracle_hostname.password` |
| Oracle HOST | `oracle-host.example.com` | `config/oracle_network/tnsnames.ora` |
| Oracle Service | `ORCL` | `config/oracle_network/tnsnames.ora` |

`192.0.2.0/24`, `198.51.100.0/24`, `203.0.113.0/24`는 문서용 TEST-NET 주소다.

## 초기 설정

```powershell
Copy-Item .\.env.local.example .\config\.env.local
Copy-Item .\config\config.example.yaml .\config\config.test.yaml
Copy-Item .\config\hostname_lookup.example.yaml .\config\hostname_lookup.yaml
Copy-Item .\config\oracle_hostname.19c.example.yaml .\config\oracle_hostname.yaml
Copy-Item .\config\oracle_hostname.password.example .\config\oracle_hostname.password
New-Item -ItemType Directory -Force .\config\oracle_network | Out-Null
Copy-Item .\config\tnsnames.example.ora .\config\oracle_network\tnsnames.ora
```

생성된 실제 파일은 `.gitignore` 대상이다.

## NETSCOUT 예시

`config/config.test.yaml`:

```yaml
netscout:
  base_url: "https://netscout.example.com:8443"
  auth_type: "basic"
  verify_ssl: true
  timeout: 30
```

`config/.env.local`:

```text
NETSCOUT_USERNAME=DUMMY_ADMIN
NETSCOUT_PASSWORD=CHANGE_ME
```

실행 예:

```powershell
python -m netscout inventory `
  --config .\config\config.test.yaml `
  --env .\config\.env.local `
  --output-dir .\output
```

## Oracle TNS 예시

`config/oracle_network/tnsnames.ora`:

```ora
HOSTNAME_DB =
  (DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = oracle-host.example.com)(PORT = 1521))
    (CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = ORCL))
  )
```

`config/oracle_hostname.password`에는 비밀번호 한 줄만 입력한다. `KEY=VALUE` 형식이나 따옴표를 사용하지 않는다.

## 공개 전 확인

```powershell
git status --short
git ls-files | Select-String -Pattern "\.env.local$|config.test.yaml$|tnsnames.ora$|\.password$|output/|results/"
```

두 번째 명령에 실제 파일이 나오면 커밋하지 않는다.
