# NETSCOUT Automation v0.87(0.8.7) → v0.90(0.9.0) 업그레이드 가이드

## 1. 권장 방식

기존 v0.87 폴더에 파일을 덮어쓰지 않고 v0.90을 새 폴더에 병행 설치한다. 조회·Oracle·등록 계획 검증이 끝난 뒤 실행 경로만 v0.90으로 전환한다. 문제가 발생하면 v0.87 폴더로 즉시 돌아갈 수 있다.

```text
C:\netscout-automation\
├─ netscout-auto_v0.87\netscout-auto\   # 기존 버전, 보존
└─ netscout-auto_v0.90\                 # 신규 버전
```

## 2. 주요 변경 사항

| 구분 | v0.87 | v0.90 |
|---|---|---|
| NETSCOUT 조회 | Inventory/Top/Detail | 동일 기능과 버전 표기 정제 |
| 수동 앱 필터 | PURE Inventory 비교 | 동일 기능 유지 및 E2E 재검증 |
| 환경 파일 | 루트 `.env.local` | `config\.env.local` |
| Hostname | HTTP 기반 조회 | HTTP + Oracle SQL*Plus/TNS |
| 등록 | 없음 | CREATE/MATCHED/ADD_IP/ADD_PORT/CONFLICT 계획 및 승인 적용 |
| Oracle | 없음 | Instant Client 19c, TNS, Oracle 11g Lab |
| 테스트 | 기존 테스트 | 35 tests + 4 subtests |

`vendor`는 v0.87과 v0.90의 캐시 제외 304개 파일이 상대 경로와 SHA-256까지 동일하다. 기존 vendor를 그대로 사용해도 되지만, 새 폴더 병행 설치에서는 v0.90에 포함된 vendor를 그대로 두는 것을 권장한다.

## 3. 업그레이드 전 백업

PowerShell:

```powershell
$old = "C:\netscout-auto_v0.87"
$backup = "C:\netscout-automation\backup_v087_$(Get-Date -Format yyyyMMdd_HHmmss)"

New-Item -ItemType Directory -Force -Path $backup | Out-Null
Copy-Item "$old\.env.local" "$backup\.env.local" -Force
Copy-Item "$old\config" "$backup\config" -Recurse -Force
Copy-Item "$old\pure_xlsx" "$backup\pure_xlsx" -Recurse -Force -ErrorAction SilentlyContinue
```

`output` 전체가 필요하면 별도 보관한다. 새 v0.90의 output과 합치지 않는다.

## 4. v0.90 코드 준비

검증된 폴더를 사용한다.

```powershell
Set-Location C:\netscout-auto_v0.90
```

다음 핵심 파일이 존재해야 한다.

```text
netscout.py
V8\v8_cli.py
src\netscout_automation\
tools\hostname_lookup_from_candidates.py
tools\oracle_query_hostname.py
tools\registration\
vendor\
```

## 5. 기존 환경값 이전

### 5.1 NETSCOUT 계정

v0.87:

```text
.env.local
```

v0.90:

```text
config\.env.local
```

복사:

```powershell
Copy-Item `
  "C:\netscout-auto_v0.87\.env.local" `
  "C:\netscout-auto_v0.90\config\.env.local" `
  -Force
```

필수 키:

```text
NETSCOUT_USERNAME
NETSCOUT_PASSWORD
```

### 5.2 NETSCOUT API 설정

v0.87과 v0.90의 `config.test.yaml`은 모두 `netscout`, `api`, `output` 구조를 사용한다. 기존 파일을 무조건 덮어쓰기보다 v0.90 예제를 기준으로 서버 주소, SSL, timeout, 장비값을 옮긴다.

```text
config\config.test.yaml
```

### 5.3 PURE Inventory

화이트리스트 비교 기준 파일을 그대로 복사한다.

```powershell
Copy-Item `
  "C:\netscout-auto_v0.87\pure_xlsx\pure_inventory.xlsx" `
  "C:\netscout-auto_v0.90\pure_xlsx\pure_inventory.xlsx" `
  -Force
```

### 5.4 이전 생성 YAML

다음 v0.87 파일은 과거 실행 산출물이므로 v0.90 `config`에 복사하지 않는다.

```text
config\targets.yaml
config\detail_targets.yaml
config\detail_targets_filtered.yaml
config\ip_port_candidates.yaml
config\manual_registered_apps.yaml
V8\config\targets.yaml
```

필요하면 백업 폴더에만 보관한다. v0.90에서는 날짜별 `output\YYYYMMDD` 아래에서 다시 생성한다.

## 6. Oracle TNS 설정

Oracle 조회를 사용할 때 필요한 파일:

```text
tools\oracle_query_hostname.py
tools\hostname_lookup_from_candidates.py
tools\oracle\instantclient_19_33\       # Git 제외, 별도 설치
config\oracle_hostname.yaml
config\oracle_hostname.password
config\oracle_network\tnsnames.ora
```

`config\oracle_hostname.yaml` 주요 값:

```yaml
oracle:
  connect_identifier: "HOSTNAME_DB"
  username: "NETSCOUT_READONLY"
  tns_admin: "oracle_network"
  sqlplus_command:
    - "tools/oracle/instantclient_19_33/sqlplus.exe"
    - "-s"
    - "/nolog"
  password_file: "oracle_hostname.password"
```

실제 비밀번호와 `tnsnames.ora`는 Git에 올리지 않는다.

## 7. 등록 모듈 추가 파일

v0.87에 다음 디렉터리를 통째로 추가한다는 개념이지만, 실제 배포에서는 v0.90 폴더에 포함된 파일을 그대로 사용한다.

```text
tools\registration\
├─ make_registration.py
├─ netscout_api.py
├─ netscout_sync.py
├─ registration_rules.py
├─ sync_planner.py
└─ tests\
```

기존 `netscout.py`에 등록 기능을 합치지 않는다. 조회 CLI와 등록 CLI를 분리한 구조를 유지한다.

## 8. 업그레이드 후 정적 검증

```powershell
Set-Location C:\netscout-auto_v0.90

python -m pytest -q
python -m compileall -q netscout.py src V8 tools
python -m netscout --help
python .\tools\registration\netscout_sync.py --help
```

통과 기준:

```text
35 passed, 4 subtests passed
compileall 오류 0
두 CLI help 정상
```

## 9. 조회 E2E 검증

### Inventory

```powershell
python -m netscout inventory `
  --config .\config\config.test.yaml `
  --env .\config\.env.local `
  --output-dir .\output `
  --output-file Inventory_UPGRADE_BASELINE.xlsx
```

### Top N

```powershell
python -m netscout top `
  --config .\config\config.test.yaml `
  --env .\config\.env.local `
  --device "DUMMY_DEVICE" `
  --interface "DUMMY_INTERFACE" `
  --duration LAST_7_DAYS `
  --top-n 10 `
  --output-dir .\output
```

### PURE 비교

```powershell
python -m netscout compare-inventory-apps `
  --pure-inventory .\pure_xlsx\pure_inventory.xlsx `
  --current-inventory .\output\YYYYMMDD\00_inventory\Inventory_UPGRADE_BASELINE.xlsx `
  --output .\output\YYYYMMDD\02_targets\manual_registered_apps.csv `
  --yaml-output .\output\YYYYMMDD\02_targets\manual_registered_apps.yaml `
  --sheet Application_List `
  --key-column ApplicationTypeCode
```

이후 Detail 대상 생성 → 필터 → dry-run → 실제 Detail → 후보 추출 순으로 `MANUAL_VALIDATION_GUIDE_V0_90.md`를 따른다.

## 10. Oracle TNS 검증

```powershell
.\tools\oracle\instantclient_19_33\sqlplus.exe -V

python .\tools\oracle_query_hostname.py `
  --config .\config\oracle_hostname.yaml `
  --ip-candidates-yaml .\docker\hostname-lab\sample_ip_port_candidates.yaml `
  --output .\output\oracle_hostname_upgrade_test.xlsx
```

운영 후보를 사용할 때는 `--ip-candidates-yaml`을 실제 `output\YYYYMMDD\04_candidates\...` 파일로 변경한다.

## 11. 등록 안전 검증

먼저 `--apply` 없이 Plan만 실행한다.

```powershell
$env:NETSCOUT_PASSWORD = Read-Host "NETSCOUT password"

python .\tools\registration\netscout_sync.py `
  --input .\registration_review.xlsx `
  --base-url https://NETSCOUT_HOST:8443 `
  --username DUMMY_ADMIN `
  --insecure
```

확인 사항:

- 예상한 `CREATE`, `MATCHED`, `ADD_IP`, `ADD_PORT`만 존재
- `CONFLICT`가 없어야 함
- 계획 파일과 Excel을 사람이 검토

실제 적용은 승인 후에만 실행한다.

```powershell
python .\tools\registration\netscout_sync.py `
  --input .\registration_review.xlsx `
  --base-url https://NETSCOUT_HOST:8443 `
  --username DUMMY_ADMIN `
  --insecure `
  --apply
```

## 12. 전환 및 롤백

### 전환

- 작업 스케줄러 또는 배치 파일의 작업 디렉터리를 v0.90으로 변경한다.
- 명령의 `--env`를 반드시 `.\config\.env.local`로 변경한다.
- 첫 운영 실행은 Inventory와 Plan-only로 제한한다.

### 롤백

- v0.90 실행을 중지한다.
- 작업 디렉터리를 기존 v0.87 경로로 되돌린다.
- v0.87의 `.env.local`, `config`, `pure_xlsx` 백업을 그대로 사용한다.

등록 모듈은 Application 삭제와 기존 IP/Port 제거를 수행하지 않으므로, 자동 롤백 대상은 코드와 실행 경로다. 이미 생성된 Application을 삭제해야 한다면 별도 운영 승인과 수동 절차가 필요하다.

## 13. 최종 PASS 기준

- pytest/compileall/CLI PASS
- Inventory/Top/Detail 실패 0
- PURE 비교 결과 검토 완료
- Oracle TNS 접속과 Hostname Excel 생성 성공
- Registration Plan이 예상 상태
- `CONFLICT` 0
- 적용 후 GET 재검증 성공
- 동일 입력 재실행 시 `MATCHED`
- BEFORE/AFTER Application 수가 예상과 일치

