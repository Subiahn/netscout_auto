# v0.9.1 다중 장비·인터페이스 등록 자동화 가이드

이 문서의 주소와 이름은 모두 **더미 데이터**다. 실제 설정 및 비밀번호 파일은 Git에 커밋하지 않는다.

## 전체 흐름

```text
Inventory
  → collection_scope.yaml 생성·검토
  → 전체 Interface Top-N 수집
  → detail_scope.yaml 생성·검토
  → 전체 Detail 수집
  → IP/Port 후보 통합
  → registration_review.xlsx 생성·검토
  → 기존 registration plan/apply/verify
```

## 1. Inventory 생성

```powershell
python -m netscout inventory `
  --config .\config\config.test.yaml `
  --env .\.env.local `
  --output-dir .\output
```

로그에 출력된 `Netscout_Inventory_*.xlsx` 경로를 다음 단계에 사용한다.

## 2. 장비·인터페이스 범위 자동 생성

```powershell
python .\tools\registration\registration_pipeline.py scope-generate `
  --inventory .\output\YYYYMMDD\00_inventory\Netscout_Inventory_YYYYMMDD_HHMMSS.xlsx `
  --output .\config\collection_scope.yaml `
  --duration LAST_24_HOURS `
  --top-n 50
```

생성된 YAML을 열어 수집하지 않을 Interface만 `enabled: false`로 바꾼다. 이름이 같은 Interface라도 번호가 다르면 각각 보존된다.

## 3. 활성 Interface 전체 Top-N 수집

먼저 호출 없이 계획을 확인한다.

```powershell
python -m netscout run-targets `
  --config .\config\config.test.yaml `
  --env .\.env.local `
  --targets .\config\collection_scope.yaml `
  --output-dir .\output `
  --dry-run
```

확인 후 `--dry-run`만 제거하여 실행한다. 일부 Interface가 실패해도 나머지는 계속 처리되며 실패 이유가 로그에 남는다.

## 4. 전체 Top-N에서 Detail 범위 생성

```powershell
python .\tools\registration\registration_pipeline.py detail-scope-build `
  --top-root .\output\YYYYMMDD\01_top `
  --output .\config\detail_scope.yaml `
  --detail-top-n 20 `
  --limit 1000
```

`detail_scope.yaml`에서 실제 상세 조회 대상 Application을 검토하고 불필요한 항목은 `enabled: false`로 바꾼다.

## 5. 전체 Detail 수집

```powershell
python -m netscout run-detail-targets `
  --config .\config\config.test.yaml `
  --env .\.env.local `
  --targets .\config\detail_scope.yaml `
  --output-dir .\output `
  --dry-run
```

계획 확인 후 `--dry-run`을 제거하여 실행한다.

## 6. IP/Port 후보 통합

```powershell
python .\tools\registration\registration_pipeline.py candidates-build `
  --detail-root .\output\YYYYMMDD\03_detail `
  --output .\output\YYYYMMDD\04_candidates\all_candidates.xlsx `
  --yaml-output .\output\YYYYMMDD\04_candidates\all_candidates.yaml
```

동일 Application/IP/Port의 발견 결과가 합쳐지고 장비·인터페이스 출처와 근거 건수가 보존된다.

## 7. 등록 정책 준비

```powershell
Copy-Item .\config\registration_policy.example.yaml .\config\registration_policy.yaml
notepad .\config\registration_policy.yaml
```

예시:

```yaml
defaults:
  registration_type: "SBA"
  parent_application: "TCP"
  protocol: "TCP"
  application_name: ""

application_rules:
  - match: "HTTPS"
    application_name: "DUMMY_HTTPS_{port}"
    parent_application: "TCP"
    protocol: "TCP"
```

매핑되지 않은 Application 이름은 자동으로 추측하지 않고 빈 값과 `Review` 사유로 남는다.

## 8. 등록 검토 Excel 생성

```powershell
python .\tools\registration\registration_pipeline.py review-build `
  --candidates .\output\YYYYMMDD\04_candidates\all_candidates.xlsx `
  --policy .\config\registration_policy.yaml `
  --output .\output\YYYYMMDD\04_registration_review\registration_review.xlsx
```

모든 행은 `Register=N`으로 생성된다. `ApplicationName`, `RegistrationType`, `ParentApplication`, `Protocol`, IP/Port 및 근거 열을 검토한 뒤 승인할 행만 `Y`로 바꾼다.

## 9. 기존 등록 모듈로 계획 확인

```powershell
$env:NETSCOUT_PASSWORD = Read-Host "NETSCOUT password" -MaskInput

python .\tools\registration\netscout_sync.py `
  --input .\output\YYYYMMDD\04_registration_review\registration_review.xlsx `
  --base-url https://netscout.example.com:8443 `
  --username DUMMY_ADMIN `
  --insecure
```

이 단계는 `--apply`가 없으므로 장비를 변경하지 않는다. `CREATE`, `ADD_IP`, `ADD_PORT`, `MATCHED`, `CONFLICT` 결과를 검토한다.

## 10. 적용 및 재검증

계획 승인 후에만 같은 명령에 `--apply`를 추가한다. 적용 후 `--apply` 없이 다시 실행하여 대상이 `MATCHED`인지 확인한다.

## Ubuntu

명령의 `python`을 `python3`으로 바꾸고 PowerShell 줄 연결 문자 `` ` ``를 Bash의 `\`로 바꾸면 된다. 비밀번호는 다음처럼 입력한다.

```bash
read -rsp "NETSCOUT password: " NETSCOUT_PASSWORD
export NETSCOUT_PASSWORD
echo
```
