# v0.90 파일명을 지정하지 않는 수동 테스트 가이드

이 문서의 주소·계정·장비명은 모두 **더미 데이터**다. 실제 값은 Git에 저장하지 않고 로컬 설정 파일에만 입력한다.

## 1. 동작 원칙

공통 CLI의 `--output-file`을 생략하면 실행 시각을 포함한 Excel 파일명이 자동으로 생성된다.

```text
<prefix>_YYYYMMDD_HHMMSS.xlsx
```

`--output-dir`도 생략하면 현재 디렉터리를 기준으로 날짜 및 단계별 폴더가 생성된다. 반복 실행해도 타임스탬프가 달라 기존 결과를 덮어쓰지 않는다.

## 2. Windows PowerShell

프로젝트 디렉터리에서 실행한다.

```powershell
cd C:\path\to\netscout_auto
```

저장소 안에 `netscout-auto` 하위 디렉터리를 사용하는 배치라면 그 디렉터리까지 이동한다.

### Inventory

```powershell
python -m netscout inventory `
  --config .\config\config.test.yaml `
  --env .\.env.local `
  --output-dir .\output
```

자동 생성 예:

```text
output\YYYYMMDD\00_inventory\Netscout_Inventory_YYYYMMDD_HHMMSS.xlsx
```

### Top N

```powershell
python -m netscout top `
  --config .\config\config.test.yaml `
  --env .\.env.local `
  --device DUMMY_DEVICE `
  --interface DUMMY_INTERFACE `
  --duration LAST_HOUR `
  --top-n 50 `
  --output-dir .\output
```

자동 생성 예:

```text
output\YYYYMMDD\01_top\DUMMY_DEVICE_DUMMY_INTERFACE_Top50_LAST_HOUR_YYYYMMDD_HHMMSS.xlsx
```

### Application Detail

```powershell
python -m netscout detail `
  --config .\config\config.test.yaml `
  --env .\.env.local `
  --device DUMMY_DEVICE `
  --interface DUMMY_INTERFACE `
  --application HTTPS `
  --duration LAST_HOUR `
  --output-dir .\output
```

자동 생성 예:

```text
output\YYYYMMDD\03_detail\HTTPS\DUMMY_DEVICE_DUMMY_INTERFACE_HTTPS_Detail_LAST_HOUR_YYYYMMDD_HHMMSS.xlsx
```

### 등록 검토 Excel

`make_registration.py`의 `--output`을 생략하면 현재 디렉터리에 `registration.xlsx`가 생성된다.

```powershell
python .\tools\registration\make_registration.py `
  --input .\output\YYYYMMDD\05_hostname\HTTPS\hostname_lookup.xlsx
```

## 3. Ubuntu

PowerShell과 동일하게 `--output-file`을 쓰지 않는다.

```bash
cd /path/to/netscout_auto

python3 -m netscout inventory \
  --config ./config/config.test.yaml \
  --env ./.env.local \
  --output-dir ./output
```

가장 최근 Inventory 결과 확인:

```bash
find ./output -type f -name 'Netscout_Inventory_*.xlsx' -printf '%T@ %p\n' | sort -nr | head -1
```

Windows PowerShell에서 가장 최근 결과 확인:

```powershell
Get-ChildItem .\output -Recurse -Filter 'Netscout_Inventory_*.xlsx' |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1 FullName, LastWriteTime
```

## 4. 생략할 수 없는 파일 옵션

다음 명령의 출력 경로는 현재 CLI에서 필수이므로 생략하면 안 된다.

- `extract-candidates --output`
- `compare-inventory-apps --output`
- `filter-detail-targets --output`
- `tools/oracle_query_hostname.py --output`
- `tools/hostname_lookup_from_candidates.py --output`
- `tcp-other-report --output`
- `registration-candidates --output`

이 옵션을 생략하면 `argparse`가 필수 인자 오류를 출력하며 파일은 생성되지 않는다.

## 5. 결과 판정

1. 명령 종료 코드가 `0`인지 확인한다.
2. 로그의 `output=` 경로를 확인한다.
3. 해당 경로에 Excel 파일이 존재하는지 확인한다.
4. 같은 명령을 다시 실행했을 때 새 타임스탬프 파일이 생기고 기존 파일이 유지되는지 확인한다.
5. `registration.xlsx`는 고정 기본 이름이므로 같은 위치에서 다시 실행하면 덮어쓸 수 있다. 보존이 필요하면 `--output`을 직접 지정한다.
