# Oracle 11g Hostname 조회 Lab

이 Lab은 Oracle Database 11g Release 2 XE와 SQL*Plus를 사용해 `ip_port_candidates.yaml`의 IP를 Hostname으로 조회하는 기능을 검증한다.

> 이 Lab의 주소, 사용자, 서비스명 및 조회 데이터는 기능 검증용 **더미 데이터**입니다. 운영 DB 접속정보를 이 디렉터리에 저장하거나 Git에 커밋하지 마십시오.

## 주의

- `gvenzl/oracle-xe:11`은 Oracle 공식 이미지가 아닌 커뮤니티 이미지다. 로컬 기능 테스트 전용으로 사용한다.
- 운영 환경에서는 회사가 승인한 Oracle 이미지와 라이선스 정책을 따른다.
- 11g XE 서비스 이름은 `XE`이며 PDB가 없다.
- Docker Desktop은 Linux 컨테이너 모드와 amd64 실행 환경이 필요하다.
- Oracle XE는 최소 1GB shared memory가 필요하므로 Compose에 `shm_size: 1gb`를 지정했다.

## 1. 비밀번호 파일 준비

프로젝트 루트에서 실행한다.

```powershell
Set-Content .\config\oracle11g-lab-system.password "로컬_SYSTEM_비밀번호"
Set-Content .\config\oracle11g-lab-app.password "로컬_APP_비밀번호"
```

두 파일은 Git에 커밋하지 않는다. Python 도구도 `config\oracle11g-lab-app.password`를 사용한다.

## 2. Oracle 11g Lab 시작

```powershell
docker compose -f .\docker\oracle11g-lab\docker-compose.yml up -d
docker compose -f .\docker\oracle11g-lab\docker-compose.yml ps
docker logs -f netscout-oracle11g-lab
```

첫 기동은 데이터베이스 초기화 때문에 오래 걸린다. 로그에 `DATABASE IS READY TO USE`가 표시되고 health 상태가 healthy가 된 후 진행한다.

## 3. SQL*Plus 직접 확인

컨테이너 SQL*Plus를 실행한다.

```powershell
docker exec -it netscout-oracle11g-lab sqlplus NETSCOUT_LAB@//localhost:1521/XE
```

비밀번호 프롬프트에 `config\oracle11g-lab-app.password`의 값을 입력한 후 확인한다.

```sql
SELECT * FROM HOSTNAME_LOOKUP ORDER BY IP_ADDRESS;
```

## 4. Python 조회 테스트

```powershell
python .\tools\oracle_query_hostname.py `
  --config .\config\oracle_hostname.lab.yaml `
  --ip-candidates-yaml .\docker\hostname-lab\sample_ip_port_candidates.yaml `
  --output .\output\oracle11g_hostname_lookup.xlsx
```

결과 Excel 구조는 기존 Hostname 조회와 동일하다.

- `Hostname_Lookup`
- `Run_Summary`

## 5. 운영 TNS 설정

```powershell
Copy-Item .\config\oracle_hostname.example.yaml .\config\oracle_hostname.yaml
notepad .\config\oracle_hostname.yaml
```

DBA에게 받은 다음 정보를 반영한다.

- TNS alias 또는 Host/Port/Service Name
- 읽기 전용 사용자명
- 비밀번호
- IP 컬럼과 Hostname 컬럼이 포함된 SELECT 문
- `tnsnames.ora` 위치

TNS alias를 사용할 경우 `oracle_hostname.yaml`의 `tns_admin`에 `tnsnames.ora`가 있는 디렉터리를 직접 입력한다. 도구가 SQL*Plus 프로세스에만 `TNS_ADMIN`을 전달하므로 Windows 전역 환경변수를 따로 변경하지 않는다. `config\tnsnames.example.ora`를 형식 참고용으로 사용할 수 있다. Easy Connect를 사용하면 `//host:port/service_name`을 직접 지정할 수 있다.

실행:

```powershell
python .\tools\oracle_query_hostname.py `
  --config .\config\oracle_hostname.yaml `
  --ip-candidates-yaml .\output\YYYYMMDD\04_candidates\HTTPS\ip_port_candidates.yaml `
  --output .\output\YYYYMMDD\05_hostname\HTTPS\hostname_lookup_oracle.xlsx
```

## 6. 종료

컨테이너만 종료하고 데이터는 유지한다.

```powershell
docker compose -f .\docker\oracle11g-lab\docker-compose.yml down
```

Lab 데이터까지 삭제하려면 볼륨 삭제가 필요하다. 이 작업은 복구할 수 없으므로 별도 승인 후 수행한다.
