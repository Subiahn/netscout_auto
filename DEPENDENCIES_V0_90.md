# v0.90 프로그램 및 패키지 버전

2026-11-05 E2E 환경에서 확인한 버전이다.

| 구분 | 항목 | 버전 | 포함 여부 |
|---|---|---:|---|
| Runtime | Python | 3.12.3 | 별도 설치 |
| Python | PyYAML | 6.0.2 | `vendor` 포함 |
| Python | openpyxl | 3.1.5 | `vendor` 포함 |
| Python | requests | 2.32.3 | `vendor` 포함 |
| Python | urllib3 | 2.3.0 | `vendor` 포함 |
| Python | certifi | 2024.12.14 | `vendor` 포함 |
| Python | charset-normalizer | 3.4.1 | `vendor` 포함 |
| Python | idna | 3.10 | `vendor` 포함 |
| Oracle | Instant Client Basic | 19.33.0.0.0 | 라이선스·용량 때문에 Git 미포함 |
| Oracle | SQL*Plus | 19.33.0.0.0 | Git 미포함 |
| Lab DB | Oracle XE 11g R2 | 11.2.0.2.0 | Docker Lab |
| Container | Docker Engine | 29.6.1 | 별도 설치 |
| Container | Docker Compose | 5.3.0 | 별도 설치 |
| Ubuntu 진단 | netcat-openbsd | 배포판 제공 버전 | 선택 사항, 포트 점검용 |

`subprocess`, `argparse`, `pathlib`, `ipaddress`는 Python 표준 라이브러리이므로 별도 설치하지 않는다.

`vendor`는 v0.87 기준 캐시 제외 304개 파일과 경로 및 SHA-256이 동일하다. 포함된 Windows 바이너리 확장은 CPython 3.12 x64 기준이다. 다른 Python 버전이나 Ubuntu에서는 `requirements.txt`를 사용해 해당 환경에 맞게 다시 설치한다.

Oracle Client 19c의 공식 최소 DB 서버는 11.2.0.4다. 11.2.0.2 Lab에서는 기능 테스트가 성공했지만 공식 지원 조합은 아니다.
