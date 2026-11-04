# netscout-auto v0.90 공개 E2E 검증 요약

> 이 문서의 서버, 계정, 장비 정보는 공개용 더미 데이터다. 실제 운영 식별정보와 비밀번호는 포함하지 않는다.

- 실행일: 2026-11-05
- NETSCOUT 예시: `https://netscout.example.com:8443`
- Device/Interface 예시: `DUMMY_DEVICE` / `DUMMY_INTERFACE`
- 실제 장비 변경: 0건
- 판정: PASS (기존 승인 앱 MATCHED 및 무변경 apply)

| 단계 | 결과 |
|---|---|
| pytest | 35 passed, 4 subtests passed |
| compileall/CLI | PASS |
| Oracle 19c SQLPlus/TNS Lab | PASS |
| Oracle Hostname | 입력 4 / FOUND 3 / NOT_FOUND 1 |
| Registration Excel 변환 | 3행 생성 / 미조회 1행 제외 |
| Inventory | Application 1,756 / Device 1 / Interface 3 |
| PURE 비교 | PURE 1,752 / CURRENT 1,756 / 후보 4 |
| Detail | success 10 / failed 0 |
| HTTPS 후보 | unique IP/Port 721 / unknown port 0 |
| TCP_OTHER 후보 | unique IP/Port 964 / unknown port 4 |
| Registration Plan | MATCHED 1 / 적용 0 |
| AFTER Inventory | Application 1,756 / Device 1 / Interface 3 |

실제 Excel, JSON, YAML, 서버 URL, 계정 및 비밀번호는 GitHub에 포함하지 않는다. 재현 절차는 `E2E_VALIDATION_SCENARIO_V0_90.md`와 `MANUAL_VALIDATION_GUIDE_V0_90.md`를 따른다.
