# netscout-auto v0.9.1 릴리스 노트

## 추가 기능

- Inventory Excel에서 전체 장비·인터페이스 수집 범위 YAML 생성
- 여러 Top-N Excel을 하나의 Detail 실행 범위로 통합
- 여러 장비·인터페이스의 Detail 결과에서 IP/Port 후보 통합
- Application 정책 YAML을 이용한 등록 검토 Excel 초안 생성
- 검토 Excel에 장비 수, 인터페이스 수, 근거 건수와 출처 목록 기록

## 안전 원칙

- 새 파이프라인은 NETSCOUT 설정을 변경하지 않는다.
- 검토 Excel의 `Register` 기본값은 항상 `N`이다.
- 이름이나 유형을 확정할 수 없으면 `Review`에 사유를 기록한다.
- 기존 검토 Excel은 덮어쓰지 않는다.
- 실제 적용은 기존 `netscout_sync.py --apply`를 명시해야만 실행된다.
