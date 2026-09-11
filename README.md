# weekly-report-toolkit

팀 주간 보고를 **엑셀 하나**로 굴리고, 주간 미팅 발표본 HTML 은 스크립트가 뽑는다.
Claude Code 플러그인.

## 왜

주간 보고에서 흔히 깨지는 두 가지를 구조로 막는다.

**같은 걸 두 번 쓰지 않는다.** 「이번 주 완료 / 진행 중 / 다음 주 계획 / 이슈」를 따로 적는 칸을
두지 않는다. 전부 과제 행의 **비고 한 칸**에서 파생된다. 두 군데에 적으면 반드시 갈라진다.

**상태 색을 사람이 찍지 않는다.** 색은 상태 열에서 계산하고, 기한이 지난 과제는 다른 무엇보다
빨강이 이긴다. 손으로 찍게 하면 전부 초록이 된다(watermelon reporting).

## 설치

```
/plugin marketplace add kai-gme/weekly-report-toolkit
/plugin install weekly-report@weekly-report-toolkit
```

Claude 에게 「주간보고 발표본 뽑아줘」 + 엑셀 경로를 주면 된다.

스크립트만 쓰려면:

```bash
pip install openpyxl
python plugins/weekly-report/scripts/xlsx2html.py <주간보고.xlsx>
```

## 엑셀 형식

시트 5개 — `안내·목록` / `진행현황` / `배포이력` / `이슈` / `KPI(월)`.
`진행현황` 9열 = `★ / 도메인 / 과제명 / 담당 / 단계 (체크리스트) / 상태 / Due / 비고 (주간 로그) / Jira`.

팀 이름·도메인 목록·Jira 주소는 **코드에 없다** — `안내·목록` 시트와 `진행현황` 1~3행에서 읽는다.

비고 칸 접두어 3종:

```
다음: 화면설계 착수                  → 다음 주 계획
이슈: 킥오프 미개최 — 요청 대상       → 이슈
09/2주: 화면설계서 3/8 화면 작성      → 진행 중 (이번 주 한 일)
```

단계는 **담당이 바뀌는 지점에서만** 나눈다:

```
[x] 기획 — 담당A (09/2주)
[ ] 개발 — 담당B
[ ] QA — 담당C
```

자세한 규칙은 [`plugins/weekly-report/skills/weekly-report/SKILL.md`](plugins/weekly-report/skills/weekly-report/SKILL.md).

## 라이선스

MIT
