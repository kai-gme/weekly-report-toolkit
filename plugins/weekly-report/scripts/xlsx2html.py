# -*- coding: utf-8 -*-
"""주간보고 엑셀 → 미팅 발표본 HTML.

사용:  python xlsx2html.py <주간보고.xlsx>   |   --selftest
엑셀이 원본(SSOT), HTML 은 매주 여기서 다시 뽑는다. HTML 을 손으로 고치지 말 것.

보이는 모양 = 흑백 + 얇은 회색 선, 불릿 머리말, 체크박스 리스트. 장식 금지.
이슈·협업요청·팀운영은 한 절로 통합 · KPI 는 월별 하단 · ★ 주요 과제 · 비고 = 주간 히스토리.

팀마다 다른 값(Jira 주소·도메인 목록)은 **코드가 아니라 워크북**에서 읽는다 — `sheet_config()` 참고.
도메인은 `안내·목록` 의 **드롭다운 원본 열을 그대로** 읽는다(이미 있는 걸 읽지, 설정 문법을 새로
만들어 사람에게 또 적게 하지 않는다). 없으면 각각 「링크 없음」·「오타 검출 없음」 으로 돌고 경고를 낸다.
"""
import sys, os, io, html, re, shutil
from datetime import date, datetime
from openpyxl import load_workbook

# Windows 기본 콘솔은 cp949 라 경고의 한글·기호에서 UnicodeEncodeError 로 죽는다.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):        # 파이프로 넘길 때는 조용히 넘어간다
    pass

HEADER_ROW = 6                      # 그 아래가 전부 데이터
# 예시 행은 «위치» 가 아니라 «과제명이 (예시) 로 시작하나» 로 가른다 —
# 위치로 박으면 예시 행을 지운 워크북에서 첫 과제가 조용히 사라진다.
EXAMPLE = "(예시)"
OTHER = "기타"                      # 도메인이 비었거나 목록 밖일 때 접어 넣는 칸


def cells(ws, r, n):
    return [ws.cell(row=r, column=c).value for c in range(1, n + 1)]


def esc(v):
    return html.escape(str(v)) if v not in (None, "") else ""


# 상태 → 색. Jira 상태 카테고리와 같은 축(To Do 회색 · In Progress 파랑 · Done 초록).
DOT = {"완료": "done", "진행중": "go", "준비중": "", "보류중": "warn"}


def signal(status, due):
    """색은 상태 열에서 계산한다 — 사람이 손으로 찍지 않는다(찍게 하면 전부 초록이 된다).

    Due 가 지났는데 안 끝났으면 그게 다른 모든 것을 이긴다(🔴).
    """
    st = str(status or "").strip()
    if st != "완료":
        d = due.date() if isinstance(due, datetime) else due
        if not isinstance(d, date):
            m = re.match(r"\s*(\d{4})-(\d{2})-(\d{2})", str(due or ""))
            d = date(*map(int, m.groups())) if m else None
        if d and (d - date.today()).days < 0:
            return "bad"
    return DOT.get(st, "")


def status(v, cls, note):
    """활동 상태 + 색점 + 체크리스트 진행률. 텍스트만 있으면 표에서 눈에 안 띈다."""
    if not v:
        return ""
    done, doing, todo = checklist(note)
    total = len(done) + len(doing) + len(todo)
    cur = doing or todo                      # 지금 이 단계가 있으면 그쪽이 현재 담당이다
    who = cur[0][1] if cur else ""
    pct = (f'<div class="w">{len(done)}/{total}'
           f'{" · " + esc(who) if who else ""}</div>') if total else ""
    return f'<span class="d {cls}"></span>' + esc(str(v).strip()) + pct


BOX = re.compile(r"^\[([ xXoO])\]\s*(.*)$", re.M)   # x=완료 · o=지금 이 단계 · 공백=예정
ITEM = re.compile(r"^(.*?)\s+—\s+(.+)$")        # '개발 — 개발자' → 항목 / 담당
WEEK = re.compile(r"\((\d{1,2}/\d주)\)\s*$")     # '… (09/3주)' → 완료 주차


def norm_week(s):
    """'9/2주' → '09/2주'. 받는 형식과 비교하는 형식이 다르면 조용한 누락이 된다.

    파서(WEEK·weeklog)는 한 자리 월을 받아주는데 2행에서 만든 라벨은 늘 zero-pad 라,
    정규화 없이 문자열 완전일치로 비교하면 '9/2주' 로 적은 행이 전 절에서 증발한다.
    """
    m = re.match(r"\s*(\d{1,2})/(\d)주\s*$", str(s or ""))
    return f"{int(m.group(1)):02d}/{m.group(2)}주" if m else str(s or "").strip()


def meta_items(txt):
    """머리말 한 행(2행·3행) → `<li>` 안에 들어갈 항목들.

    2·3행은 사람이 '·' 로 끊어 쓴 문장이라 그 구분자를 그대로 항목 경계로 쓴다.
    'k: v' 면 k 를 굵게, 아니면 통째로. 값에 ':' 가 또 있어도(10:30) 첫 ':' 에서만 끊는다.
    """
    for part in (p.strip() for p in str(txt or "").split("·")):
        if not part:
            continue
        k, _, v = part.partition(":")
        yield (f"<strong>{esc(k.strip())}</strong>: {esc(v.strip())}"
               if v.strip() else esc(part))


def week_label(txt):
    """'주차: 2026년 9월 2주차 …' → '09/2주'. 못 읽으면 ''."""
    m = re.search(r"(\d{1,2})월\s*(\d)주차", str(txt or ""))
    return f"{int(m.group(1)):02d}/{m.group(2)}주" if m else ""


def fold_domain(v, domains):
    """(표에 낼 도메인, 엑셀에 적힌 값). 목록 밖·빈 칸은 '기타' 로 접는다.

    접지 않으면 ★ 행에서 자동번호가 없어 산출물 전체가 KeyError 로 죽는다.
    접은 사실은 경고로 남긴다 — 조용히 접으면 빠뜨린 건지 진짜 기타인지 구별이 안 된다.
    `domains` 가 비면(엑셀에 목록 선언이 없으면) 접지 않는다 = 오타를 검출할 근거가 없다.
    """
    raw = str(v).strip() if v not in (None, "") else ""
    if not domains:
        return (raw or OTHER), raw
    return (raw if raw in domains else OTHER), raw


GUIDE = "안내·목록"                  # 드롭다운 원본이자 설정이 사는 시트
BLANK_GAP = 3                       # 헤더와 목록 사이에 이만큼 비어 있어도 목록으로 본다


def safe_url(v):
    """링크로 쓸 주소. http(s) 가 아니면 버린다.

    이 값은 엑셀에서 오는 사용자 입력이고 `href` 에 그대로 들어간다.
    `esc()` 는 속성 밖으로 나가는 것만 막지 `javascript:`·`data:` 스킴은 못 막는다.
    """
    s = str(v or "").strip()
    return s if re.match(r"^https?://", s, re.I) else ""


def lookup_column(wb, header):
    """`안내·목록` 에서 header 칸을 찾아 그 아래 칸들을 목록으로.

    🔑 도메인 목록은 **드롭다운 원본이라 이미 워크북 안에 있다** — 코드에 복사하지도,
    설정 문법을 새로 만들어 사람에게 또 적게 하지도 않는다. 있는 걸 읽는다.
    """
    if GUIDE not in wb.sheetnames:
        return []
    g = wb[GUIDE]
    for row in g.iter_rows():
        for c in row:
            if str(c.value or "").strip() != header:
                continue
            out, blank = [], 0
            for r in range(c.row + 1, g.max_row + 1):
                v = g.cell(row=r, column=c.column).value
                if v in (None, ""):
                    # 헤더 바로 아래 한 줄쯤 비워 두는 시트가 흔하다. 값이 한 번이라도
                    # 나온 뒤의 빈 칸이 진짜 끝이다 — 첫 칸에서 끊으면 「칸이 없다」로 오진한다.
                    blank += 1
                    if out or blank > BLANK_GAP:
                        break
                    continue
                if str(v).strip() == header:
                    break           # 헤더가 또 나왔다 = 다음 블록. 헤더 글자는 값이 아니다
                out.append(str(v).strip())
            if out:
                return out
    return []


def lookup_value(wb, label):
    """`안내·목록` 에서 label 칸의 **오른쪽** 값. 없으면 ''.

    완전일치를 먼저 본다 — `startswith` 만 쓰면 「Jira 주소 안내: …」 같은 산문 칸이
    진짜 라벨보다 위에 있을 때 그쪽을 집는다.
    """
    if GUIDE not in wb.sheetnames:
        return ""
    g = wb[GUIDE]
    loose = ""
    for row in g.iter_rows():
        for c in row:
            t = str(c.value or "").strip()
            if not t.startswith(label):
                continue
            v = g.cell(row=c.row, column=c.column + 1).value
            v = str(v).strip() if v not in (None, "") else ""
            if t == label:
                return v
            loose = loose or v
    return loose


def sheet_config(wb):
    """워크북에서 팀 고유값을 읽는다 → {'Jira': …, '도메인': [...]}.

    팀마다 다른 값이라 코드에 박지 않는다. 둘 다 없으면 빈 값으로 돌고 경고가 뜬다.
    """
    cfg = {}
    doms = lookup_column(wb, "도메인")
    if doms:
        if OTHER not in doms:
            doms.append(OTHER)          # 접을 칸이 없으면 접기가 성립하지 않는다
        cfg["도메인"] = doms
    jira = lookup_value(wb, "Jira 주소")
    if jira:
        cfg["Jira"] = jira
    return cfg


def domain_heading(n, i, dom):
    """§1 의 도메인 절 제목.

    도메인은 이제 코드 상수가 아니라 **엑셀에서 온 사용자 입력**이다. 이스케이프를 빼면
    `</h3><h2>가짜 절</h2>` 같은 값으로 발표본의 절 구조를 위조할 수 있다.
    """
    return f"<h3>{n}.{i} {esc(dom)}</h3>"


def domain_order(rows, declared):
    """§1 에 낼 도메인과 그 순서. 선언이 있으면 그대로, 없으면 엑셀에 나온 순서.

    🔴 **빈 리스트를 돌려주면 안 된다** — 자동번호가 한 행도 안 매겨져 ★ 표가 KeyError 로 죽는다.
    """
    return list(declared) or list(dict.fromkeys(d["domain"] for d in rows))


def split_item(txt):
    """'개발 — 개발자 (09/3주)' → (항목, 담당, 주차). 없는 조각은 빈 문자열."""
    t = str(txt).strip()
    m = WEEK.search(t)
    week = norm_week(m.group(1)) if m else ""
    if m:
        t = t[:m.start()].strip()
    m = ITEM.match(t)
    return (m.group(1), m.group(2).strip(), week) if m else (t, "", week)


def checklist(txt):
    """체크 줄을 판다 → (완료, 지금, 예정). 각 원소 = (항목, 담당, 주차).

    `[x]` 끝남 · `[o]` **지금 이 단계** · `[ ]` 아직. 「지금 담당」은 `[o]` 가 있으면 그쪽이 먼저다 —
    `[ ]` 첫 줄로만 보면 아직 시작도 안 한 단계의 사람이 현재 담당으로 잡힌다.

    진행률·이번 주 완료·진행 중·다음 계획이 전부 여기서 파생된다.
    같은 걸 '주간' 시트에 또 쓰지 않는다 — 두 번 쓰면 갈라진다.
    """
    hit = [m for m in (BOX.match(l.strip()) for l in str(txt or "").splitlines()) if m]
    done = [split_item(m.group(2)) for m in hit if m.group(1) in "xX"]
    doing = [split_item(m.group(2)) for m in hit if m.group(1) in "oO"]
    todo = [split_item(m.group(2)) for m in hit if m.group(1) == " "]
    return done, doing, todo


def done_fallback(d, week):
    """단계에 이번 주 `[x]` 가 없는 행을 §이번 주 완료로 올릴까.

    단계를 안 쪼갰거나 지난 주에 이미 다 끝낸 과제 — 상태가 '완료' 고 이번 주 로그가 있으면 올린다.
    `not d["done"]` 로 보면 지난 주에 단계를 다 끝낸 행이 영영 안 올라온다.
    """
    return (not any(i[2] == week for i in d["done"]) and bool(d["log"])
            and str(d["status"]).strip() == "완료")


def progress_fallback(d):
    """남은 `[ ]` 가 없는 행을 §진행 중으로 올릴까 (이번 주 로그가 있는 행만 들어온다).

    단계를 안 쪼갰거나 전부 `[x]` 인데 아직 릴리즈 전인 과제. 여기에 `not d["done"]` 을 얹으면
    「로그를 썼는데 발표본 어느 절에도 안 나오는」 행이 생긴다 — 이 도구의 존재 이유를 부정한다.
    """
    return not (d["doing"] or d["todo"]) and str(d["status"]).strip() != "완료"


def nextline(txt):
    """비고의 '다음: ' 줄 = 다음 주 할 일. 없으면 ''."""
    for l in str(txt or "").splitlines():
        m = re.match(r"^\s*다음\s*[:：]\s*(.*)$", l)
        if m and m.group(1).strip():
            return m.group(1).strip()
    return ""


def weeklog(txt, week):
    """비고 셀에서 이번 주차 로그 줄만 → 본문. 없으면 '' (= 이번 주 안 움직인 행)."""
    for l in str(txt or "").splitlines():
        l = l.strip()
        if BOX.match(l):
            continue
        m = re.match(r"^([^:：]{2,12})[:：]\s*(.*)$", l)
        if m and norm_week(m.group(1)) == week:
            return m.group(2).strip()
    return ""


def history(txt, week=""):
    """비고 = 체크리스트 + 주간 히스토리 한 칸.

    [x]/[ ] 줄은 하위 항목이라 전부 보인다(접으면 진행률의 근거가 안 보인다).
    그 외 줄은 주차 로그라 최신 한 줄만 두고 접는다.
    """
    if not txt:
        return ""
    lines = [l.strip() for l in str(txt).splitlines() if l.strip()]
    # 이슈·다음 줄은 각자 §이슈·§다음 주 계획 절에서 낸다 — 여기 두면 최신 로그를 가린다
    lines = [l for l in lines if not re.match(r"^(이슈|다음)\s*[:：]", l)]
    box, log = [l for l in lines if BOX.match(l)], [l for l in lines if not BOX.match(l)]
    # 접히지 않고 보이는 줄은 하나뿐이라 그 자리는 이번 주 로그가 갖는다.
    # 맨 위에 접두어 없는 메모('참고: …')를 얹으면 정작 이번 주 한 일이 <details> 로 숨는다.
    if week:    # sort 는 안정 정렬이라 나머지 줄의 순서는 그대로다
        log.sort(key=lambda l: norm_week(re.split(r"[:：]", l, 1)[0]) != week)

    def one(l):
        m = re.match(r"^([^:：]{2,12})[:：]\s*(.*)$", l)
        return (f'<span class="w">{esc(m.group(1))}</span> {esc(m.group(2))}'
                if m else esc(l))

    out = ""
    if box:
        def item(l):
            m = BOX.match(l)
            name, who, week = split_item(m.group(2))
            tail = " ".join(x for x in (f"— {who}" if who else "", week and f"({week})") if x)
            body = esc(name) + (f' <span class="w">{esc(tail)}</span>' if tail else "")
            mark = {"x": "☑", "X": "☑", "o": "▶", "O": "▶"}.get(m.group(1), "☐")
            return f'<div>{mark} {body}</div>'
        out += "".join(item(l) for l in box)
    if log:
        # 이번 주 줄은 몇 줄이든 다 보인다. 한 줄만 펴고 접으면 둘째 줄이 '이전 N주' 로
        # 접혀 지난 주 일로 오라벨된다 — 접히는 건 정말 지난 주 줄만이어야 한다.
        cur = sum(1 for l in log if week and norm_week(re.split(r"[:：]", l, 1)[0]) == week) or 1
        out += "".join(f'<div>{one(l)}</div>' for l in log[:cur])
        if len(log) > cur:
            rest = "".join(f'<div class="old">{one(l)}</div>' for l in log[cur:])
            out += f'<details><summary>이전 {len(log)-cur}주</summary>{rest}</details>'
    return out


TICKET = re.compile(r"^[A-Z][A-Z0-9]+-\d+$")


def jira_cell(v, base=""):
    """티켓 키면 링크로. 'IT-3577, IT-764' 처럼 여러 개면 각각 건다.

    `base` 는 엑셀 `안내·목록` 의 'Jira: …' 줄에서 온다. 없으면 링크 없이 키만 나온다.
    """
    if v in (None, "", "티켓 없음"):
        return ""
    keys = [k.strip() for k in re.split(r"[,\s/]+", str(v)) if k.strip()]
    out = [f'<a href="{esc(base)}{esc(k)}" target="_blank">{esc(k)}</a>'
           if base and TICKET.match(k) else esc(k) for k in keys]
    return " ".join(out)


COLS = ("3%", "3%", "auto", "7%", "13%", "8%", "9%", "21%", "8%")


def task_table(group, week="", jira_base=""):
    cg = "".join(f'<col style="width:{w}">' for w in COLS)
    th = ("<tr><th>#</th><th>★</th><th>과제명</th><th>담당</th><th>단계</th><th>상태</th><th>Due</th>"
          "<th>비고 (주간 로그)</th><th>Jira</th></tr>")
    tr = "".join(
        f'<tr><td class="c n">{d["no"]}</td>'      # n = nowrap. 두 자리(10~)가 세로로 쪼개지지 않게
        f'<td class="c n">{"★" if d["star"] else ""}</td>'
        f'<td>{esc(d["name"])}</td>'
        f'<td class="c">{esc(d["owner"])}</td>'
        f'<td>{history(d["steps"])}</td>'
        f'<td class="n">{status(d["status"], d["sig"], d["steps"])}</td>'
        f'<td class="n c">{esc(d["due"])}</td>'
        f'<td>{history(d["note"], week)}</td>'
        f'<td class="n">{jira_cell(d["jira"], jira_base)}</td></tr>'
        for d in group)
    return f'<div class="tw"><table><colgroup>{cg}</colgroup><thead>{th}</thead><tbody>{tr}</tbody></table></div>'


def plain_table(headers, rows, widths=None):
    cg = ("<colgroup>" + "".join(f'<col style="width:{w}">' for w in widths) + "</colgroup>") if widths else ""
    th = "".join(f"<th>{h}</th>" for h in headers)
    tr = "".join("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in row) + "</tr>" for row in rows)
    return f'<div class="tw"><table>{cg}<thead><tr>{th}</tr></thead><tbody>{tr}</tbody></table></div>'


def heading(v):
    """1행 = 발표본 제목. 공백만 있는 칸도 빈 칸으로 본다."""
    return str(v or "").strip() or "주간업무 보고"


def lint(rows, week, meta, domains=(), jira_base="", jira_raw=""):
    """규격 밖 값 목록. 전부 «에러» 가 아니라 «조용히 빠짐» 으로 나타나는 것들이다.

    print 가 아니라 리스트로 돌려준다 — 경고 자체가 사라지는 변이를 selftest 가 잡을 수 있어야 한다.
    """
    w = []
    if not week:
        w.append("⚠ 주차 미인식: 진행현황 2행에 '2026년 9월 3주차' 형태가 없다"
                 " → §완료·§진행 중·§다음 주 계획이 통째로 빈다")
    elif not any(d["log"] for d in rows):
        # 주차는 읽혔는데 아무도 그 라벨을 안 썼을 때도 결과는 똑같이 '세 절이 빔' 이다.
        # 리더가 2행만 다음 주로 고치고 뽑으면 바로 이 상태 — 요약줄의 '움직인 행 0' 은 눈에 안 띈다.
        w.append(f"⚠ 이번 주({week}) 로그 0건: 비고에 '{week}: …' 를 쓴 행이 하나도 없다"
                 " → §완료·§진행 중·§다음 주 계획이 통째로 빈다")
    for i, v in enumerate(meta, start=2):
        if not str(v or "").strip():
            w.append(f"⚠ 머리말 {i}행이 비었다 — 그 줄 항목이 발표본 머리말에서 빠진다")
    # 설정이 «없는» 것도 조용한 실패다. 없으면 기능이 빠진 채로 멀쩡히 돌기 때문에 여기서 말한다.
    if not domains:
        w.append(f"⚠ 도메인 목록 없음: '{GUIDE}' 시트에 「도메인」 칸이 없다(그 아래가 목록이다)"
                 " → 도메인 오타를 검출할 근거가 없다(그대로 표가 갈라진다)")
    if not jira_base:
        # 「줄이 없다」와 「줄은 있는데 주소가 틀렸다」는 처방이 다르다 — 뭉뚱그리면 시키는 대로
        # 고쳐도 같은 경고가 또 뜬다(순환). raw 를 받아 둘을 가른다.
        if str(jira_raw or "").strip():
            w.append(f"⚠ Jira 주소 무시: '{jira_raw}' — http:// 또는 https:// 로 시작해야 한다"
                     " → 티켓 키가 링크 없이 글자로만 나온다")
        else:
            w.append(f"⚠ Jira 주소 없음: '{GUIDE}' 시트에 「Jira 주소」 칸을 만들고 **그 오른쪽 칸**에"
                     " 'https://<회사>.atlassian.net/browse/' → 티켓 키가 링크 없이 글자로만 나온다")
    for d in rows:
        # 담당 칸(필터용 1명)과 단계 첫 [ ] 의 담당이 어긋나면 알린다 — 단계 넘길 때 같이 안 바꾼 것
        cur = d["doing"] or d["todo"]
        want = cur[0][1] if cur else ""
        if want and want not in str(d["owner"] or ""):     # 담당은 'A / B' 도 된다 — 포함 여부로 본다
            w.append(f"⚠ 담당 불일치: {d['name']} — 담당칸 '{d['owner'] or ''}' vs 단계 '{want}'")
        if domains and d["domain_raw"] not in domains:     # 목록 밖·빈 칸은 '기타' 로 접어 냈다
            w.append(f"⚠ 도메인 밖: {d['name']} — '{d['domain_raw'] or '(빈칸)'}'"
                     f" (허용 {'/'.join(domains)}) → '{OTHER}' 표로 접어 냈다")
        # 상태가 '완료' 인데 안 끝난 단계가 남아 있으면 그 행은 §완료·§진행 중에 동시에 실린다.
        # 둘 중 뭐가 참인지 코드가 정할 수 없다 — 사람이 정해야 한다.
        if str(d["status"] or "").strip() == "완료" and (d["doing"] or d["todo"]):
            # 로그가 없으면 §완료·§진행 중 둘 다 못 들어간다 — 「같이 나온다」는 거짓이 된다
            where = " → §완료·§진행 중에 같이 나온다" if d["log"] else " → 어느 절에도 안 나온다"
            w.append(f"⚠ 완료인데 단계가 남았다: {d['name']} —"
                     f" 안 끝난 단계 {len(d['doing']) + len(d['todo'])}개"
                     + where + ". 단계를 [x] 로 바꾸거나 상태를 되돌려라")
        if str(d["status"] or "").strip() not in DOT:      # 상태가 틀리면 색·완료 판정이 죽는다
            w.append(f"⚠ 상태 밖: {d['name']} — '{d['status'] or ''}'"
                     f" (허용 {'/'.join(DOT)}) → 색·§완료 판정 안 된다")
        if d["due"] and not isinstance(d["due"], (date, datetime)) \
                and not re.match(r"^\d{4}-\d{2}-\d{2}", str(d["due"]).strip()):
            w.append(f"⚠ Due 형식: {d['name']} — '{d['due']}' → 기한 지나도 🔴 가 안 붙는다")
    return w


def main(path):
    wb = load_workbook(path)
    ws = wb["진행현황"]
    cfg = sheet_config(wb)                       # 팀마다 다른 값은 코드가 아니라 엑셀에 있다
    jira_raw = cfg.get("Jira", "")
    jira_base, declared = safe_url(jira_raw), cfg.get("도메인", [])

    rows = []
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        v = cells(ws, r, 9)
        if not v[2] or str(v[2]).strip().startswith(EXAMPLE):
            continue
        dom, dom_raw = fold_domain(v[1], declared)
        rows.append(dict(domain=dom, domain_raw=dom_raw, star=bool(v[0]), name=v[2],
                         owner=v[3], steps=v[4], status=v[5], due=v[6],
                         note=v[7], jira=v[8]))
    # 보고 주차 라벨 — 2행 '주차: 2026년 9월 2주차 …' 에서 뽑아 '09/2주' 로. 리더가 매주 그 셀만 고친다.
    week = week_label(ws.cell(row=2, column=1).value)

    for d in rows:
        d["sig"] = signal(d["status"], d["due"])
        d["done"], d["doing"], d["todo"] = checklist(d["steps"])
        d["log"] = weeklog(d["note"], week) if week else ""

    domains = domain_order(rows, declared)

    # # 은 엑셀에 두지 않는다(행 추가·삭제 때마다 손으로 다시 매겨야 함) — 여기서 도메인별로 자동 부여
    for dom in domains:
        for i, d in enumerate([x for x in rows if x["domain"] == dom], start=1):
            d["no"] = i

    B = []          # 본문
    n = 0           # 섹션 번호

    # ── 1. 도메인별 진행 현황
    n += 1
    B.append(f"<h2>{n}. 도메인별 진행 현황</h2>")
    stars = [d for d in rows if d["star"]]
    if stars:
        B.append(f"<h3>{n}.0 주요 과제 ★</h3>")
        B.append(f'<div class="star">{task_table(stars, week, jira_base)}</div>')
    for i, dom in enumerate(domains, start=1):
        group = [d for d in rows if d["domain"] == dom]
        B.append(domain_heading(n, i, dom))
        # 과제가 없어도 절은 낸다 — 도메인이 통째로 사라지면 「빠뜨린 건지 없는 건지」를 못 읽는다
        if not group:
            B.append('<p class="none">이번 주 진행 건 없음.</p>')
        # 이름 전체가 괄호면 자리표시(「(아직 없음)」) — 이름 앞에만 괄호가 붙은 실제 과제
        # (「(사내) 근태 모듈」 같은 이름)를 회색 문장으로 삼키지 않도록 fullmatch 로 본다
        elif len(group) == 1 and re.fullmatch(r"\(.*\)", str(group[0]["name"]).strip()):
            B.append(f'<p class="none">{esc(str(group[0]["name"]).strip()[1:-1])}.</p>')
        else:
            B.append(task_table(group, week, jira_base))

    # ── 2~4. 이번 주 완료 / 진행 중 / 다음 주 계획 — 전부 '진행현황' 체크리스트에서 파생.
    # 따로 적는 시트를 두지 않는다 — 같은 걸 두 군데 쓰면 갈라진다.
    moved = [d for d in rows if d["log"]]          # 이번 주 로그가 있는 행 = 움직인 행
    # 각 원소 = (행, 문구, 문구에 담당이 이미 붙었나). 단계를 쪼갠 항목만 담당을 달고 온다.
    sections = [
        ("이번 주 완료 (Done)",
         [(d, f'{i[0]}{" — " + i[1] if i[1] else ""}', bool(i[1]))
          for d in rows for i in d["done"] if i[2] == week]
         # 이번 주 체크 항목이 없는 과제(단계를 안 쪼갰거나 지난 주에 다 끝냈거나)
         # → 상태 '완료' + 이번 주 로그로 잡는다
         + [(d, d["log"], False) for d in rows if done_fallback(d, week)]),
        ("진행 중 (In Progress)",
         [(d, f'{(d["doing"] or d["todo"])[0][0]}'
              f'{" — " + (d["doing"] or d["todo"])[0][1] if (d["doing"] or d["todo"])[0][1] else ""}'
              f'{" / " + d["log"] if d["log"] else ""}', bool((d["doing"] or d["todo"])[0][1]))
          for d in moved if d["doing"] or d["todo"]]
         + [(d, d["log"], False) for d in moved if progress_fallback(d)]),
        # 진행 중인 일은 다음 주에도 계속하므로 §진행 중에 이미 있다 — 여기는 '새로 시작하는 것' 만.
        # 소스 = 비고의 '다음: ' 줄(사람이 정하는 결정이라 자동으로 뽑을 데가 없다).
        ("다음 주 계획 (Plan)",
         # moved 로 제한 = 이번 주에 움직인 행만. 갱신을 놓친 행의 지난주 '다음:' 줄이
         # 이번 주 계획인 척하는 것을 막는다(로그가 신선도 보증 역할).
         # 체크리스트에서 «다음 단계» 를 뽑지 않는다 — 목록 순서가 곧 일정 순서는 아니라
         # 엉뚱한 단계가 계획으로 올라간다. 계획은 사람이 정하는 결정이다.
         [(d, nextline(d["note"]), False) for d in moved if nextline(d["note"])]),
    ]
    for title, items in sections:
        n += 1
        B.append(f"<h2>{n}. {title}</h2>")
        if not items:
            B.append('<p class="none">없음.</p>')
            continue
        # 같은 과제의 여러 항목은 한 줄로 묶는다 — 과제명이 반복되면 읽는 사람이 셋을 못 센다
        merged = {}
        for d, t, owned in items:
            e = merged.setdefault(id(d), (d, [], []))
            e[1].append(t); e[2].append(owned)
        # §1 표와 같은 순서로 — 갈래(체크리스트/로그)별로 모으면 도메인 순서가 뒤집힌다
        merged = dict(sorted(merged.items(), key=lambda kv: rows.index(kv[1][0])))
        B.append("<ul>" + "".join(
            f'<li><input type="checkbox" disabled> ({esc(d["domain"])}) '
            f'<strong>{esc(d["name"])}</strong>'
            # 항목 문구가 과제명을 되풀이하면(단계 없는 과제의 로그) 과제명만 남긴다
            + ("" if len(ts) == 1 and str(ts[0]).strip() in str(d["name"])
               else " · " + esc(" / ".join(ts)))
            # 단계를 쪼갠 항목은 이미 '— 담당' 을 달고 온다. 안 쪼갠 과제는 담당 열에서 붙인다.
            + (f' <span class="w">— {esc(d["owner"])}</span>'
               if d["owner"] and not any(owned) else "")
            + "</li>" for d, ts, owned in merged.values()) + "</ul>")

    # ── 5. 배포 이력
    n += 1
    w3 = wb["배포이력"]
    dep = [v for v in (cells(w3, r, 7) for r in range(5, w3.max_row + 1)) if any(x for x in v)]
    B.append(f"<h2>{n}. 배포 이력</h2>")
    B.append(plain_table(["일자", "대상 시스템", "버전 / 브랜치", "담당", "결과", "결함", "비고"], dep)
             if dep else '<p class="none">이번 주 배포 없음.</p>')

    # ── 6. 이슈 — 비고 칸의 '이슈: ' 줄에서 뽑는다. 따로 적는 시트를 두지 않는다.
    n += 1
    if "이슈" in wb.sheetnames:                       # 이슈 시트가 있으면 그쪽이 원본
        wi = wb["이슈"]
        iss = [v[:7] for v in (cells(wi, r, 8) for r in range(5, wi.max_row + 1)) if any(x for x in v)]
        head = ["구분", "관련 과제", "내용", "영향", "대응 · 요청", "요청 대상", "기한"]
        wid = ["9%", "18%", "auto", "15%", "20%", "10%", "9%"]
    else:                                            # 없으면 비고의 '이슈: ' 줄에서
        iss, head, wid = [], ["도메인", "관련 과제", "내용", "요청 대상"], ["9%", "22%", "auto", "12%"]
        for d in rows:
            for l in str(d["note"] or "").splitlines():
                m = re.match(r"^\s*이슈\s*[:：]\s*(.*)$", l)
                if m:
                    p = m.group(1).strip().rsplit(" — ", 1)
                    iss.append([d["domain"], d["name"], p[0], p[1] if len(p) > 1 else ""])
    B.append(f"<h2>{n}. 이슈 / 리스크 / 협업 요청 / 팀 운영 / 특이사항</h2>")
    B.append(plain_table(head, iss, wid) if iss else '<p class="none">없음.</p>')

    # ── 7. 월별 KPI — 맨 아래. 비어 있으면 절 자체를 안 낸다 — 빈 절은 발표본에서 잡음이다.
    w5 = wb["KPI(월)"]
    kpi = [v for v in (cells(w5, r, 7) for r in range(5, w5.max_row + 1)) if any(x for x in v[1:])]
    if kpi:
        n += 1
        B.append(f"<h2>{n}. 월별 KPI</h2>")
        B.append(plain_table(["월", "배포 (Prod)", "배포 (Staging)", "납기 준수율",
                              "배포 후 결함", "진행 과제 수", "비고"], kpi))

    # 머리말도 엑셀 2·3행에서 뽑는다 — 여기 박아 두면 다음 주에 지난 주 날짜가 그대로 나간다.
    # 두 행 모두 '·' 로 끊은 사람 읽는 문장이라 그 구분자대로 항목을 낸다.
    meta_html = "".join(f"<li>{m}</li>"
                        for row in (2, 3)
                        for m in meta_items(ws.cell(row=row, column=1).value))
    # 제목은 1행에서. 파일명에서 뽑으면 사본 접미사('…_v1 1')가 인쇄물 머리에 박힌다.
    h1 = heading(ws.cell(row=1, column=1).value)
    title = f"{h1} {week}".strip()

    doc = f"""<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
<title>{esc(title)}</title>
<style>body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans KR",sans-serif;max-width:1080px;margin:24px auto;padding:0 24px;color:#1f2328;line-height:1.6;font-size:14px;}}
h1{{font-size:22px;border-bottom:2px solid #d0d7de;padding-bottom:6px;}}
h2{{font-size:17px;border-bottom:1px solid #d0d7de;padding-bottom:4px;margin-top:26px;}}
h3{{font-size:14px;color:#0969da;margin-top:16px;}}
hr{{border:0;border-top:1px solid #d0d7de;margin:16px 0;}}
.tw{{overflow-x:auto;}}
table{{border-collapse:collapse;width:100%;min-width:840px;margin:8px 0;font-size:13px;table-layout:fixed;}}
th,td{{border:1px solid #d0d7de;padding:6px 10px;text-align:left;vertical-align:top;word-break:break-word;}}
th{{background:#f6f8fa;}}
td.c{{text-align:center;}}
td.n,th.n{{white-space:nowrap;}}
ul{{padding-left:20px;}} li{{margin:2px 0;}}
code{{background:#f6f8fa;padding:1px 4px;border-radius:3px;font-family:Consolas,monospace;font-size:12px;}}
a{{color:#0969da;}}
.d{{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:5px;vertical-align:middle;background:#8c959f;}}
.d.done{{background:#2da44e;}} .d.go{{background:#0969da;}} .d.warn{{background:#d4a72c;}} .d.bad{{background:#cf222e;}}
.w{{color:#6a737d;}}
.old{{color:#6a737d;}}
details summary{{cursor:pointer;color:#6a737d;font-size:12px;}}
.star table{{background:#fffdf5;}}
.none{{color:#6a737d;}}
@media print{{body{{margin:0;}} details{{display:block;}} details summary{{display:none;}} table{{page-break-inside:auto;}} tr{{page-break-inside:avoid;}}}}
</style></head><body><h1>{esc(h1)}</h1>
<ul>{meta_html}</ul>
<hr>
{''.join(B)}
<hr>
<p class="none" style="font-size:12px">원본: <code>{esc(os.path.basename(path))}</code> — 이 HTML 은 엑셀에서 자동 생성됩니다. 직접 수정하지 마세요.</p>
</body></html>"""

    out = os.path.splitext(path)[0] + ".html"
    with open(out, "w", encoding="utf-8") as fp:
        fp.write(doc)
    for w in lint(rows, week, [ws.cell(row=r, column=1).value for r in (2, 3)],
                  declared, jira_base, jira_raw):
        print(" ", w)

    print("saved:", out,
          f"| 주차 {week or '?'} · 과제 {len(rows)} · 움직인 행 {len(moved)} "
          f"· 배포 {len(dep)} · 이슈 {len(iss)}")


def endtoend():
    """진짜 엑셀을 만들어 `main()` 을 통째로 돌린다.

    🔑 **헬퍼만 단언하면 호출부가 비어 있는 걸 못 본다.** 독립 QA 가 두 라운드 연속 같은 계열을
    잡았다 — 헬퍼는 맞는데 `main()` 에서 그 헬퍼를 안 쓰도록 되돌리면 검사가 통과했다.
    그래서 여기서는 산출물 HTML 을 보고 판정한다.
    """
    import contextlib, tempfile
    from openpyxl import Workbook

    def run(path):                       # main() 의 진단 출력은 여기선 잡음이다
        with contextlib.redirect_stdout(io.StringIO()):
            main(path)
        return io.open(os.path.splitext(path)[0] + ".html", encoding="utf-8").read()

    def build(path, jira, dom, declared=("도메인A",), example=True,
              steps="[o] 개발 — 개발자", status="진행중", note="09/2주: 한 일\n다음: 다음 것"):
        wb = Workbook()
        g = wb.active
        g.title = GUIDE
        g["A1"] = "Jira 주소"                      # 라벨 · 값은 오른쪽 칸
        g["B1"] = jira
        g["D1"] = "도메인"                          # 드롭다운 원본 = 헤더 아래로
        for i, d in enumerate(declared, start=2):
            g.cell(row=i, column=4).value = d
        ws = wb.create_sheet("진행현황")
        ws["A1"] = "주간업무 보고"
        ws["A2"] = "주차: 2026년 9월 2주차 · 작성일: 2026-09-10"
        ws["A3"] = "리더 아무개"
        # 예시 행이 있으면 첫 과제는 그 아래, 지웠으면 **헤더 바로 다음 줄**로 올라온다.
        # 그 자리를 안 써 보면 「데이터 시작행을 위치로 박는」 회귀를 못 잡는다.
        r0 = HEADER_ROW + 1
        if example:
            ws.cell(row=r0, column=3).value = f"{EXAMPLE} 예시 과제"
            r0 += 1
        for c, v in enumerate(["★", dom, "과제 하나", "개발자", steps,
                               status, "2026-11-30", note, "IT-1"], 1):
            ws.cell(row=r0, column=c).value = v
        for s in ("배포이력", "이슈", "KPI(월)"):
            wb.create_sheet(s)
        wb.save(path)

    tmp = tempfile.mkdtemp(prefix="weekly_report_selftest_")
    try:
        # ① 도메인 칸으로 절 구조를 위조할 수 있나 (호출부가 esc 를 안 타면 여기서 깨진다)
        # 도메인 선언이 없으면 접히지 않고 엑셀 값이 그대로 절 제목이 된다 = 실제 공격 경로
        p = os.path.join(tmp, "inj.xlsx")
        build(p, "https://x.example/browse/", "</h3><h2>가짜 절</h2><h3>", declared=())
        out = run(p)
        assert "<h2>가짜 절</h2>" not in out, "도메인 값이 절 구조를 위조했다"
        assert "&lt;/h3&gt;&lt;h2&gt;" in out
        assert '<a href="https://x.example/browse/IT-1"' in out, "Jira 링크 배선이 끊겼다"
        assert EXAMPLE not in out, "예시 행이 발표본에 나왔다"
        assert "예시 과제" not in out, "예시 행 내용이 발표본에 나왔다"   # 겹쳐 쓰면 위 단언이 공허해진다
        assert "과제 하나" in out, "예시 행 판정이 진짜 과제를 삼켰다"
        assert "▶" in out, "[o] 진행 중 표시가 안 나왔다"

        # ② 위험한 스킴이 href 에 들어가나 (호출부가 safe_url 을 안 타면 여기서 깨진다)
        p = os.path.join(tmp, "js.xlsx")
        build(p, "javascript:alert(1)//", "도메인A")
        out = run(p)
        assert "javascript:" not in out, "javascript: 스킴이 href 로 나갔다"

        # ③ ★ 행의 도메인이 목록 밖이어도 산출물이 나온다 (자동번호 배선)
        p = os.path.join(tmp, "star.xlsx")
        build(p, "https://x.example/browse/", "목록에없는도메인")
        assert len(run(p)) > 0

        def sec(out, title):
            """절 제목 다음 <h2> 전까지 — 「어느 절에 실렸나」를 산출물로 판정한다.

            🔴 이 함수가 문서 전체를 돌려주면 아래 «있다» 단언이 전부 공허해진다(과제명은 §1 표에
            늘 있다). 그러면 이 함수를 망가뜨리는 변이 하나로 절 배정 검사가 통째로 죽는다 —
            그래서 호출부마다 «없다» 단언을 짝으로 둔다.
            """
            i = out.find(title)
            j = out.find("<h2>", i) if i >= 0 else -1
            return out[i:j if j > 0 else len(out)] if i >= 0 else ""

        # ④ 단계가 `[x]`+`[o]` 뿐(남은 `[ ]` 없음)이어도 §진행 중에 실린다.
        #    호출부에서 `doing` 을 빼면 로그를 쓴 행이 어느 절에도 안 나온다.
        p = os.path.join(tmp, "doing.xlsx")
        build(p, "https://x.example/browse/", "도메인A",
              steps="[x] 기획 — 기획자 (09/1주)\n[o] 개발 — 개발자")
        out = run(p)
        assert "과제 하나" in sec(out, "진행 중 (In Progress)"), "[o] 행이 §진행 중에서 빠졌다"
        # «없다» 를 짝으로 — sec() 가 문서 전체를 돌려주면 위 단언이 공허해진다
        assert "과제 하나" not in sec(out, "배포 이력"), "sec() 가 절을 안 가른다"
        assert "1/2" in out, f"진행률 분모가 doing 을 안 센다"          # [x]+[o] = 1/2
        assert "다음 것" in sec(out, "다음 주 계획"), "'다음:' 줄이 §계획에 안 실렸다"

        # ⑤ 지난 주에 단계를 다 끝낸 «완료» 행이 §이번 주 완료에 실린다 (done_fallback 배선)
        p = os.path.join(tmp, "donefb.xlsx")
        build(p, "https://x.example/browse/", "도메인A",
              steps="[x] 개발 — 개발자 (09/1주)", status="완료", note="09/2주: 릴리즈 끝")
        out = run(p)
        assert "과제 하나" in sec(out, "이번 주 완료"), "done_fallback 배선이 끊겼다"
        assert "과제 하나" not in sec(out, "배포 이력"), "sec() 가 절을 안 가른다"

        # ⑥ 이번 주에 안 움직인 행의 낡은 '다음:' 줄이 이번 주 계획인 척하지 않는다
        p = os.path.join(tmp, "stale.xlsx")
        build(p, "https://x.example/browse/", "도메인A", note="08/4주: 옛날 일\n다음: 낡은 계획")
        out = run(p)
        assert "낡은 계획" not in sec(out, "다음 주 계획"), "moved 신선도 가드가 풀렸다"

        # ⑦ 예시 행을 지운 워크북에서도 첫 과제가 살아 있다 (위치 고정으로 되돌리면 깨진다)
        p = os.path.join(tmp, "noex.xlsx")
        build(p, "https://x.example/browse/", "도메인A", example=False)
        assert "과제 하나" in run(p), "예시 행이 없으면 첫 과제가 사라진다"

        # ⑧ href 속성·상태 열도 엑셀 값이다 — 속성 탈출·태그 주입이 막혀야 한다
        p = os.path.join(tmp, "attr.xlsx")
        build(p, 'https://x.example/"onmouseover="alert(1)',
              "도메인A", steps="[o] 개발 — <img src=x onerror=alert(1)>")
        out = run(p)
        assert '"onmouseover=' not in out, "Jira 주소가 href 속성을 탈출했다"
        assert "<img src=x onerror" not in out, "담당 이름이 태그로 들어갔다"

        # ⑨ Windows 기본 콘솔(cp949)에서 경고의 한글·기호가 죽지 않는다.
        #    이 자리에서만 잡힌다 — 검사 프로세스가 이미 utf-8 이면 단언으로는 안 보인다.
        import subprocess
        env = dict(os.environ, PYTHONIOENCODING="cp949")
        r = subprocess.run([sys.executable, os.path.abspath(__file__), p],
                           capture_output=True, env=env)
        assert r.returncode == 0, ("cp949 콘솔에서 죽는다 — sys.stdout.reconfigure 가 빠졌다: "
                                   + r.stderr.decode("utf-8", "replace").strip()[-200:])
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def selftest():
    """엑셀 없이 도는 파서 점검. 비고 칸 규격을 바꾸면 여기가 먼저 깨져야 한다."""
    DOM = ["도메인A", "도메인B", "기타"]                    # 실제 값은 엑셀에서 온다
    JB = "https://example.invalid/browse/"
    assert split_item("개발 — 개발자 (09/3주)") == ("개발", "개발자", "09/3주")
    assert split_item("PRD 작성") == ("PRD 작성", "", "")
    done, doing, todo = checklist("[x] 기획 — 기획자 (09/2주)\n[o] 개발 — 개발자\n[ ] QA — QA담당\n09/2주: 진행")
    assert len(done) == 1 and len(doing) == 1 and len(todo) == 1 and done[0][2] == "09/2주"
    assert doing[0][1] == "개발자"
    # 「지금 이 단계」가 있으면 그쪽이 현재 담당 — [ ] 첫 줄로만 보면 아직 시작도 안 한 사람이 잡힌다
    assert "개발자" in status("진행중", "go", "[x] 기획 — 기획자\n[o] 개발 — 개발자\n[ ] QA — QA담당")
    assert "▶" in history("[o] 개발 — 개발자")          # 진행 중 표시가 완료와 구별된다
    assert "☑" in history("[x] 개발 — 개발자")
    assert "☐" in history("[ ] 개발 — 개발자")
    assert nextline("다음: 화면설계 착수\n09/2주: x") == "화면설계 착수"
    assert weeklog("09/2주: 한 일\n09/1주: 지난 주", "09/2주") == "한 일"
    assert weeklog("09/1주: 지난 주", "09/2주") == ""          # 안 움직인 행
    assert signal("진행중", "2020-01-01") == "bad"             # 기한 지남이 다 이긴다
    assert signal("완료", "2020-01-01") == "done"

    # 한 자리 월도 두 자리와 같은 주차다 — 이 등식이 깨지면 그 행이 전 절에서 증발한다
    assert norm_week("9/2주") == norm_week("09/2주") == "09/2주"
    assert split_item("개발 — 개발자 (9/3주)")[2] == "09/3주"
    assert weeklog("9/2주: 한 일", "09/2주") == "한 일"
    assert week_label("주차: 2026년 9월 2주차 · 작성일: 2026-09-10") == "09/2주"
    assert week_label("2026년 10월 1주차") == "10/1주"
    assert week_label("주차: 미정") == ""

    # 규격 밖·빈 칸 도메인은 '기타' 로 접되 원본을 남긴다(접은 사실을 경고로 내야 하므로)
    assert fold_domain("도메인A", DOM) == ("도메인A", "도메인A")
    assert fold_domain("도메인a", DOM) == ("기타", "도메인a")
    assert fold_domain(None, DOM) == ("기타", "")
    assert fold_domain("아무거나", []) == ("아무거나", "아무거나")   # 목록 없으면 접지 않는다

    # 머리말은 2·3행에서 파생된다 — 여기가 죽으면 다음 주에 지난 주 날짜가 나간다
    assert list(meta_items("주차: 2026년 9월 2주차 · 리더 아무개")) == [
        "<strong>주차</strong>: 2026년 9월 2주차", "리더 아무개"]
    assert list(meta_items("시작: 10:30")) == ["<strong>시작</strong>: 10:30"]
    assert list(meta_items("")) == []

    # F2 회귀 방어 — 지난 주에 단계를 다 끝낸 행(todo=[], done 은 지난 주)이 이번 주 로그를
    # 썼으면 어느 절에든 나와야 한다. 두 술어에 `not d["done"]` 이 돌아오면 여기가 깨진다.
    old = dict(done=[("개발", "개발자", "09/1주")], doing=[], todo=[],
               log="릴리즈 대기", status="진행중")
    assert progress_fallback(old) and not done_fallback(old, "09/2주")
    assert done_fallback(dict(old, status="완료"), "09/2주")      # 완료면 §완료 쪽으로
    assert not progress_fallback(dict(old, status="완료"))
    assert not done_fallback(dict(old, done=[("개발", "개발자", "09/2주")]), "09/2주")  # 이번 주 [x] 는 항목이 낸다
    assert not done_fallback(dict(old, log="", status="완료"), "09/2주")              # 로그 없으면 안 움직인 행
    assert not progress_fallback(dict(old, todo=[("QA", "QA담당", "")]))              # 남은 [ ] 는 1절이 낸다
    assert not progress_fallback(dict(old, doing=[("개발", "개발자", "")]))           # [o] 도 1절이 낸다

    # 이번 주 로그는 접히지 않는 첫 줄 자리를 갖는다
    assert heading(None) == heading("   ") == "주간업무 보고"   # 빈 <h1> 금지
    assert heading(" 주간업무 보고 ") == "주간업무 보고"

    # 팀 고유값은 워크북의 «이미 있는» 칸에서 읽는다 — 코드로 되돌아가면 여기가 깨진다
    class _Cell:                                   # openpyxl 없이 시트 읽기를 확인한다
        def __init__(s, v, row, col): s.value, s.row, s.column = v, row, col

    class _Sheet:
        """행마다 (1열, 2열, 3열…) 튜플. 실제 시트처럼 여러 열을 준다 —
        1열만 주면 오른쪽 칸·아래 칸 탐색이 한 번도 검사되지 않는다."""
        def __init__(s, grid): s.g, s.max_row = grid, len(grid)
        def _row(s, r):
            v = s.g[r - 1] if 1 <= r <= len(s.g) else ()
            return (v,) if isinstance(v, str) else v
        def cell(s, row, column):
            r = s._row(row)
            return _Cell(r[column - 1] if column - 1 < len(r) else None, row, column)
        def iter_rows(s):
            return [[s.cell(r, c) for c in range(1, max(len(s._row(r)), 1) + 1)]
                    for r in range(1, s.max_row + 1)]

    class _WB:
        def __init__(s, grid, name=GUIDE): s.sheetnames, s._s = [name], _Sheet(grid)
        def __getitem__(s, k): return s._s

    # 드롭다운 원본 = 헤더 칸 아래로 값이 이어진다. 빈 칸에서 끊는다.
    wb_ = _WB([("도메인", "Jira 주소", "https://x.example/browse/"),
               ("도메인A",), ("도메인B",), (None,), ("안 읽힘",)])
    cfg = sheet_config(wb_)
    assert cfg["도메인"] == ["도메인A", "도메인B", "기타"]   # 접을 칸은 없으면 붙인다
    assert cfg["Jira"] == "https://x.example/browse/"      # 라벨의 «오른쪽» 칸
    assert sheet_config(_WB([("메모",), ("안내",)])) == {}
    assert sheet_config(_WB([], name="다른시트")) == {}
    assert lookup_column(_WB([("도메인",), ("기타",)]), "도메인") == ["기타"]
    assert lookup_value(_WB([("Jira 주소",)]), "Jira 주소") == ""    # 오른쪽이 비면 빈 값
    # 헤더 바로 아래 한 줄이 비어도 목록은 살아 있다 — 첫 칸에서 끊으면 「칸이 없다」로 오진한다
    assert lookup_column(_WB([("도메인",), (None,), ("A",), ("B",)]), "도메인") == ["A", "B"]
    assert lookup_column(_WB([("도메인",), (None,), (None,), ("A",)]), "도메인") == ["A"]   # 두 줄도
    assert lookup_column(_WB([("도메인",), ("A",), (None,), ("안 읽힘",)]), "도메인") == ["A"]  # 값 뒤 빈 칸이 끝
    # 라벨로 «시작하는» 산문 칸이 위에 있어도 완전일치 칸이 이긴다
    assert lookup_value(_WB([("Jira 주소 안내", "이 값 아님"), ("Jira 주소", "https://x.example/b/")]),
                        "Jira 주소") == "https://x.example/b/"
    # 헤더가 여러 번 나오면 «값이 있는» 쪽이 이긴다 (빈 헤더 때문에 목록을 잃지 않는다)
    assert lookup_column(_WB([("도메인",), (None,), (None,), ("도메인",), ("A",)]), "도메인") == ["A"]
    # 완전일치가 없으면 첫 non-empty loose 가 이긴다 (뒤엣것이 앞엣것을 덮지 않는다)
    assert lookup_value(_WB([("Jira 주소 안내", ""), ("Jira 주소 참고", "앞값"),
                             ("Jira 주소 비고", "뒷값")]), "Jira 주소") == "앞값"

    # href 에 들어가는 값은 스킴을 검사한다 — esc() 는 속성 탈출만 막는다
    assert safe_url("https://x.example/browse/") == "https://x.example/browse/"
    assert safe_url("javascript:alert(1)//") == ""
    assert safe_url("data:text/html,<script>") == ""
    assert safe_url(None) == ""

    # 엑셀에서 온 값이 HTML 구조로 새지 않는다 — 도메인 칸 하나로 절을 위조할 수 있었다
    assert domain_heading(1, 2, "도메인A") == "<h3>1.2 도메인A</h3>"
    assert domain_heading(1, 2, "</h3><h2>가짜</h2><h3>") == \
        "<h3>1.2 &lt;/h3&gt;&lt;h2&gt;가짜&lt;/h2&gt;&lt;h3&gt;</h3>"
    assert "<img" not in domain_heading(1, 1, "<img src=x onerror=alert(1)>")

    # 도메인 순서 — 빈 리스트면 자동번호가 안 매겨져 ★ 표가 KeyError 로 죽는다
    rr = [dict(domain="X"), dict(domain="Y"), dict(domain="X")]
    assert domain_order(rr, []) == ["X", "Y"]
    assert domain_order(rr, ["Y", "X"]) == ["Y", "X"]
    assert domain_order([], ["A"]) == ["A"]
    assert domain_order(rr, []) != []
    assert jira_cell("IT-3577", JB) == f'<a href="{JB}IT-3577" target="_blank">IT-3577</a>'
    assert jira_cell("IT-3577") == "IT-3577"                  # 주소 없으면 링크 없이 키만
    assert jira_cell("티켓 없음", JB) == ""

    # 경고 자체가 사라지는 것도 조용한 실패다 — 여기가 그걸 잡는다
    ok = dict(name="A", domain_raw="도메인A", status="진행중", due="2026-11-30",
              owner="기획자", todo=[], doing=[], done=[], log="한 일")
    meta = ["주차: 2026년 9월 2주차", "리더 아무개"]
    def L(rows, week="09/2주", meta=meta, domains=DOM, jb=JB):
        return lint(rows, week, meta, domains, jb)

    assert L([ok]) == []                                                  # 정상은 조용하다
    assert any("도메인 밖" in x for x in L([dict(ok, domain_raw="도메인a")]))
    assert any("상태 밖" in x for x in L([dict(ok, status="진행 중")]))
    assert any("Due 형식" in x for x in L([dict(ok, due="11월 말")]))
    assert any("주차 미인식" in x for x in L([ok], week=""))
    assert any("로그 0건" in x for x in L([dict(ok, log="")]))
    assert any("머리말 3행" in x for x in L([ok], meta=["주차: …", ""]))
    assert any("담당 불일치" in x for x in L([dict(ok, todo=[("개발", "개발자", "")])]))
    # 완료인데 단계가 남으면 §완료·§진행 중에 동시 등재된다 — 모순이라 사람이 정해야 한다
    assert any("완료인데 단계가 남았다" in x
               for x in L([dict(ok, status="완료", todo=[("QA", "QA담당", "")])]))
    assert any("완료인데 단계가 남았다" in x
               for x in L([dict(ok, status="완료", doing=[("개발", "개발자", "")])]))
    assert not any("완료인데 단계가 남았다" in x for x in L([dict(ok, status="완료")]))
    # 「지금 이 단계」가 있으면 담당 대조 기준도 그쪽이다 — [ ] 로 보면 엉뚱한 사람과 비교한다
    assert any("담당 불일치" in x
               for x in L([dict(ok, doing=[("개발", "개발자", "")], todo=[("QA", "기획자", "")])]))
    assert not any("담당 불일치" in x
                   for x in L([dict(ok, doing=[("기획", "기획자", "")], todo=[("개발", "개발자", "")])]))
    # 설정이 «없는» 것도 보고돼야 한다 — 없으면 기능이 빠진 채 멀쩡히 돌기 때문
    assert any("도메인 목록 없음" in x for x in L([ok], domains=[]))
    assert any("Jira 주소 없음" in x for x in L([ok], jb=""))
    # 「줄이 없다」와 「줄은 있는데 주소가 틀렸다」는 처방이 다르다 — 뭉뚱그리면 고쳐도 또 뜬다
    bad = lint([ok], "09/2주", meta, DOM, "", "javascript:alert(1)//")
    assert any("Jira 주소 무시" in x for x in bad)
    assert not any("Jira 주소 없음" in x for x in bad)

    h = history("참고: 메모\n09/2주: 이번 주", "09/2주")
    assert h.index("이번 주") < h.index("메모")
    # 이번 주 줄이 2개면 둘 다 펴고, 접히는 개수는 지난 주 줄 수와 같아야 한다
    h = history("09/2주: 둘째\n09/2주: 첫째\n09/1주: 지난 주", "09/2주")
    assert "이전 1주" in h and h.index("첫째") < h.index("<details")
    assert h.index("둘째") < h.index("<details")
    endtoend()          # 헬퍼 단언만으로는 호출부가 비어 있는 걸 못 본다
    print("selftest ok")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("사용:  python xlsx2html.py <주간보고.xlsx>  |  --selftest")
    selftest() if sys.argv[1] == "--selftest" else main(sys.argv[1])
