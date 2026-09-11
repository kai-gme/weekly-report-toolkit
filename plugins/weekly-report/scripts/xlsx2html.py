# -*- coding: utf-8 -*-
"""주간보고 엑셀 → 주간 미팅 발표본 HTML.

사용:  python xlsx2html.py <주간보고.xlsx>
엑셀이 원본(SSOT), HTML 은 매주 여기서 다시 뽑는다. HTML 을 손으로 고치지 말 것.

팀 이름·도메인 목록·Jira 주소는 코드에 두지 않는다 — 전부 워크북에서 읽는다.
흑백 + 얇은 회색 선. 장식 금지.
"""
import sys, os, html, re
from datetime import date, datetime
from openpyxl import load_workbook

# Windows 기본 콘솔은 cp949 라 한글·기호 출력에서 UnicodeEncodeError 로 죽는다.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):        # 파이프로 넘길 때는 조용히 넘어간다
    pass

HEADER_ROW = 6                      # 그 아래가 전부 데이터.
# 예시 행은 «위치» 가 아니라 «과제명이 (예시) 로 시작하나» 로 가른다 —
# 위치로 박으면 예시 행을 지운 워크북에서 첫 과제가 조용히 사라진다.
FALLBACK_DOMAINS = ["기타"]          # '안내·목록' 에 도메인 목록이 없을 때만


def lookup_column(wb, header):
    """'안내·목록' 시트에서 header 를 찾아 그 아래 칸들을 목록으로 돌려준다.

    도메인 목록은 드롭다운 원본이라 이미 워크북 안에 있다 — 코드에 복사하지 않는다.
    """
    if "안내·목록" not in wb.sheetnames:
        return []
    g = wb["안내·목록"]
    for row in g.iter_rows():
        for c in row:
            if str(c.value or "").strip() == header:
                out = []
                for r in range(c.row + 1, g.max_row + 1):
                    v = g.cell(row=r, column=c.column).value
                    if v in (None, ""):
                        break
                    out.append(str(v).strip())
                return out
    return []


def lookup_value(wb, label):
    """'안내·목록' 에서 label 로 시작하는 칸의 오른쪽 값. 없으면 ''."""
    if "안내·목록" not in wb.sheetnames:
        return ""
    g = wb["안내·목록"]
    for row in g.iter_rows():
        for c in row:
            if str(c.value or "").strip().startswith(label):
                v = g.cell(row=c.row, column=c.column + 1).value
                return str(v).strip() if v else ""
    return ""


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
    cur = doing or todo
    who = cur[0][1] if cur else ""
    pct = (f'<div class="w">{len(done)}/{total}'
           f'{" · " + esc(who) if who else ""}</div>') if total else ""
    return f'<span class="d {cls}"></span>' + esc(str(v).strip()) + pct


BOX = re.compile(r"^\[([ xXoO])\]\s*(.*)$", re.M)  # x=완료 · o=지금 이 단계 · 공백=예정
ITEM = re.compile(r"^(.*?)\s+—\s+(.+)$")        # '개발 — 담당자' → 항목 / 담당
WEEK = re.compile(r"\((\d{1,2}/\d주)\)\s*$")     # '… (09/3주)' → 완료 주차


def split_item(txt):
    """'개발 — 담당자 (09/3주)' → (항목, 담당, 주차). 없는 조각은 빈 문자열."""
    t = str(txt).strip()
    m = WEEK.search(t)
    week = m.group(1) if m else ""
    if m:
        t = t[:m.start()].strip()
    m = ITEM.match(t)
    return (m.group(1), m.group(2).strip(), week) if m else (t, "", week)


def checklist(txt):
    """체크 줄을 판다 → (완료, 지금, 예정). 각 원소 = (항목, 담당, 주차).

    [x] 끝남 · [o] 지금 이 단계 · [ ] 아직. 「지금 담당」 은 [o] 가 있으면 그쪽이 먼저다.

    진행률·이번 주 완료·진행 중·다음 계획이 전부 여기서 파생된다(체크리스트가 진행률의 근거).
    같은 걸 '주간' 시트에 또 쓰지 않는다 — 두 번 쓰면 갈라진다.
    """
    hit = [m for m in (BOX.match(l.strip()) for l in str(txt or "").splitlines()) if m]
    done = [split_item(m.group(2)) for m in hit if m.group(1) in "xX"]
    doing = [split_item(m.group(2)) for m in hit if m.group(1) in "oO"]
    todo = [split_item(m.group(2)) for m in hit if m.group(1) == " "]
    return done, doing, todo


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
        if m and m.group(1).strip() == week:
            return m.group(2).strip()
    return ""


def history(txt):
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
        out += f'<div>{one(log[0])}</div>'
        if len(log) > 1:
            rest = "".join(f'<div class="old">{one(l)}</div>' for l in log[1:])
            out += f'<details><summary>이전 {len(log)-1}주</summary>{rest}</details>'
    return out


# Jira 주소는 '안내·목록' 시트의 「Jira 주소」 칸에서 읽는다. 없으면 링크 없이 키만 나온다.
TICKET = re.compile(r"^[A-Z][A-Z0-9]+-\d+$")


def jira_cell(v, base):
    """티켓 키면 링크로. 'ABC-123, ABC-456' 처럼 여러 개면 각각 건다."""
    if v in (None, "", "티켓 없음"):
        return ""
    keys = [k.strip() for k in re.split(r"[,\s/]+", str(v)) if k.strip()]
    out = [f'<a href="{base}{esc(k)}" target="_blank">{esc(k)}</a>'
           if base and TICKET.match(k) else esc(k) for k in keys]
    return " ".join(out)


COLS = ("3%", "3%", "auto", "7%", "13%", "8%", "9%", "21%", "8%")


def task_table(group, jira_base=""):
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
        f'<td>{history(d["note"])}</td>'
        f'<td class="n">{jira_cell(d["jira"], jira_base)}</td></tr>'
        for d in group)
    return f'<div class="tw"><table><colgroup>{cg}</colgroup><thead>{th}</thead><tbody>{tr}</tbody></table></div>'


def plain_table(headers, rows, widths=None):
    cg = ("<colgroup>" + "".join(f'<col style="width:{w}">' for w in widths) + "</colgroup>") if widths else ""
    th = "".join(f"<th>{h}</th>" for h in headers)
    tr = "".join("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in row) + "</tr>" for row in rows)
    return f'<div class="tw"><table>{cg}<thead><tr>{th}</tr></thead><tbody>{tr}</tbody></table></div>'


def main(path):
    wb = load_workbook(path)
    ws = wb["진행현황"]

    rows = []
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        v = cells(ws, r, 9)
        if not v[2] or str(v[2]).strip().startswith("(예시)"):
            continue
        rows.append(dict(domain=v[1] or "기타", star=bool(v[0]), name=v[2],
                         owner=v[3], steps=v[4], status=v[5], due=v[6],
                         note=v[7], jira=v[8]))
    # 보고 주차 라벨 — 2행 '주차: 2026년 9월 2주차 …' 에서 뽑아 '09/2주' 로. 리더가 매주 그 셀만 고친다.
    m = re.search(r"(\d{1,2})월\s*(\d)주차", str(ws.cell(row=2, column=1).value or ""))
    week = f"{int(m.group(1)):02d}/{m.group(2)}주" if m else ""

    # 팀 고유값은 전부 워크북에서. 코드에 이름·도메인·Jira 주소를 두지 않는다.
    domains = lookup_column(wb, "도메인") or FALLBACK_DOMAINS
    if "기타" not in domains:
        domains.append("기타")
    jira_base = lookup_value(wb, "Jira 주소")

    for d in rows:
        d["sig"] = signal(d["status"], d["due"])
        d["done"], d["doing"], d["todo"] = checklist(d["steps"])
        d["log"] = weeklog(d["note"], week)

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
        B.append(f'<div class="star">{task_table(stars, jira_base)}</div>')
    for i, dom in enumerate(domains, start=1):
        group = [d for d in rows if d["domain"] == dom]
        B.append(f"<h3>{n}.{i} {dom}</h3>")
        # 과제가 없어도 절은 낸다 — 도메인이 통째로 사라지면 「빠뜨린 건지 없는 건지」를 못 읽는다
        if not group:
            B.append('<p class="none">이번 주 진행 건 없음.</p>')
        elif len(group) == 1 and str(group[0]["name"]).startswith("("):
            B.append(f'<p class="none">{esc(group[0]["name"]).strip("()")}.</p>')
        else:
            B.append(task_table(group, jira_base))

    # ── 2~4. 이번 주 완료 / 진행 중 / 다음 주 계획 — 전부 '진행현황' 체크리스트에서 파생.
    # 따로 적는 시트를 두지 않는다 — 같은 걸 두 번 쓰면 갈라진다.
    moved = [d for d in rows if d["log"]]          # 이번 주 로그가 있는 행 = 움직인 행
    # 각 원소 = (행, 문구, 문구에 담당이 이미 붙었나). 단계를 쪼갠 항목만 담당을 달고 온다.
    sections = [
        ("이번 주 완료 (Done)",
         [(d, f'{i[0]}{" — " + i[1] if i[1] else ""}', bool(i[1]))
          for d in rows for i in d["done"] if i[2] == week]
         # 단계를 안 쪼갠 과제는 체크 항목이 없다 → 상태 '완료' + 이번 주 로그로 잡는다
         + [(d, d["log"], False) for d in rows
            if not d["done"] and d["log"] and str(d["status"]).strip() == "완료"]),
        ("진행 중 (In Progress)",
         [(d, f'{(d["doing"] or d["todo"])[0][0]}'
              f'{" — " + (d["doing"] or d["todo"])[0][1] if (d["doing"] or d["todo"])[0][1] else ""}'
              f'{" / " + d["log"] if d["log"] else ""}', bool((d["doing"] or d["todo"])[0][1]))
          for d in moved if d["doing"] or d["todo"]]
         # 단계를 안 쪼갠 과제는 로그 줄이 곧 '이번 주 한 일'
         + [(d, d["log"], False) for d in moved
            if not d["todo"] and not d["doing"] and not d["done"]
            and str(d["status"]).strip() != "완료"]),
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

    # ── 7. 월별 KPI. 비어 있으면 절 자체를 안 낸다 — 빈 절은 발표본에서 잡음이다.
    w5 = wb["KPI(월)"]
    kpi = [v for v in (cells(w5, r, 7) for r in range(5, w5.max_row + 1)) if any(x for x in v[1:])]
    if kpi:
        n += 1
        B.append(f"<h2>{n}. 월별 KPI</h2>")
        B.append(plain_table(["월", "배포 (Prod)", "배포 (Staging)", "납기 준수율",
                              "배포 후 결함", "진행 과제 수", "비고"], kpi))

    # 제목·머리말은 '진행현황' 1~3행 그대로. 팀 구성이 바뀌어도 스크립트를 안 고친다.
    title = str(ws.cell(row=1, column=1).value or "주간업무 보고").strip()
    head = "".join(f"<li>{esc(str(ws.cell(row=r, column=1).value).strip())}</li>"
                   for r in (2, 3) if ws.cell(row=r, column=1).value)

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
</style></head><body><h1>{esc(title)}</h1>
<ul>{head}</ul>
<hr>
{''.join(B)}
<hr>
<p class="none" style="font-size:12px">원본: <code>{esc(os.path.basename(path))}</code> — 이 HTML 은 엑셀에서 자동 생성됩니다. 직접 수정하지 마세요.</p>
</body></html>"""

    out = os.path.splitext(path)[0] + ".html"
    with open(out, "w", encoding="utf-8") as fp:
        fp.write(doc)
    # 담당 칸(필터용 1명)과 단계 첫 [ ] 의 담당이 어긋나면 알린다 — 단계 넘길 때 같이 안 바꾼 것
    for d in rows:
        cur = d["doing"] or d["todo"]
        want = cur[0][1] if cur else ""
        if want and want not in str(d["owner"] or ""):     # 담당은 'A / B' 도 된다 — 포함 여부로 본다
            print(f"  ⚠ 담당 불일치: {d['name']} — 담당칸 '{d['owner'] or ''}' vs 단계 '{want}'")

    print("saved:", out,
          f"| 주차 {week or '?'} · 과제 {len(rows)} · 움직인 행 {len(moved)} "
          f"· 배포 {len(dep)} · 이슈 {len(iss)}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("사용:  python xlsx2html.py <주간보고.xlsx>")
    main(sys.argv[1])
