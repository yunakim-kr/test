#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""협의체 안건 도출 에이전트 (연습용)

하는 일
  select : 동향 자료를 검증하고 이슈별 점수를 매겨 '선별 후보'를 보여준 뒤 멈춘다(확인 지점 1).
  build  : 담당자가 확정한 이슈로 안건서 Word 파일을 만든다. 만들기 전에 업무 규칙을 점검하고,
           오류가 있으면 파일을 만들지 않고 멈춘다.

  run    : 위 두 단계를 화면 안내에 따라 이어서 진행한다(더블클릭 실행 파일이 이 명령을 부른다).
  중복 실행 방지: 이미 실행 중이면 두 번째 실행은 안내 메시지를 보이고 끝난다.

하지 않는 일(연습용이라 뺀 것)
  웹 검색, AI 호출, 인터넷 연결, 이메일 발송, 화면(GUI), 예약 실행.
  프로그램은 폴더 안의 자료만 사용하며 자료에 없는 사실은 만들지 않는다.
"""
import argparse
import atexit
import csv
import datetime
import difflib
import json
import os
import re
import shutil
import signal
import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
RATING = {"상": 3, "중": 2, "하": 1}
GUBUN = {"산", "학", "관", "해외"}


# ---------------------------------------------------------------- 공통 도구
def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def parse_date(text):
    m = re.fullmatch(r"(\d{4})-(\d{2})(?:-(\d{2}))?", (text or "").strip())
    if not m:
        return None
    return int(m[1]), int(m[2]), int(m[3]) if m[3] else None


def parse_period(values):
    """기간(시작, 끝) 입력을 날짜로 바꾼다. 형식이 틀리면 쉬운 말로 알려준다."""
    try:
        start, end = (datetime.date.fromisoformat(v.strip()) for v in values)
    except ValueError:
        print("오류: 기간은 2026-08-29 형식(연-월-일)으로 입력해 주세요.")
        return None
    if start > end:
        print("오류: 시작일이 끝일보다 늦습니다.")
        return None
    return start, end


def as_date(t):
    return datetime.date(t[0], t[1], t[2] or 1)


def fmt_date(t):
    y, mo, d = t
    return f"'{str(y)[2:]}. {mo}. {d}." if d else f"'{str(y)[2:]}. {mo}."


def norm(text):
    """숫자 비교용: 쉼표·공백 제거."""
    return re.sub(r"[,\s]", "", text or "")


def log(out_dir, msg):
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(out_dir / "실행기록.log", "a", encoding="utf-8") as f:
        f.write(f"[{stamp}] {msg}\n")


# ---------------------------------------------------------------- 동향 자료
def load_trends(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["_date"] = parse_date(r.get("일자", ""))
    return rows


def validate_trends(rows, start, end):
    """필수 항목·형식·중복을 점검하고 기간 내/외를 표시한다."""
    msgs, seen_no, seen_key = [], set(), set()
    for r in rows:
        no = r.get("번호", "?")
        for col in ("번호", "일자", "구분", "내용", "출처"):
            if not (r.get(col) or "").strip():
                msgs.append(("오류", f"번호 {no}: '{col}' 항목이 비어 있음"))
        if r["_date"] is None:
            msgs.append(("오류", f"번호 {no}: 일자 형식이 올바르지 않음(예: 2026-09-28 또는 2026-09)"))
        if r.get("구분") and r["구분"] not in GUBUN:
            msgs.append(("오류", f"번호 {no}: 구분은 산·학·관·해외 중 하나여야 함"))
        if no in seen_no:
            msgs.append(("오류", f"번호 {no}: 번호가 중복됨"))
        seen_no.add(no)
        key = (r.get("출처"), r.get("일자"), r.get("주체"))
        if key in seen_key:
            msgs.append(("주의", f"번호 {no}: 같은 출처·일자·주체의 자료가 이미 있음(중복 가능성)"))
        seen_key.add(key)
        r["_in"] = bool(r["_date"]) and start <= as_date(r["_date"]) <= end
    out = sum(1 for r in rows if not r["_in"])
    msgs.append(("정보", f"수집 기간 {start}~{end}: 기간 내 {len(rows) - out}건, 기간 외 {out}건(기간 외 자료는 '기간 내' 판단에서 제외)"))
    return msgs


# ---------------------------------------------------------------- 이슈 선별
def evaluate(rows, defs):
    result = []
    for issue in defs["이슈"]:
        code = issue["코드"]
        mine = [r for r in rows if r.get("이슈코드") == code]
        n_actor = len(issue.get("필요주체", []))
        purpose = "상" if n_actor >= 3 else "중" if n_actor == 2 else "하"
        in_period = any(r["_in"] for r in mine)
        follow = any((r.get("후속일정") or "").strip() for r in mine)
        timely = "상" if (in_period and follow) else "중" if (in_period or follow) else "하"
        actors = {a.strip() for r in mine for a in re.split(r"[·,/]", r.get("주체", "")) if a.strip()}
        policy = any(r.get("정책연계") == "Y" for r in mine)
        impact = "상" if policy else "중" if len(actors) >= 2 else "하"
        why = (f"필요 주체 {'·'.join(issue.get('필요주체', [])) or '없음'}; "
               f"기간 내 자료 {'있음' if in_period else '없음'}, 후속 일정 {'있음' if follow else '없음'}; "
               f"정책 연계 {'있음' if policy else '없음'}, 관련 주체 {len(actors)}곳")
        result.append({"코드": code, "이슈명": issue["이슈명"], "분류": issue.get("분류", ""),
                       "목적 부합성": purpose, "시의성": timely, "파급력": impact,
                       "자료 수": len(mine), "판단 근거": why})
    result.sort(key=lambda x: (-RATING[x["목적 부합성"]], -RATING[x["시의성"]], -RATING[x["파급력"]], -x["자료 수"], x["코드"]))
    return result


def cmd_select(args):
    rules = load_json(args.규칙)
    rows = load_trends(args.동향)
    defs = load_json(args.정의)
    period = parse_period(args.기간)
    if period is None:
        return 1
    start, end = period
    out = Path(args.출력)
    out.mkdir(parents=True, exist_ok=True)
    (out / "candidates.txt").unlink(missing_ok=True)  # 지난 실행의 후보가 남지 않게 지움

    msgs = validate_trends(rows, start, end)
    lines = ["[1단계] 동향 자료 점검 결과", "-" * 60]
    lines += [f"[{lv}] {m}" for lv, m in msgs]
    n_err = sum(1 for lv, _ in msgs if lv == "오류")
    lines.append(f"=> 오류 {n_err}건, 주의 {sum(1 for lv, _ in msgs if lv == '주의')}건")
    (out / "1_동향_점검결과.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))
    if n_err:
        print("\n오류가 있어 이슈 선별을 중단합니다. 동향 자료를 고친 뒤 다시 실행하세요.")
        log(out, f"select 중단: 동향 자료 오류 {n_err}건")
        return 1

    ranked = evaluate(rows, defs)
    n_pick = rules["선별"]["선별_개수"]
    path = out / "2_이슈_선별_결과.csv"
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["순위", "이슈코드", "이슈명", "분류", "목적 부합성", "시의성", "파급력", "자료 수", "선별 후보", "판단 근거"])
        for i, r in enumerate(ranked, 1):
            w.writerow([i, r["코드"], r["이슈명"], r["분류"], r["목적 부합성"], r["시의성"], r["파급력"],
                        r["자료 수"], "선별" if i <= n_pick else "", r["판단 근거"]])
    print("\n[2단계] 이슈 선별 결과 (목적 부합성 → 시의성 → 파급력 순)")
    print("-" * 60)
    for i, r in enumerate(ranked, 1):
        mark = "★선별 후보" if i <= n_pick else ""
        print(f"{i}위 [{r['코드']}] {r['이슈명']}  | 목적 {r['목적 부합성']} · 시의성 {r['시의성']} · 파급력 {r['파급력']}  {mark}")
    codes = ",".join(r["코드"] for r in ranked[:n_pick])
    (out / "candidates.txt").write_text(codes + "\n", encoding="utf-8")  # 실행.bat이 읽는 파일
    print(f"\n저장: {path}")
    if os.environ.get("AGENT_LAUNCHER"):  # 실행 파일(더블클릭)로 열었을 때는 곧바로 확정 질문이 나온다
        print("\n★ 확인 지점 1: 위 선별 후보는 초안입니다. 아래에서 담당자가 이슈를 확정해 주세요.")
    else:
        print("\n★ 확인 지점 1: 위 선별 후보는 초안입니다. 담당자가 이슈를 확정한 뒤 다음 명령으로 안건서를 만드세요.")
        print(f"   python3 agenda_agent.py build --이슈 {codes}")
    log(out, f"select 완료: 선별 후보 {codes}")
    return 0


# ---------------------------------------------------------------- 안건서 점검(업무 규칙)
NUM_RE_CACHE = {}


def num_tokens(text, units):
    key = tuple(units)
    if key not in NUM_RE_CACHE:
        alt = "|".join(re.escape(u) for u in sorted(units, key=len, reverse=True))
        NUM_RE_CACHE[key] = re.compile(r"(\d[\d,\.]*)\s?(" + alt + r")")
    text = re.sub(r"\d{1,2}\.\s?\d{1,2}\.", " ", text)  # 날짜(9. 28.)는 제외
    return {norm(m.group(1)) + m.group(2).replace(" ", "") for m in NUM_RE_CACHE[key].finditer(text)}


def agenda_texts(a):
    """(위치, 문장) 목록: 안건서에 실제로 들어가는 모든 문장."""
    t = [("제목", a["제목"])]
    t += [("요약", s) for s in a["요약"]]
    for h in a["현황"]:
        t.append(("현황 판단", h["판단"]))
        t += [("현황 근거", e["문구"]) for e in h.get("근거", [])]
    for d in a["논의"]:
        t += [("논의 질문", d["질문"]), ("논의 쟁점", d["쟁점"]), ("논의 선택지", d["선택지"])]
        t += [("논의 의견", e["문구"]) for e in d.get("의견", [])]
        t += [("논의 추가", s) for s in d.get("추가", [])]
    t.append(("회의 산출물", a["산출물"]))
    for c in a["조치"]:
        t.append(("향후 조치", c["문구"]))
        t += [("향후 조치", s) for s in c.get("하위", [])]
    t += [("확인 필요사항", s) for s in a.get("확인필요", [])]
    return t


def lint(idx, code, a, rows_by_no, code_rows, rules, tally):
    """안건 하나를 업무 규칙으로 점검한다. (수준, 안건, 항목, 메시지) 목록을 돌려준다."""
    R = rules["안건서"]
    out = []

    def bad(level, item, msg):
        tally[item] = tally.get(item, 0) + 1
        out.append((level, f"안건 {idx}", item, msg))

    def touch(item):
        tally.setdefault(item, 0)

    for item in ("요약 문장 수", "금지 표현", "제목 표현", "미확인 표시", "문장 끝 마침표", "근거 번호",
                 "자료에 없는 수치", "논의사항 구성", "회의 산출물", "향후 조치 1:1", "담당·시기 임의 기재", "산·학·관 표기"):
        touch(item)

    # 1) 요약 3문장 이내
    n_sent = len([s for s in a["요약"] if s.strip()])
    if n_sent > R["요약_최대문장"]:
        bad("오류", "요약 문장 수", f"요약박스가 {n_sent}문장(최대 {R['요약_최대문장']}문장)")
    # 2) 금지 표현 / 3) 제목 금지어 / 4) 미확인 표시 / 5) 마침표
    for where, s in agenda_texts(a):
        for w in R["금지_표현"]:
            if w in s:
                bad("오류", "금지 표현", f"[{where}] '{w}' 사용: {s[:40]}…")
        if where != "확인 필요사항":
            for w in R["본문_미확인표시"]:
                if w in s:
                    bad("주의", "미확인 표시", f"[{where}] 본문에 미확인 표시 '{w}': 문서 말미 '확인 필요사항'으로 이동 권장")
        if where not in ("제목", "요약") and s.rstrip().endswith("."):
            bad("주의", "문장 끝 마침표", f"[{where}] 마침표로 끝남: {s[:40]}…")
    for w in R["제목_금지어"]:
        if w in a["제목"]:
            bad("주의", "제목 표현", f"제목에 '{w}' 사용: {a['제목']}")
    # 6) 근거 번호 존재
    cited = set()
    for h in a["현황"]:
        for e in h.get("근거", []):
            cited.update(e["번호"])
    for d in a["논의"]:
        cited.update(d.get("근거번호", []))
    for n in sorted(cited):
        if str(n) not in rows_by_no:
            bad("오류", "근거 번호", f"동향 자료에 번호 {n}이(가) 없음")
    # 7) 자료에 없는 수치
    src = norm(" ".join(rows_by_no[str(n)]["내용"] + " " + rows_by_no[str(n)].get("후속일정", "")
                        for n in cited if str(n) in rows_by_no)
               + " ".join(r["내용"] + " " + r.get("후속일정", "") for r in code_rows))
    for where, s in agenda_texts(a):
        for tok in sorted(num_tokens(s, R["수치_단위"])):
            if tok not in src:
                bad("오류", "자료에 없는 수치", f"[{where}] '{tok}'이(가) 근거 자료에 없음: {s[:40]}…")
    # 8) 논의사항 구성(쟁점·선택지·의견)
    for k, d in enumerate(a["논의"], 1):
        for col in ("질문", "쟁점", "선택지"):
            if not (d.get(col) or "").strip():
                bad("오류", "논의사항 구성", f"논의사항 {k}: '{col}'이(가) 비어 있음")
        if not d.get("의견"):
            bad("오류", "논의사항 구성", f"논의사항 {k}: 산·학·관 의견 요청이 없음")
    # 9) 회의 산출물
    if not (a.get("산출물") or "").strip():
        bad("오류", "회의 산출물", "회의 산출물이 비어 있음")
    # 10) 향후 조치 1:1
    n_q = len(a["논의"])
    mapped = [n for c in a["조치"] for n in c.get("논의번호", [])]
    if sorted(mapped) != list(range(1, n_q + 1)) or len(a["조치"]) != n_q:
        bad("오류", "향후 조치 1:1", f"논의사항 {n_q}개와 향후 조치 {len(a['조치'])}개가 1:1로 연결되지 않음(연결 번호 {mapped})")
    # 11) 담당·시기 임의 기재
    for c in a["조치"]:
        for w in R["담당_시기_금지_표시"]:
            if w in c["문구"]:
                bad("오류", "담당·시기 임의 기재", f"자료에 없는 담당·시기 표기 '{w}': {c['문구'][:40]}…")
    # 12) 산·학·관 표기
    for d in a["논의"]:
        for e in d.get("의견", []):
            head = re.split(r"[(（]", e["주체"])[0]
            if not all(p in R["허용_주체"] for p in head.split("·")):
                bad("오류", "산·학·관 표기", f"허용되지 않은 주체 표기 '{e['주체']}'(산·학·관만 사용)")
    return out


def lint_overlap(agendas, rules, tally):
    """안건 간 비슷한 문장(중복) 점검."""
    th = rules["안건서"]["중복_기준"]
    tally.setdefault("안건 간 중복", 0)
    out = []
    lines = []
    for i, (code, a) in enumerate(agendas, 1):
        for where, s in agenda_texts(a):
            if where in ("현황 근거", "논의 쟁점", "논의 선택지", "향후 조치") and len(s) > 15:
                lines.append((i, where, s))
    for x in range(len(lines)):
        for y in range(x + 1, len(lines)):
            if lines[x][0] != lines[y][0]:
                ratio = difflib.SequenceMatcher(None, lines[x][2], lines[y][2]).ratio()
                if ratio >= th:
                    tally["안건 간 중복"] += 1
                    out.append(("주의", f"안건 {lines[x][0]}·{lines[y][0]}", "안건 간 중복",
                                f"비슷한 문장(유사도 {ratio:.2f}): '{lines[x][2][:30]}…' / '{lines[y][2][:30]}…'"))
    return out


# ---------------------------------------------------------------- 안건서 조립
def emphasize(text, phrase):
    return text.replace(phrase, f"**{phrase}**", 1) if phrase and phrase in text else text


def build_agenda_json(idx, a, cited_rows):
    C = lambda t: {"type": "circle", "text": t}
    S = lambda t: {"type": "sub", "text": t}
    summary = ". ".join(s.strip() for s in a["요약"]) + "."
    s1 = []
    for h in a["현황"]:
        s1.append(C(emphasize(h["판단"], h.get("강조", ""))))
        s1 += [S(e["문구"]) for e in h.get("근거", [])]
    s2 = []
    for d in a["논의"]:
        s2.append(C(d["질문"]))
        s2.append(S("쟁점: " + d["쟁점"]))
        s2.append(S("선택지: " + d["선택지"]))
        s2 += [S(f"{e['주체']}: {e['문구']}") for e in d.get("의견", [])]
        s2 += [S(t) for t in d.get("추가", [])]
    s2.append({"type": "conclusion", "lines": [f"**회의 산출물**: {a['산출물']}"]})
    s3 = []
    for c in a["조치"]:
        s3.append(C(c["문구"]))
        s3 += [S(t) for t in c.get("하위", [])]
    return {"title": a["제목"], "subtitle": f"(안건 {idx} · {a['핵심질문']})", "summary": summary,
            "sections": [{"heading": "추진배경 및 현황 (설명사항)", "blocks": s1},
                         {"heading": "주요 논의사항 (의견 요청)", "blocks": s2},
                         {"heading": "향후 조치", "blocks": s3}]}


def build_appendix(agendas, rows_by_no, rules):
    """참고 자료(근거 표, 확인 필요사항, 출처 목록)를 자동으로 만든다."""
    C = lambda t: {"type": "circle", "text": t}
    S = lambda t: {"type": "sub", "text": t}
    table_rows, used = [], {}
    for idx, (code, a) in enumerate(agendas, 1):
        nums = set()
        for h in a["현황"]:
            for e in h.get("근거", []):
                nums.update(e["번호"])
        for d in a["논의"]:
            nums.update(d.get("근거번호", []))
        for n in sorted(nums, key=lambda n: rows_by_no[str(n)]["_date"]):
            r = rows_by_no[str(n)]
            table_rows.append([str(idx), fmt_date(r["_date"]), r["내용"], f"[{n}] {r['출처']}"])
            used[n] = r
    blocks = [{"type": "table", "headers": ["안건", "일자", "근거 사실", "출처"], "widths": [700, 1100, 4900, 2326],
               "align": ["center", "center", "left", "left"], "rows": table_rows},
              {"type": "heading", "text": "확인 필요사항"}]
    for idx, (code, a) in enumerate(agendas, 1):
        if a.get("확인필요"):
            blocks.append(C(f"안건 {idx}"))
            blocks += [S(t) for t in a["확인필요"]]
    blocks += [C("공통"), S(rules["안건서"]["공통_확인필요"])]
    src_rows = [[f"[{n}]", r["출처"], fmt_date(r["_date"]), r.get("주소", "")] for n, r in sorted(used.items())]
    return [{"title": "참고 1. 안건별 근거자료 및 확인 필요사항", "blocks": blocks},
            {"title": "참고 2. 출처 목록", "blocks": [{"type": "table", "headers": ["번호", "기관·매체", "일자", "주소"],
             "widths": [600, 2600, 1100, 4726], "align": ["center", "left", "center", "left"], "rows": src_rows}]}]


def cmd_build(args):
    rules = load_json(args.규칙)
    rows = load_trends(args.동향)
    defs = load_json(args.정의)
    period = parse_period(args.기간)
    if period is None:
        return 1
    start, end = period
    out = Path(args.출력)
    out.mkdir(parents=True, exist_ok=True)
    (out / "3_안건서_점검결과.txt").unlink(missing_ok=True)  # 지난 실행의 점검결과가 남지 않게 지움
    validate_trends(rows, start, end)
    rows_by_no = {r["번호"]: r for r in rows}
    issues = {i["코드"]: i for i in defs["이슈"]}
    codes = [c.strip() for c in args.이슈.split(",") if c.strip()]

    msgs = []
    for c in codes:
        if c not in issues:
            print(f"오류: 이슈코드 '{c}'가 이슈 정의에 없습니다.")
            return 1
        if "안건" not in issues[c]:
            print(f"오류: 이슈 '{c}'에는 안건 초안이 없어 안건서를 만들 수 없습니다(안건 초안을 이슈 정의에 추가하세요).")
            return 1

    ranked = [r["코드"] for r in evaluate(rows, defs)][: rules["선별"]["선별_개수"]]
    for c in codes:
        if c not in ranked:
            msgs.append(("주의", "선별 확인", "확인 지점", f"이슈 {c}는 선별 후보(상위 {len(ranked)}개)에 없음: 담당자 판단으로 확정된 것으로 처리"))

    tally, agendas = {}, []
    for idx, c in enumerate(codes, 1):
        a = issues[c]["안건"]
        agendas.append((c, a))
        code_rows = [r for r in rows if r.get("이슈코드") == c]
        msgs += lint(idx, c, a, rows_by_no, code_rows, rules, tally)
    msgs += lint_overlap(agendas, rules, tally)

    n_err = sum(1 for m in msgs if m[0] == "오류")
    n_warn = sum(1 for m in msgs if m[0] == "주의")
    lines = ["[3~5단계] 안건서 업무 규칙 점검 결과", "-" * 60,
             f"확정 이슈: {', '.join(codes)}  |  정의 파일: {Path(args.정의).name}", ""]
    lines.append("규칙별 결과")
    for item, n in tally.items():
        lines.append(f"  {'통과' if n == 0 else f'위반 {n}건':<8} {item}")
    lines.append("")
    if msgs:
        lines.append("세부 내용")
        lines += [f"  [{lv}] {where} - {item}: {msg}" for lv, where, item, msg in msgs]
    lines.append("")
    lines.append(f"=> 오류 {n_err}건, 주의 {n_warn}건")
    report = "\n".join(lines) + "\n"
    (out / "3_안건서_점검결과.txt").write_text(report, encoding="utf-8")
    print(report)

    if n_err and not args.강제:
        print("오류가 있어 안건서를 만들지 않고 중단합니다. 이슈 정의를 고친 뒤 다시 실행하세요.")
        log(out, f"build 중단: 이슈 {','.join(codes)} 오류 {n_err}건")
        return 1

    doc = {"compact": True, "agendas": []}
    for idx, (c, a) in enumerate(agendas, 1):
        doc["agendas"].append(build_agenda_json(idx, a, None))
    doc["agendas"][-1]["appendix"] = build_appendix(agendas, rows_by_no, rules)
    stamp = args.작성일 or datetime.date.today().strftime("%Y%m%d")
    json_path = out / f"안건서_{stamp}.json"
    docx_path = out / f"안건서_{stamp}.docx"
    json_path.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    if not shutil.which("node"):
        print("Word 파일을 만들려면 Node.js가 필요합니다(https://nodejs.org/ 에서 설치). 중간 파일만 저장했습니다:", json_path)
        return 1
    run = subprocess.run(["node", str(BASE / "build_agenda.js"), str(json_path), str(docx_path)],
                         capture_output=True, text=True, encoding="utf-8", errors="replace")
    if run.returncode != 0:
        print("Word 파일 생성 실패:\n", run.stderr)
        log(out, "build 실패: Word 생성 오류")
        return 1
    print(f"저장: {docx_path}")
    print(f"저장: {out / '3_안건서_점검결과.txt'}")
    log(out, f"build 완료: 이슈 {','.join(codes)}, 오류 {n_err}건, 주의 {n_warn}건 -> {docx_path.name}")
    return 0


# ---------------------------------------------------------------- 중복 실행 방지
LOCK = BASE / ".agent.lock"
LOCK_MAX_HOURS = 12  # 이 시간이 지난 잠금은 남은 찌꺼기로 보고 무시한다


def pid_alive(pid):
    """그 번호의 프로그램이 지금 실행 중인지 확인한다."""
    if pid <= 0:
        return False
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes
            k32 = ctypes.WinDLL("kernel32", use_last_error=True)
            k32.OpenProcess.restype = wintypes.HANDLE
            k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
            k32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
            k32.CloseHandle.argtypes = [wintypes.HANDLE]
            handle = k32.OpenProcess(0x1000, False, pid)  # 정보 조회 권한
            if not handle:
                return False
            code = wintypes.DWORD()
            ok = k32.GetExitCodeProcess(handle, ctypes.byref(code))
            k32.CloseHandle(handle)
            return bool(ok) and code.value == 259  # 259 = 아직 실행 중
        except Exception:
            return True  # 확인이 안 되면 실행 중으로 본다(안전한 쪽)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def read_lock():
    try:
        pid_text, started = LOCK.read_text(encoding="utf-8").split("\n")[:2]
        return int(pid_text), started.strip()
    except Exception:
        return None


def lock_is_stale(info):
    if info is None:
        return True
    pid, started = info
    if not pid_alive(pid):
        return True
    try:
        age = datetime.datetime.now() - datetime.datetime.fromisoformat(started)
        return age > datetime.timedelta(hours=LOCK_MAX_HOURS)
    except ValueError:
        return True


def acquire_lock():
    """잠금을 얻으면 None, 이미 실행 중이면 (실행 번호, 시작 시각)을 돌려준다."""
    for _ in range(3):
        try:
            fd = os.open(str(LOCK), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            info = read_lock()
            if lock_is_stale(info):
                if read_lock() == info:  # 그 사이 다른 실행이 잠금을 바꾸지 않았는지 다시 확인
                    LOCK.unlink(missing_ok=True)
                continue
            return info
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(f"{os.getpid()}\n{datetime.datetime.now().isoformat(timespec='seconds')}\n")
        return None
    return read_lock() or (0, "알 수 없음")


def release_lock():
    info = read_lock()
    if info and info[0] == os.getpid():
        LOCK.unlink(missing_ok=True)


# ---------------------------------------------------------------- 화면 안내로 이어서 실행(run)
def ask(prompt):
    try:
        return input(prompt)
    except EOFError:
        return None


def open_path(path):
    if os.environ.get("AGENT_NO_OPEN"):
        return
    try:
        if os.name == "nt":
            os.startfile(str(path))  # noqa
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


def cmd_run(args):
    os.environ["AGENT_LAUNCHER"] = "1"
    print("=" * 60)
    print("  협의체 안건 도출 에이전트 (연습용)")
    print("=" * 60)
    print(f"\n수집 기간: {args.기간[0]} ~ {args.기간[1]}")
    ans = ask("기간을 바꾸려면 Y, 그대로 쓰려면 엔터: ")
    if ans is not None and ans.strip().upper() == "Y":
        start = ask("시작일 (예 2026-08-29): ") or args.기간[0]
        end = ask("끝일 (예 2026-09-29): ") or args.기간[1]
        args.기간 = [start.strip(), end.strip()]

    print("\n[1단계] 동향 자료 점검과 이슈 선별을 시작합니다.\n")
    rc = cmd_select(args)
    if rc:
        print("\n[멈춤] 실행 중 문제가 발생했습니다. 위 내용을 확인해 주세요.")
        return rc
    out = Path(args.출력)
    cand_file = out / "candidates.txt"
    if not cand_file.exists():
        print("[멈춤] 선별 후보 파일이 없습니다.")
        return 1
    cand = cand_file.read_text(encoding="utf-8").strip()

    print("\n" + "-" * 60)
    print(" 확인 지점: 선별 후보는 초안입니다. 담당자가 확정해 주세요.")
    print(f" 선별 후보 이슈코드: {cand}")
    print(" 안건 순서는 입력한 순서대로 정해집니다.")
    print("-" * 60)
    while True:
        ans = ask("후보 그대로 확정은 Y, 다른 이슈코드로 확정은 코드 입력(예 A,B), 종료는 N: ")
        if ans is None:
            print("\n입력이 없어 종료합니다.")
            return 0
        ans = ans.strip()
        if ans:
            break
    if ans.upper() == "N":
        print("종료합니다.")
        return 0
    args.이슈 = cand if ans.upper() == "Y" else ans
    args.작성일 = None
    args.강제 = False

    print(f"\n[2단계] 안건서를 만듭니다. 확정 이슈: {args.이슈}\n")
    rc = cmd_build(args)
    if rc:
        print("\n[멈춤] 안건서를 만들지 못했습니다. 위 내용과 결과 폴더의 3_안건서_점검결과.txt를 확인해 주세요.")
        report = out / "3_안건서_점검결과.txt"
        if report.exists():
            open_path(report)
        return rc
    print("\n완료되었습니다. 결과 폴더를 엽니다.")
    open_path(out)
    return 0


# ---------------------------------------------------------------- 실행
def main():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")  # 콘솔 인코딩 문제로 프로그램이 멈추지 않게 함
        except Exception:
            pass
    ap = argparse.ArgumentParser(description="협의체 안건 도출 에이전트(연습용)")
    sub = ap.add_subparsers(dest="명령", required=True)
    for name, helptext in (("select", "동향 자료 점검 + 이슈 선별(확인 지점 1에서 멈춤)"),
                           ("build", "확정 이슈로 안건서 만들기(업무 규칙 점검 포함)"),
                           ("run", "화면 안내에 따라 선별부터 안건서까지 이어서 진행")):
        p = sub.add_parser(name, help=helptext)
        p.add_argument("--동향", default=str(BASE / "연습자료" / "동향_자료.csv"))
        p.add_argument("--정의", default=str(BASE / "연습자료" / "이슈_정의.json"))
        p.add_argument("--규칙", default=str(BASE / "업무규칙.json"))
        p.add_argument("--기간", "--period", nargs=2, default=["2026-08-29", "2026-09-29"], metavar=("시작", "끝"))
        p.add_argument("--출력", default=str(BASE / "결과"))
        if name == "build":
            p.add_argument("--이슈", "--issues", required=True, help="확정한 이슈코드(쉼표로 구분, 예: A,B,C)")
            p.add_argument("--작성일", help="파일 이름에 넣을 날짜(기본: 오늘, 예: 20260929)")
            p.add_argument("--강제", action="store_true", help="오류가 있어도 파일 생성(권장하지 않음)")
    args = ap.parse_args()

    # 중복 실행 방지: 이미 실행 중이면 안내만 하고 끝낸다
    busy = acquire_lock()
    if busy:
        pid, started = busy
        print("이미 에이전트가 실행 중입니다. 새로 시작하지 않고 종료합니다.")
        print(f"  - 실행 시작: {started}")
        print("  - 열려 있는 실행 창에서 작업을 먼저 끝내 주세요(결과 파일이 겹쳐 덮어써지는 것을 막기 위함).")
        print(f"  - 실행 창이 이미 닫혔는데도 이 안내가 계속 나오면 '{LOCK.name}' 파일을 삭제한 뒤 다시 실행하세요.")
        sys.exit(3)
    atexit.register(release_lock)
    if os.name != "nt":
        signal.signal(signal.SIGTERM, lambda *_: sys.exit(1))  # 강제 종료 요청에도 잠금을 정리한다

    if args.명령 == "select":
        code = cmd_select(args)
    elif args.명령 == "build":
        code = cmd_build(args)
    else:
        code = cmd_run(args)
    sys.exit(code)


if __name__ == "__main__":
    main()
