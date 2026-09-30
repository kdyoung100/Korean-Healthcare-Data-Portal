#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
엑셀 2개 -> data.json  (중간 md 파일 없음)

    python build.py            # data.json 만 만든다
    python build.py --md       # guides/*.md 도 같이 떨군다 (눈으로 확인할 때만)

  자료입력양식.xlsx           -> 상세 페이지 본문
  중개포털_데이터_리스트.xlsx  -> 검색 목록
  둘은 '자료집ID' 열로 이어진다. 제목이 바뀌어도 안 끊긴다.

배포할 파일은 index.html 과 data.json 두 개뿐이다.
상세 본문이 data.json 안에 들어가고, index.html 이 그걸로 GUIDE_CACHE 를
미리 채우므로 상세를 열 때 따로 받아오지 않는다.
"""
import json, os, re, sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("[오류] openpyxl 이 없습니다.  pip install openpyxl")
    sys.exit(1)

BASE = Path(__file__).resolve().parent
OUT = BASE / "data.json"
IDS = BASE / "자료집ID.json"        # 발급한 ID 기록 (번호 재사용 방지). 커밋할 것.
WRITE_MD = "--md" in sys.argv

# ── 자료입력양식 (상세) ────────────────────────────────────────────
FORM_SHEET = "자료입력"
# 항목(=상세 페이지의 ## 제목)이 되지 않는 관리용 열
SKIP = {"순번", "분류", "기관", "사업", "제목", "난이도", "자료집ID",
        "메모", "비고", "확인", "작성자", "수정일"}
STARS = {"하": "★☆☆", "중": "★★☆", "상": "★★★"}
# 엑셀의 '분류' 는 자료집 문서 구조상의 이름이라 '부록' 처럼 홈페이지에
# 그대로 쓰기엔 어색한 것이 있다. 여기서만 바꿔 준다. (엑셀은 안 건드림)
CAT_MAP = {"부록": "결합·연계데이터"}

# 이 열은 상세 페이지 맨 위 '데이터셋 개요' 자리에 넣는다.
# (index.html 이 첫 '## ' 앞의 문단을 개요로 읽는다)
# 카드로도 같이 보이게 하려면 아래를 "" 로 두면 된다.
INTRO_COL = "개요"

# ── 데이터 리스트 (검색) ───────────────────────────────────────────
SHEET_DS = "데이터셋"
SHEET_TB = "데이터_테이블"
COLMAP = {
    "데이터셋ID": "id",
    "자료집ID": "guideId",          # ← 상세와 잇는 열
    "기관명": "org",
    "데이터명": "name",
    "제공유형": "type",
    "수집 시작연도": "startYear",
    "수집 종료연도": "endYear",
    "데이터 수집기간": "period",
    "데이터 항목정보(Item)": "itemInfo",
    "데이터 소개": "desc",
    "환자 수": "patientCount",
    "테이블 수": "tableCount",
    "컬럼 수": "columnCount",
    "데이터 파일 수": "fileCount",
    "최종 업데이트 일자": "lastUpdate",
    "업데이트 주기": "updateCycle",
    "주요 키워드": "keywords",
    "데이터 샘플정보(Sample)": "sampleInfo",
    "데이터 신청절차": "applyProcess",
    "비고": "note",
}

NOISE = [
    r'데이터\.+사라짐\.+', r'회신\s*대기중', r'회신기다리는중', r'확인요청\s*완료',
    r'확인필요\s*확인\s*및\s*수정완?', r'확인\s*및\s*수정완?', r'\(공개X\)',
    r'문의하신 내용에 별도로.*?감사드리겠습니다\.',
    r'같은홈페이지에서.*?받고있음',
    r'\d{3}-\d{3,4}-\d{4}\([^)]*\)',
]


def _datekey(p):
    """파일명 안의 6자리 이상 숫자(260828 같은 날짜)를 최신순 판단에 쓴다.
    '(최종)', '_수정' 처럼 글자가 붙어도 날짜가 큰 쪽이 최신본이 된다."""
    nums = [int(n) for n in re.findall(r"\d{6,}", p.stem)]
    return (max(nums) if nums else -1, p.name)


def find_latest(pattern, desc):
    hits = [p for p in BASE.glob(pattern) if not p.name.startswith("~$")]
    if not hits:
        return None
    hits.sort(key=_datekey)
    if len(hits) > 1:
        print(f"[안내] {desc} 후보 {len(hits)}개 → 최신본 사용: {hits[-1].name}")
        for p in hits[:-1]:
            print(f"       (건너뜀) {p.name}")
    return hits[-1]


def cell(v):
    """셀 값을 문자열로 정리. 1085.0 -> '1085', 날짜 -> 'YYYY. M.'"""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return f"{v.year}. {v.month}."
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if s in ("-", "None", "nan"):
        return "" if s != "-" else "-"
    return re.sub(r"[ \t\u00a0]{2,}", " ", s)


def rows_of(ws):
    header = [cell(c.value) for c in ws[1]]
    for r in ws.iter_rows(min_row=2):
        vals = [c.value for c in r]
        if not any(v is not None and str(v).strip() for v in vals):
            continue
        yield dict(zip(header, vals))


def clean(t):
    for pat in NOISE:
        t = re.sub(pat, '', t)
    return re.sub(r'[ \t]{2,}', ' ', t).strip()


def slug(name):
    s = re.sub(r'[()·<>]', '', name)
    return re.sub(r'[^\w가-힣]+', '-', s).strip('-')


# ══ 1) 자료입력양식 → 상세 본문(markdown 문자열) ═══════════════════

def render_lines(text):
    """엑셀 한 칸(Alt+Enter 로 줄바꿈) -> md 줄 목록."""
    out = []
    for ln in str(text).split('\n'):
        ln = clean(ln)
        if not ln:
            continue
        # 줄 전체가 <...> 인 부가설명은 ※ 로
        m0 = re.match(r'^-?\s*<(.+)>$', ln)
        if m0:
            out.append(f"- ※ {m0.group(1).strip()}")
            continue
        ln = re.sub(r'^[·•]\s*', '  - ', ln)
        if ln.startswith('  - '):
            out.append(ln)
            continue
        ln = re.sub(r'^-\s*', '', ln)
        m = re.match(r'^([가-힣A-Za-z0-9/\s]{2,20})[:：]\s*(.*)$', ln)
        if m and m.group(2):
            out.append(f"- **{m.group(1).strip()}**: {m.group(2).strip()}")
        elif m:
            out.append(f"- **{m.group(1).strip()}**")
        else:
            out.append(f"- {ln}")
    return out


def intro_text(val):
    """'개요' 칸 -> 상세 맨 위에 들어갈 문단.

    index.html 은 첫 '## ' 앞의 줄들을 공백으로 이어 붙여 개요로 쓴다.
    그래서 글머리표(-, ·)는 떼고 문장만 남긴다.
    """
    out = []
    for ln in str(val).split('\n'):
        ln = clean(re.sub(r'^\s*[-·•]\s*', '', ln))
        if ln:
            out.append(ln)
    return out


def to_md(rec, fields, gid):
    # ID 주석은 넣지 않는다. index.html 의 개요 추출 루프가 '#' 도 '>' 도
    # 아닌 줄을 전부 개요로 잡아서, 주석이 화면에 그대로 찍힌다.
    # ID 는 data.json 의 guideId 에 이미 들어 있다.
    out = [f"# {rec['제목']}\n"]
    meta = []
    cat = str(rec.get('분류') or '').strip()
    if cat:
        meta.append(f"**분류**: {CAT_MAP.get(cat, cat)}")
    d = str(rec.get('난이도') or '').strip()
    if d in STARS:
        meta.append(f"**구득 난이도**: {d} {STARS[d]}")
    if meta:
        out.append("> " + " | ".join(meta) + "\n")

    # '개요' 는 카드가 아니라 맨 위 개요 자리로 보낸다
    if INTRO_COL and str(rec.get(INTRO_COL) or '').strip():
        out.extend(intro_text(rec[INTRO_COL]))
        out.append("")

    for col in fields:
        if col == INTRO_COL:
            continue
        val = rec.get(col)
        if not str(val or '').strip():
            continue
        label = "구성(변수)" if col.strip() == "구성" else col.strip()
        out.append(f"## {label}\n")
        out.extend(render_lines(val))
        out.append("")
    return '\n'.join(out)


def build_guides():
    """자료입력양식.xlsx -> {자료집ID: markdown 본문}"""
    src = find_latest("자료입력양식*.xlsx", "자료입력양식(xlsx)")
    if not src:
        print(f"[오류] 자료입력양식(xlsx)을 찾지 못했습니다. 위치: {BASE}")
        sys.exit(1)
    print(f"[상세] {src.name}")

    wb = openpyxl.load_workbook(src, data_only=True)
    if FORM_SHEET not in wb.sheetnames:
        print(f"[오류] '{FORM_SHEET}' 시트가 없습니다. 시트: {wb.sheetnames}")
        sys.exit(1)
    ws = wb[FORM_SHEET]

    hdr = [str(c.value).strip() if c.value else "" for c in ws[1]]
    if "제목" not in hdr:
        print(f"[오류] '{FORM_SHEET}' 시트 1행에 '제목' 열이 없습니다.")
        sys.exit(1)
    fields = [h for h in hdr if h and h not in SKIP]

    ledger = json.load(open(IDS, encoding="utf-8")) if IDS.exists() else {}
    bodies, titles, seen, issued, dup = {}, {}, set(), [], []
    reused = []

    # 엑셀의 '자료집ID' 칸이 비어 있을 때 쓸 보조 색인: 제목(과 옛 제목) → ID.
    # 칸을 못 채웠다고 번호가 매번 새로 발급되면 안 되기 때문이다.
    by_title = {}
    for g, info in ledger.items():
        keys = [str(info.get("title") or "").strip()]
        keys += [str(a).strip() for a in (info.get("aliases") or [])]
        for k in keys:
            if k:
                by_title.setdefault(k, g)

    for r in ws.iter_rows(min_row=2, values_only=True):
        rec = {h: v for h, v in zip(hdr, r) if h}
        title = str(rec.get("제목") or "").strip()
        if not title:
            continue
        rec["제목"] = title

        gid = str(rec.get("자료집ID") or "").strip().upper()
        if not gid:                    # 칸이 비었다 → 먼저 제목으로 기존 번호를 찾는다
            found = by_title.get(title, "")
            if found and found not in seen:
                gid = found
                reused.append((gid, title))
        if not gid:                                   # 새 자료 → 번호 발급
            n = 1
            while f"G{n:02d}" in ledger or f"G{n:02d}" in seen:
                n += 1
            gid = f"G{n:02d}"
            issued.append((gid, title))
        if gid in seen:
            dup.append((gid, title))
            continue
        seen.add(gid)

        # 대장은 '이 번호를 쓴 적 있다' 는 기록일 뿐이다. 연결에는 쓰이지 않는다.
        # (번호 재사용을 막고, 제목이 언제 바뀌었는지 남겨 두기 위한 것)
        prev = ledger.get(gid, {}).get("title")
        if prev and prev != title:
            al = ledger[gid].setdefault("aliases", [])
            if prev not in al:
                al.append(prev)
            print(f"[개명] {gid}  {prev}  →  {title}   (ID 그대로라 링크 안 끊김)")
        ledger.setdefault(gid, {"title": title, "aliases": []})["title"] = title

        bodies[gid] = to_md(rec, fields, gid)
        titles[gid] = title

    json.dump(ledger, open(IDS, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print(f"       상세 {len(bodies)}건 / 항목(열): {', '.join(fields)}")
    if reused:
        print(f"[안내] 자료집ID 칸이 비어 있지만 제목으로 기존 번호를 찾아 붙였습니다 {len(reused)}개:")
        for g, t in reused:
            print(f"       {g}  {t}")
        print("       ※ 엑셀 '자료집ID' 열에 적어 두면 제목을 바꿔도 안전합니다.")
    if issued:
        print(f"[확인 요망] 자료집ID 가 비어 있어 새로 발급한 항목 {len(issued)}개:")
        for g, t in issued:
            print(f"       {g}  {t}")
        print("       ※ 엑셀의 '자료집ID' 열에 이 번호를 적고, 자료집ID.json 을 커밋하세요.")
    if dup:
        print(f"[경고] 자료집ID 가 중복된 줄 {len(dup)}개 (뒤엣것은 건너뜀):")
        for g, t in dup:
            print(f"       {g}  {t}")

    if WRITE_MD:
        outdir = BASE / "guides"
        outdir.mkdir(exist_ok=True)
        for gid, body in bodies.items():
            (outdir / f"{slug(titles[gid])}.md").write_text(body, encoding="utf-8")
        print(f"       (--md) guides/*.md {len(bodies)}개도 떨궜습니다. 배포엔 필요 없습니다.")

    return bodies, titles


# ══ 2) 데이터 리스트 → 데이터셋 목록 ═══════════════════════════════

def load_tables(wb):
    """데이터셋ID -> 테이블 목록"""
    if SHEET_TB not in wb.sheetnames:
        return {}
    out = {}
    for r in rows_of(wb[SHEET_TB]):
        ds_id = cell(r.get("데이터셋ID"))
        if not ds_id:
            continue
        en, ko = cell(r.get("테이블명(영문)")), cell(r.get("테이블명(국문)"))
        if en == "-":
            en = ""
        if not (en or ko):
            continue
        out.setdefault(ds_id, []).append({
            "en": en,
            "ko": ko,
            "combined": cell(r.get("결합 제공")).upper() == "O",
            "solo": cell(r.get("단독 제공")).upper() == "O",
            "startYear": cell(r.get("시작연도")),
            "endYear": cell(r.get("종료연도")),
            "note": cell(r.get("비고")),
        })
    return out


def finish_years(rec):
    """검색용 연도 목록 (2007 로 검색하면 1998~2024 데이터도 걸리게)"""
    try:
        s, e = int(rec["startYear"]), int(rec["endYear"])
        rec["years"] = list(range(min(s, e), max(s, e) + 1)) if 1900 < s < 2100 else []
    except (ValueError, TypeError):
        rec["years"] = []


def fill_period(rec):
    """'데이터 수집기간' 칸이 비어 있으면 수집 시작·종료연도로 채운다 (예: 2013 ~ 2024).
    엑셀은 건드리지 않고 data.json 에만 넣는다. 채운 건수를 알 수 있게 True 를 돌려준다."""
    if rec["period"] and rec["period"] != "-":
        return False
    s, e = rec["startYear"].strip(), rec["endYear"].strip()
    s, e = ("" if s == "-" else s), ("" if e == "-" else e)
    if not s and not e:
        return False
    rec["period"] = s if s == e else (f"{s} ~ {e}" if s and e else (f"{s} ~" if s else f"~ {e}"))
    return True


# ── PubMed 논문 연결 ───────────────────────────────────────────────
# 논문은 논문DB.xlsx 에서 읽는다(paper_db.py fetch 로 모으고, 검수 후 커밋).
# 빌드는 PubMed 를 부르지 않는다 — 검수 안 된 논문이 CI 에서 몰래 늘어나지 않게.
PUBMED_MAX = 200                    # 데이터셋당 data.json 에 넣을 논문 수 (연구자 검색에 전부 필요)
SHOW_STATUS = {"확인"}             # 사이트에 보일 검수상태. 검수 전 논문까지 보이려면 {"확인", "미검수"}


def build_papers(dataset_ids):
    """→ (papers, researchers). 논문DB.xlsx 가 없으면 둘 다 빈 값"""
    if not (BASE / "논문DB.xlsx").exists():
        return {}, {}
    from paper_db import load_for_build
    papers, researchers = load_for_build(dataset_ids, SHOW_STATUS, PUBMED_MAX)
    if papers:
        print(f"[논문] 논문DB 연결 {len(papers)}개 데이터셋: "
              + ", ".join(f"{k} {v['total']}편(검수 {v['verified']})" for k, v in papers.items())
              + f" / 연구자 {len(researchers)}명")
    return papers, researchers


def main():
    bodies, titles = build_guides()

    src = find_latest("중개포털_데이터_리스트*.xlsx", "데이터 리스트(xlsx)")
    if not src:
        print(f"[오류] 데이터 리스트(xlsx)를 찾지 못했습니다. 위치: {BASE}")
        sys.exit(1)
    print(f"[목록] {src.name}")

    wb = openpyxl.load_workbook(src, data_only=True)
    if SHEET_DS not in wb.sheetnames:
        print(f"[오류] '{SHEET_DS}' 시트가 없습니다. 시트: {wb.sheetnames}")
        sys.exit(1)
    tables = load_tables(wb)

    datasets, seen_ids, linked, no_guide, bad_id = [], set(), 0, [], []
    period_filled = []
    for r in rows_of(wb[SHEET_DS]):
        rec = {key: cell(r.get(col)) for col, key in COLMAP.items()}
        if not rec["name"]:
            continue
        if not rec["id"]:
            rec["id"] = f"AUTO-{len(datasets)+1:04d}"
        if rec["id"] in seen_ids:
            print(f"[경고] 데이터셋ID 중복 → 건너뜀: {rec['id']} ({rec['name']})")
            continue
        seen_ids.add(rec["id"])

        rec["tables"] = tables.get(rec["id"], [])
        if not rec["tableCount"] and rec["tables"]:
            rec["tableCount"] = str(len(rec["tables"]))

        gid = rec.pop("guideId", "").strip().upper()
        if gid and gid in bodies:
            rec["guideId"] = gid
            rec["guideFile"] = gid          # index.html 이 캐시 키로 쓴다
            linked += 1
        else:
            rec["guideId"] = ""
            rec["guideFile"] = ""
            if gid:
                bad_id.append((rec["id"], rec["name"], gid))
            else:
                no_guide.append(rec["name"])

        if fill_period(rec):
            period_filled.append(rec["name"])
        finish_years(rec)
        datasets.append(rec)

    if period_filled:
        print(f"[안내] 데이터 수집기간이 비어 있어 시작·종료연도로 채운 데이터셋 {len(period_filled)}개")

    papers, researchers = build_papers(seen_ids)
    payload = {
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "source": src.name,
        "count": len(datasets),
        "datasets": datasets,
        "guides": bodies,               # {자료집ID: markdown 본문}
        "papers": papers,               # {데이터셋ID: PubMed 논문} (논문DB.xlsx)
        "researchers": researchers,     # {연구자키: 이름·기관·논문PMID·데이터셋} (논문DB.xlsx 저자 시트)
    }
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))

    orgs = len({d["org"] for d in datasets if d["org"]})
    used = {d["guideId"] for d in datasets if d["guideId"]}
    print(f"[완료] {len(datasets)}개 데이터셋 / {orgs}개 기관 → {OUT.name}"
          f" ({OUT.stat().st_size/1024:.0f} KB)")
    print(f"       상세 연결 {linked}개 / 미연결 {len(no_guide)}개")
    print(f"       배포할 파일: index.html + data.json  (끝)")
    if bad_id:
        print(f"[경고] 자료입력양식에 없는 자료집ID {len(bad_id)}건:")
        for a, b, g in bad_id[:10]:
            print(f"       {a}  {b}  →  {g}")
    orphan = [(g, titles[g]) for g in bodies if g not in used]
    if orphan:
        print(f"[안내] 아무 데이터셋도 보지 않는 상세 {len(orphan)}건:")
        for g, t in orphan:
            print(f"       {g}  {t}")
    if no_guide:
        head = ", ".join(dict.fromkeys(no_guide))[:160]
        print(f"[안내] 상세가 없는 데이터셋 {len(no_guide)}건 (정상): {head}...")


if __name__ == "__main__":
    main()
