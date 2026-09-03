# -*- coding: utf-8 -*-
"""워크시트 내보내기 v3: EDL + 전사 -> 어절 단위 워크시트 TSV

v2(문장 단위)에서 바뀐 점:
  - 행 단위가 문장 -> **어절**. 어절 하나만 골라 자를 수 있다.
  - 어절 시간은 Scribe 전사의 **실제 단어 타임스탬프**를 쓴다.
    (v1은 문장 길이를 글자수로 비례배분한 추정값이라 단어 중간이 잘렸다.)
  - **공백(무음) 행을 넣는다.** 말이 없는 구간 — 딴 데 보거나 뜸 들이는 곳 —
    은 어절 행 사이에 있어서 v1/v2 로는 지정할 방법이 아예 없었다.

사용자는 F(컷) / G(수정) / H(강조) 세 컬럼만 만진다.

`build_timeline()` 은 ws_import 도 그대로 불러 쓴다. 두 쪽이 같은 순서로
같은 행을 만들어야 시트의 '번호'가 어긋나지 않는다.

Usage:
    python helpers/ws_export.py c0017
"""
from __future__ import annotations
import json, re, subprocess, sys
from pathlib import Path

GAP_MIN = 0.3          # 이보다 짧은 무음은 행으로 만들지 않는다 (시트가 불어나기만 한다)
GAP_CUT_HINT = 0.9     # 이보다 긴 무음은 컷 후보로 표시

GOLD_CANDIDATES = [
    "비욘드캠퍼스", "비욘드워크", "BEYONDWORK", "캠퍼스", "교육", "네트워킹",
    "프로그램", "강사", "창업", "공유오피스", "오피스", "경쟁력", "정체성",
    "DNA", "미래", "수도권", "지점", "플랫폼", "스마트스토어", "강의", "입주자",
]

HEADER = ["구분", "번호", "시각(영상)", "문장#", "어절",
          "컷", "수정", "강조",
          "문장(전체)", "컷후보", "강조후보", "비고",
          "(시스템)start", "(시스템)end"]


def resolve_edit(footage: str | Path | None) -> Path:
    """푸티지 폴더 -> 그 안의 edit/.

    프로젝트마다 폴더가 다르므로 경로를 박아두지 않는다. 인자가 없을 때만
    현재 폴더에서 찾는다(터미널에서 직접 부를 때의 편의).
    """
    if footage:
        f = Path(footage).resolve()
        e = f if f.name == "edit" else f / "edit"
        if not e.is_dir():
            raise SystemExit(f"edit 폴더가 없습니다: {e}")
        return e
    cwd = Path.cwd()
    for c in (cwd if cwd.name == "edit" else cwd / "edit", cwd):
        if (c / "transcripts").is_dir():
            return c
    raise SystemExit("푸티지 폴더를 --footage 로 지정하세요")


def parse_srt(p: Path) -> list:
    """SRT -> [(t0, t1, text)]. 줄바꿈은 ' / ' 로 펼친다."""
    cues = []
    for block in re.split(r"\n\s*\n", p.read_text(encoding="utf-8").strip()):
        lines = [l for l in block.strip().splitlines() if l.strip()]
        if len(lines) < 3:
            continue
        m = re.match(r"(\d+):(\d+):(\d+)[,.](\d+)\s*-->\s*(\d+):(\d+):(\d+)[,.](\d+)",
                     lines[1])
        if not m:
            continue
        g = list(map(int, m.groups()))
        cues.append((g[0] * 3600 + g[1] * 60 + g[2] + g[3] / 1000,
                     g[4] * 3600 + g[5] * 60 + g[6] + g[7] / 1000,
                     " / ".join(lines[2:]).strip()))
    return cues


def is_card(r: dict) -> bool:
    return str(r.get("source", "")).startswith("card_") or r.get("beat") == "CARD"


def build_timeline(edit: Path, name: str):
    """EDL + 전사 -> (카드 행, 타임라인 행). ws_import 도 이 함수를 쓴다.

    타임라인 행은 어절과 공백이 시간순으로 섞인 리스트다. 각 행의 'no' 는
    1부터의 연번이고, 시트의 '번호' 열이 바로 이 값이다.
    """
    edl = json.loads((edit / f"edl_{name}.json").read_text(encoding="utf-8"))
    cues = parse_srt(edit / f"cues_{name}.srt")

    tr_cache = {}
    def words_of(src: str) -> list:
        if src not in tr_cache:
            d = json.loads((edit / "transcripts" / f"{src}.json").read_text(encoding="utf-8"))
            tr_cache[src] = [w for w in d["words"]
                             if w.get("type") == "word" and w.get("start") is not None]
        return tr_cache[src]

    cards_json = {}
    if (edit / "cards.json").exists():
        for c in json.loads((edit / "cards.json").read_text(encoding="utf-8"))["cards"]:
            cards_json[c["name"]] = c

    cards, items, off = [], [], 0.0
    for r in edl["ranges"]:
        seg = float(r["end"]) - float(r["start"])
        if is_card(r):
            src = str(r["source"])
            nm = src[len("card_"):] if src.startswith("card_") else src
            cards.append({"name": nm, "out_start": off, "out_end": off + seg,
                          "text": card_text(cards_json.get(nm, {}))})
            off += seg
            continue
        s, e = float(r["start"]), float(r["end"])
        ws = [w for w in words_of(str(r["source"]))
              if w["start"] < e and (w.get("end") or w["start"]) > s]
        prev = s                      # 세그먼트 안에서 직전에 소리가 끝난 지점
        spoke = False                 # 이 세그먼트에서 아직 말이 나왔는지
        for w in ws:
            a = max(float(w["start"]), s)
            b = min(float(w.get("end") or w["start"]), e)
            if a - prev >= GAP_MIN:
                # 세그먼트 맨 앞의 무음은 앞쪽이 하드컷이라 여백을 둘 필요가 없다.
                items.append({"kind": "공백", "out_start": off + (prev - s),
                              "out_end": off + (a - s), "text": "",
                              "edge_start": not spoke, "edge_end": False})
            items.append({"kind": "어절", "out_start": off + (a - s),
                          "out_end": off + (b - s), "text": (w.get("text") or "").strip()})
            prev, spoke = b, True
        if e - prev >= GAP_MIN:
            items.append({"kind": "공백", "out_start": off + (prev - s),
                          "out_end": off + (e - s), "text": "",
                          "edge_start": not spoke, "edge_end": True})
        off += seg

    # 어절을 자막 큐에 붙인다 — 시트의 문장 컬럼과 자막 재구성에 쓴다.
    for it in items:
        it["cue"] = None
        if it["kind"] != "어절":
            continue
        mid = (it["out_start"] + it["out_end"]) / 2
        for ci, (t0, t1, _txt) in enumerate(cues, 1):
            if t0 <= mid < t1:
                it["cue"] = ci
                break
    for n, it in enumerate(items, 1):
        it["no"] = n
    return cards, items, cues, edl


def rendered_offsets(edit: Path, edl: dict) -> list | None:
    """실제 렌더된 영상의 세그먼트 시작 시각.

    render.py 는 세그먼트마다 프레임 경계로 맞추느라 EDL 이론값보다 약 1프레임씩
    길게 뽑는다. 41세그먼트가 쌓이면 끝에서 1.2초쯤 밀린다 — 시트 시각을 보고
    영상에서 그 지점을 찾을 때 이 차이가 그대로 오차가 된다.
    """
    offs, off = [], 0.0
    for i, r in enumerate(edl["ranges"]):
        p = edit / "clips_graded" / f"seg_{i:02d}_{r['source']}.mp4"
        if not p.exists():
            return None
        try:
            d = float(subprocess.run(
                ["ffprobe", "-v", "error", "-show_entries", "format=duration",
                 "-of", "csv=p=0", str(p)],
                capture_output=True, text=True, check=True).stdout.strip())
        except (subprocess.CalledProcessError, ValueError, FileNotFoundError):
            return None
        offs.append(off)
        off += d
    return offs


def to_rendered(edl: dict, rend: list | None, t: float) -> float:
    """EDL 타임라인의 t 를 실제 영상 시각으로."""
    if rend is None:
        return t
    off, eps = 0.0, 1e-6
    for i, r in enumerate(edl["ranges"]):
        seg = float(r["end"]) - float(r["start"])
        if off - eps <= t < off + seg - eps:
            return rend[i] + (t - off)
        off += seg
    return t


def card_text(card: dict) -> str:
    p = card.get("props", {})
    parts = [p.get("line1"), p.get("line2"), p.get("title"), p.get("subtitle"),
             p.get("brand"), p.get("contact")]
    parts = [x for x in parts if x]
    parts += [x for x in p.get("lines", []) if x]
    return " / ".join(parts)


def mmss(sec: float) -> str:
    # 0.1초로 먼저 반올림한 뒤 분을 가른다. 그러지 않으면 179.96초가 02:60.0 이 된다.
    tenths = int(round(sec * 10))
    return f"{tenths // 600:02d}:{tenths % 600 / 10:04.1f}"


def gold_hint(word: str) -> str:
    flat = word.replace(" ", "")
    return "?" if any(k.replace(" ", "") in flat for k in GOLD_CANDIDATES) else ""


def build_rows(edit: Path, name: str) -> list:
    cards, items, cues, edl = build_timeline(edit, name)
    rend = rendered_offsets(edit, edl)
    R = lambda t: mmss(to_rendered(edl, rend, t))

    rows = [HEADER]
    for c in cards:
        rows.append(["카드", c["name"], R(c["out_start"]), "", c["text"],
                     "", "", "", "", "", "",
                     f"길이 {c['out_end'] - c['out_start']:.1f}s",
                     f"{c['out_start']:.3f}", f"{c['out_end']:.3f}"])
    rows.append(["———", "", "", "", "↓ 아래는 어절 / 무음 ↓"] + [""] * 9)

    seen_cue = set()
    for it in items:
        dur = it["out_end"] - it["out_start"]
        if it["kind"] == "공백":
            rows.append(["공백", str(it["no"]), R(it["out_start"]), "",
                         f"⏸ 무음 {dur:.1f}초", "", "", "", "",
                         "?" if dur >= GAP_CUT_HINT else "", "",
                         "말이 없는 구간",
                         f"{it['out_start']:.3f}", f"{it['out_end']:.3f}"])
            continue
        sent = ""
        if it["cue"] and it["cue"] not in seen_cue:
            seen_cue.add(it["cue"])
            sent = cues[it["cue"] - 1][2]
        memo = []
        if re.fullmatch(r"[어음아에흐]+[.,!?]*", it["text"]):
            memo.append("필러")
        if "..." in it["text"] or "…" in it["text"]:
            memo.append("말줄임")
        rows.append(["어절", str(it["no"]), R(it["out_start"]),
                     str(it["cue"] or ""), it["text"],
                     "", "", "", sent, "", gold_hint(it["text"]), "; ".join(memo),
                     f"{it['out_start']:.3f}", f"{it['out_end']:.3f}"])
    return rows


def write_xlsx(rows: list, out: Path) -> bool:
    """엑셀로도 뽑는다 — 더블클릭으로 열리고 구글시트가 그대로 가져간다.

    openpyxl 이 없으면 건너뛴다. TSV 는 어차피 항상 나온다.
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        return False
    wb = Workbook()
    ws = wb.active
    ws.title = "worksheet"
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor="1E3932")
        c.font = Font(bold=True, color="FFFFFF")
    for col in ("F", "G", "H"):          # 사람이 만지는 칸만 금색으로
        ws[f"{col}1"].fill = PatternFill("solid", fgColor="CBA258")
        ws[f"{col}1"].font = Font(bold=True, color="1E3932")
    widths = {"A": 6, "B": 7, "C": 11, "D": 6, "E": 22, "F": 5, "G": 16,
              "H": 9, "I": 46, "J": 7, "K": 10, "L": 16, "M": 12, "N": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(rows[0]))}{len(rows)}"
    wb.save(out)
    return True


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    name = args[0] if args else "c0017"
    footage = sys.argv[sys.argv.index("--footage") + 1] if "--footage" in sys.argv else None
    edit = resolve_edit(footage)
    rows = build_rows(edit, name)
    out = edit / f"worksheet_{name}.tsv"
    out.write_text("\n".join("\t".join(r) for r in rows), encoding="utf-8-sig")
    xlsx = edit / f"worksheet_{name}.xlsx"
    if write_xlsx(rows, xlsx):
        print(f"XLSX  {xlsx}")
    else:
        print("  (openpyxl 없음 — xlsx 건너뜀. pip install openpyxl)")

    n_card = sum(1 for r in rows[1:] if r[0] == "카드")
    n_word = sum(1 for r in rows[1:] if r[0] == "어절")
    n_gap = sum(1 for r in rows[1:] if r[0] == "공백")
    gap_s = sum(float(r[13]) - float(r[12]) for r in rows[1:] if r[0] == "공백")

    guide = f"""# 워크시트 작업가이드: {name}

행 {len(rows) - 1}개 = 카드 {n_card} + 구분선 1 + 어절 {n_word} + 무음 {n_gap}
잘라낼 수 있는 무음 총 {gap_s:.0f}초.

## 만지실 컬럼은 3개뿐입니다 (F, G, H)

| 열 | 이름 | 하실 일 |
|---|---|---|
| F | **컷** | 빼고 싶은 행에 `X`. 어절·무음·카드 어디든 됩니다 |
| G | **수정** | 그 어절을 대체할 글자. 지우려면 `-` 하나만 |
| H | **강조** | 어절 전체를 강조하려면 `★`. 일부만 강조하려면 그 글자를 적으세요 |

## 강조를 어절 일부만 하고 싶을 때

`경쟁력이라고` 어절에서 `경쟁력` 세 글자만 금색으로 하려면
H 에 `★` 대신 **`경쟁력`** 이라고 적으면 됩니다. 나머지 `이라고` 는 흰색으로 남습니다.
적은 글자가 그 어절 안에 없으면 반영 단계에서 경고가 나옵니다.

## 무음(⏸) 행

말이 없는 {GAP_MIN}초 이상 구간입니다. 딴 데 보시거나 뜸 들인 곳이 여기 잡힙니다.
`F`에 `X` 하면 그 정적이 사라집니다. 앞뒤 0.06초는 남겨서 말꼬리가 잘리지 않게 합니다.

## 띄어쓰기 교정 (어절 수가 바뀌는 경우)

`비욘드 워크` -> `비욘드워크` 처럼 두 어절을 하나로 합칠 때:
- `비욘드` 행의 G 에 `비욘드워크`
- `워크` 행의 G 에 `-`

## 끝나면

시트를 그대로 두시면 제가 읽어옵니다.
`ws_import.py` 가 컷·수정·강조를 반영해 EDL·자막·OM cues를 다시 만듭니다.
"""
    (edit / f"worksheet_{name}.index.md").write_text(guide, encoding="utf-8")
    print(f"EDIT={edit}")
    print(f"WROTE {out.name}: 카드 {n_card} + 어절 {n_word} + 무음 {n_gap} "
          f"= {len(rows) - 1}행")
    print(f"  잘라낼 수 있는 무음 {gap_s:.0f}초")


if __name__ == "__main__":
    main()
