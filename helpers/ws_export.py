# -*- coding: utf-8 -*-
"""워크시트 내보내기 v2: EDL + 자막 SRT -> 문장 단위 워크시트 TSV

v1(어절 단위 712행) 대비 바뀐 점:
  - 행 단위가 어절 -> **문장(cue)**. 컷이 어차피 문장 단위로만 동작하므로.
  - 카드/브릿지 7개를 별도 블록으로 포함. v1에는 아예 없었다.
  - 수정 컬럼이 어절 교체 -> **문장 전체 교체**. 띄어쓰기 교정이 가능해진다.
  - 강조가 '★ 체크' -> **강조할 단어를 그대로 입력**(쉼표 구분).
  - 후보(컷후보/강조후보)는 사용자 입력 컬럼과 분리해 오른쪽에 참고용으로 둔다.

사용자는 F(컷) / G(수정문구) / H(강조) 세 컬럼만 만진다.

Usage:
    python helpers/ws_export.py c0017
Output:
    <edit>/worksheet_<name>.tsv        구글시트 붙여넣기/업로드용
    <edit>/worksheet_<name>.index.md   작업 가이드
"""
from __future__ import annotations
import json, re, subprocess, sys
from pathlib import Path

# 골드 강조 자동 후보 — 시트에는 '후보' 컬럼에만 제안으로 들어간다.
GOLD_CANDIDATES = [
    "비욘드캠퍼스", "비욘드워크", "BEYONDWORK", "캠퍼스", "교육", "네트워킹",
    "프로그램", "강사", "창업", "공유오피스", "오피스", "경쟁력", "정체성",
    "DNA", "미래", "수도권", "지점", "플랫폼", "스마트스토어", "강의", "입주자",
]

HEADER = ["구분", "번호", "시각(영상)", "위치", "현재문구",
          "컷", "수정문구", "강조",
          "컷후보", "강조후보", "비고",
          "(시스템)start", "(시스템)end"]


def discover_edit() -> Path | None:
    """F: 드라이브에서 260827_*/edit 를 발견."""
    cands = []
    for parent in Path("F:/").iterdir():
        if not parent.is_dir():
            continue
        try:
            children = list(parent.iterdir())
        except (PermissionError, OSError):
            continue
        for d in children:
            if d.is_dir() and d.name.startswith("260827") and (d / "edit").is_dir():
                cands.append(d / "edit")
    if not cands:
        return None
    cands.sort(key=lambda e: 0 if (e / "cards.json").exists() else 1)
    return cands[0]


def parse_srt(p: Path) -> list:
    """SRT -> [[t0, t1, text]]. 줄바꿈은 ' / ' 로 펼쳐 한 셀에 담는다."""
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
        t0 = g[0] * 3600 + g[1] * 60 + g[2] + g[3] / 1000
        t1 = g[4] * 3600 + g[5] * 60 + g[6] + g[7] / 1000
        cues.append([t0, t1, " / ".join(lines[2:]).strip()])
    return cues


def out_windows(edl: dict) -> list:
    """EDL ranges -> [(out_start, out_end, range)] 출력 타임라인."""
    wins, off = [], 0.0
    for r in edl["ranges"]:
        seg = float(r["end"]) - float(r["start"])
        wins.append((off, off + seg, r))
        off += seg
    return wins


def window_at(wins: list, t: float):
    """부동소수 누적 오차를 감안해 t 가 속한 윈도우를 찾는다."""
    eps = 1e-6
    for s, e, r in wins:
        if s - eps <= t < e - eps:
            return s, e, r
    return None, None, None


def rendered_offsets(edit: Path, edl: dict) -> list | None:
    """실제 렌더된 영상의 세그먼트 시작 시각.

    render.py 는 세그먼트마다 프레임 경계로 맞추느라 EDL 이론값보다 약 1프레임씩
    길게 뽑는다. 41세그먼트가 쌓이면 끝에서 1.2초쯤 밀린다 — 시트 시각을 보고
    영상에서 그 지점을 찾을 때 이 차이가 그대로 오차가 된다.
    clips_graded/ 에 세그먼트 파일이 다 있으면 실측값으로 보정한다.
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
        offs.append((off, off + d))
        off += d
    return offs


def to_rendered(wins: list, rend: list | None, t: float) -> float:
    """EDL 타임라인의 t 를 실제 영상 시각으로."""
    if rend is None:
        return t
    eps = 1e-6
    for i, (s, e, _r) in enumerate(wins):
        if s - eps <= t < e - eps:
            return rend[i][0] + (t - s)
    return t


def card_text(card: dict) -> str:
    """cards.json 한 항목 -> 사람이 읽고 고칠 수 있는 한 줄."""
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


def gold_hint(text: str) -> str:
    """문장에서 골드 후보 단어를 뽑아 제안 문자열로."""
    flat = text.replace(" ", "")
    hits = []
    for kw in GOLD_CANDIDATES:
        if kw.replace(" ", "") in flat and kw not in hits:
            hits.append(kw)
    return ", ".join(hits[:3])


def memo_for(text: str, gap: float, dur: float) -> str:
    notes = []
    if gap >= 0.9:
        notes.append(f"뒤 휴지 {gap:.1f}s")
    if dur < 1.0:
        notes.append(f"짧음 {dur:.1f}s")
    if re.search(r"(^|\s)(어|음|아)[.,]?(\s|$)", text):
        notes.append("필러 포함")
    if "..." in text or "…" in text:
        notes.append("말줄임")
    return "; ".join(notes)


def build_rows(edit: Path, name: str) -> list:
    edl = json.loads((edit / f"edl_{name}.json").read_text(encoding="utf-8"))
    cues = parse_srt(edit / f"cues_{name}.srt")
    wins = out_windows(edl)
    rend = rendered_offsets(edit, edl)
    cards_by_name = {}
    cards_path = edit / "cards.json"
    if cards_path.exists():
        for c in json.loads(cards_path.read_text(encoding="utf-8"))["cards"]:
            cards_by_name[c["name"]] = c

    rows = [HEADER]

    # --- 블록 1: 카드 / 브릿지 ---
    for s, e, r in wins:
        src = r.get("source", "")
        if not (src.startswith("card_") or r.get("beat") == "CARD"):
            continue
        cname = src[len("card_"):] if src.startswith("card_") else src
        card = cards_by_name.get(cname, {})
        rows.append([
            "카드", cname, mmss(to_rendered(wins, rend, s)), "CARD",
            card_text(card) or r.get("quote", ""),
            "", "", "", "", "", f"길이 {e - s:.1f}s",
            f"{s:.3f}", f"{e:.3f}",
        ])

    rows.append(["———", "", "", "", "↓ 아래는 발화 문장 ↓", "", "", "", "", "", "", "", ""])

    # --- 블록 2: 발화 문장 ---
    for i, (t0, t1, text) in enumerate(cues):
        s, e, r = window_at(wins, t0)
        beat = r.get("beat", "?") if r is not None else "?"
        gap = (cues[i + 1][0] - t1) if i + 1 < len(cues) else 0.0
        rows.append([
            "문장", str(i + 1), mmss(to_rendered(wins, rend, t0)), beat, text,
            "", "", "",
            "?" if 0.9 <= gap <= 3.0 else "",
            gold_hint(text),
            memo_for(text, gap, t1 - t0),
            f"{t0:.3f}", f"{t1:.3f}",
        ])
    return rows


def main():
    name = sys.argv[1] if len(sys.argv) > 1 else "c0017"
    edit = discover_edit()
    if edit is None:
        raise SystemExit("cannot discover 260827_*/edit under F:/")
    rows = build_rows(edit, name)
    out = edit / f"worksheet_{name}.tsv"
    out.write_text("\n".join("\t".join(r) for r in rows), encoding="utf-8-sig")

    n_card = sum(1 for r in rows[1:] if r[0] == "카드")
    n_sent = sum(1 for r in rows[1:] if r[0] == "문장")
    n_cut = sum(1 for r in rows[1:] if r[8] == "?")
    n_gold = sum(1 for r in rows[1:] if r[9])

    guide = f"""# 워크시트 작업가이드: {name}

행 {len(rows) - 1}개 = 카드 {n_card} + 구분선 1 + 문장 {n_sent}

## 만지실 컬럼은 3개뿐입니다 (F, G, H)

| 열 | 이름 | 하실 일 |
|---|---|---|
| F | **컷** | 통째로 빼고 싶은 행에 `X`. 카드 행에도 쓸 수 있습니다 |
| G | **수정문구** | 자막(또는 카드 문구)을 고칠 때 **문장 전체를 새로 입력**. 띄어쓰기 바꿔도 됩니다. 비우면 원문 유지 |
| H | **강조** | 골드로 강조할 단어를 그대로 입력. 여러 개면 쉼표로 (`비욘드캠퍼스, 교육`) |

나머지는 참고용입니다.

| 열 | 이름 | 뜻 |
|---|---|---|
| I | 컷후보 | 뒤에 0.9~3.0초 휴지가 있는 문장 (`?` {n_cut}건). 잘라도 자연스러울 가능성이 높은 지점 |
| J | 강조후보 | 자동 추출한 키워드 제안 ({n_gold}행). 쓰시려면 H로 옮겨 적으세요 |
| K | 비고 | 휴지 길이 / 짧은 큐 / 필러 포함 여부 |
| L, M | (시스템) | **건드리지 마세요.** 컷 위치 계산에 쓰입니다 |

## 카드 문구 수정법

카드 행의 E(현재문구)는 여러 줄이 ` / ` 로 이어져 있습니다.
G(수정문구)에도 같은 방식으로 ` / ` 구분해 적어주세요.
예) `비욘드캠퍼스란? / 입주사를 위한 학교` — 앞이 title, 뒤가 subtitle.

## 끝나면

시트를 그대로 두시면 제가 읽어옵니다. (또는 TSV/CSV로 내려받아 주셔도 됩니다.)
`ws_import.py` 가 컷·수정·강조를 반영해 EDL·자막·OM cues를 다시 만듭니다.
"""
    (edit / f"worksheet_{name}.index.md").write_text(guide, encoding="utf-8")
    print(f"EDIT={edit}")
    print(f"WROTE {out.name}: 카드 {n_card} + 문장 {n_sent} = {len(rows) - 1}행 "
          f"(v1은 712행이었음)")
    print(f"  컷후보 {n_cut}건 / 강조후보 {n_gold}행")


if __name__ == "__main__":
    main()
