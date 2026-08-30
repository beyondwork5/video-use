# -*- coding: utf-8 -*-
"""워크시트 반영기 v2: 사용자가 확정한 시트 -> 컷·수정·강조를 EDL/자막/OM cues에 반영.

v1에서 고친 것:
  - v1은 cue의 **첫 행에서만** 컷 값을 읽어, 문장 중간 행의 X 를 조용히 무시했다.
    v2는 행 = 문장이므로 그 함정 자체가 없다.
  - v1 문서에 있던 '그룹번호로 부분 자름'은 구현돼 있지 않았다(딕셔너리만 만들고 버림).
    v2는 그 기능을 광고하지 않는다. 컷은 `X` 하나뿐, 단위는 문장/카드.
  - v1은 어절 교체라 띄어쓰기 교정이 불가능했다. v2는 문장 전체를 교체한다.
  - 카드(브릿지) 문구 수정을 지원한다. 바뀐 카드는 재렌더가 필요하므로 명령을 출력한다.

입력: <edit>/worksheet_<name>.tsv  또는 --input <파일|구글시트 export URL>
출력:
    edl_<name>_v2.json      컷 반영 EDL
    cues_<name>_v2.srt      컷·수정 반영 자막
    bw_260827_비욘드캠퍼스_<name>_cues.json   OM cues (강조 포함)
    cards_v2.json           카드 문구 수정본 (바뀐 게 있을 때만)
    ws_report_<name>.json   반영 요약

Usage:
    python helpers/ws_import.py c0017
    python helpers/ws_import.py c0017 --input <file_or_url>
"""
from __future__ import annotations
import csv, io, json, re, sys, urllib.request
from pathlib import Path

MAX_LINE = 18          # 자막 한 줄 최대 글자수 (자막기준 v2)
GAP_ABSORB = 3.0       # 컷한 문장 뒤 이만큼 이하의 침묵은 같이 제거
MIN_KEEP = 0.05        # 컷 후 남는 조각이 이보다 짧으면 반올림 찌꺼기로 보고 버린다
PUBLIC = Path(r"C:\Users\DHMoon\video-openmontage\remotion-composer\public")


def discover_edit() -> Path | None:
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


def load_table(path_or_url: str) -> list:
    if path_or_url.startswith("http"):
        data = urllib.request.urlopen(path_or_url, timeout=60).read().decode("utf-8-sig")
    else:
        data = Path(path_or_url).read_text(encoding="utf-8-sig")
    first = data.splitlines()[0]
    if "\t" in first:
        rows = [r.split("\t") for r in data.splitlines()]
    else:
        rows = list(csv.reader(io.StringIO(data)))
    width = max(len(r) for r in rows)
    return [r + [""] * (width - len(r)) for r in rows]


def find_cols(header: list) -> dict:
    # 실제로 쓰는 건 이 5개뿐. 나머지 컬럼은 사람이 읽으라고 있는 것이라 없어도 된다.
    want = {"kind": "구분", "no": "번호",
            "cut": "컷", "fix": "수정문구", "gold": "강조"}
    idx = {h.strip(): i for i, h in enumerate(header)}
    cols = {k: idx[v] for k, v in want.items() if v in idx}
    missing = [v for k, v in want.items() if k not in cols]
    if missing:
        raise SystemExit(f"시트 헤더에서 컬럼을 찾지 못함: {missing}\n헤더: {header}")
    return cols


def is_cut(v: str) -> bool:
    return v.strip().upper() in ("X", "O", "V", "★", "1", "Y", "TRUE")


def wrap2(t: str) -> str:
    """' / ' 로 명시된 줄바꿈을 존중하고, 없으면 18자 기준 2줄로 자른다."""
    if " / " in t:
        return "\n".join(p.strip() for p in t.split(" / ") if p.strip())
    if len(t) <= MAX_LINE:
        return t
    mid = (len(t) + 1) // 2
    spaces = [i for i, c in enumerate(t) if c == " "]
    if not spaces:
        return t
    cut = min(spaces, key=lambda i: (abs(i - mid), 0 if i <= mid else 1))
    return t[:cut].rstrip() + "\n" + t[cut + 1:].lstrip()


def parse_srt(p: Path) -> list:
    """SRT -> [(t0, t1, text)]. 줄바꿈은 ' / ' 로 펼쳐 시트 표기와 맞춘다."""
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


def fmt_ts(sec: float) -> str:
    ms = int(round(sec * 1000))
    h, rem = divmod(ms, 3600000)
    m, rem = divmod(rem, 60000)
    s, ms = divmod(rem, 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def out_windows(edl: dict) -> list:
    wins, off = [], 0.0
    for r in edl["ranges"]:
        seg = float(r["end"]) - float(r["start"])
        wins.append((off, off + seg))
        off += seg
    return wins


def window_end(wins: list, t: float) -> float:
    """t 가 속한 EDL 출력 윈도우의 끝. 컷이 세그먼트를 넘지 않게 하는 데 쓴다."""
    eps = 1e-6
    for s, e in wins:
        if s - eps <= t < e - eps:
            return e
    return t


def merge_cuts(segs: list) -> list:
    if not segs:
        return []
    segs = sorted(segs)
    out = [list(segs[0])]
    for s, e in segs[1:]:
        if s <= out[-1][1] + 1e-6:
            out[-1][1] = max(out[-1][1], e)
        else:
            out.append([s, e])
    return [tuple(x) for x in out]


def removed_before(cuts: list, t: float) -> float:
    acc = 0.0
    for s, e in cuts:
        if e <= t:
            acc += e - s
        elif s < t:
            acc += t - s
    return acc


def rebuild_edl(edl: dict, cuts: list):
    """컷 구간(출력 타임라인 초)을 EDL ranges 에 반영. 남는 조각을 소스시간으로 역변환."""
    new_ranges, off, dropped_ranges = [], 0.0, 0
    for r in edl["ranges"]:
        rs, re_ = float(r["start"]), float(r["end"])
        seg = re_ - rs
        out_s, out_e = off, off + seg
        keep, cur = [], out_s
        for cs, ce in cuts:
            if ce <= out_s or cs >= out_e:
                continue
            if cs > cur:
                keep.append((cur, cs))
            cur = max(cur, ce)
        if cur < out_e:
            keep.append((cur, out_e))
        # 시트 시간이 소수 3자리로 반올림돼 있어 경계에서 극소 조각이 남는다.
        # 한 프레임도 안 되는 조각은 컷 의도가 아니므로 버린다.
        keep = [(ks, ke) for ks, ke in keep if ke - ks > MIN_KEEP]
        off += seg
        if not keep:
            dropped_ranges += 1
            continue
        for ks, ke in keep:
            new_ranges.append({
                "source": r.get("source", ""),
                "start": round(rs + (ks - out_s), 4),
                "end": round(rs + (ke - out_s), 4),
                "beat": r.get("beat", "?"),
                "quote": r.get("quote", ""),
                "reason": r.get("reason", ""),
            })
    total = sum(x["end"] - x["start"] for x in new_ranges)
    return new_ranges, total, dropped_ranges


def card_windows(ranges: list):
    wins, off = [], 0.0
    for r in ranges:
        seg = r["end"] - r["start"]
        if r.get("beat") == "CARD" or str(r.get("source", "")).startswith("card_"):
            wins.append([round(off, 3), round(off + seg, 3)])
        off += seg
    return wins, round(off, 3)


def apply_card_text(card: dict, new_text: str) -> dict:
    """' / ' 로 구분된 새 문구를 카드 props 구조에 되돌려 넣는다."""
    p = dict(card.get("props", {}))
    parts = [x.strip() for x in new_text.split(" / ") if x.strip()]
    if "lines" in p:
        p["lines"] = parts
    elif "line1" in p:
        p["line1"] = parts[0] if parts else ""
        if len(parts) > 1:
            p["line2"] = parts[1]
    elif "brand" in p:
        p["brand"] = parts[0] if parts else ""
        if len(parts) > 1:
            p["contact"] = parts[1]
    else:
        p["title"] = parts[0] if parts else ""
        if len(parts) > 1:
            p["subtitle"] = parts[1]
        elif "subtitle" in p:
            p.pop("subtitle")
    out = dict(card)
    out["props"] = p
    return out


def main():
    name = sys.argv[1] if len(sys.argv) > 1 else "c0017"
    inpath = None
    if "--input" in sys.argv:
        inpath = sys.argv[sys.argv.index("--input") + 1]
    edit = discover_edit()
    if edit is None:
        raise SystemExit("cannot discover 260827_*/edit under F:/")

    source = inpath or str(edit / f"worksheet_{name}.tsv")
    rows = load_table(source)
    cols = find_cols(rows[0])

    edl = json.loads((edit / f"edl_{name}.json").read_text(encoding="utf-8"))
    wins = out_windows(edl)

    # 시트에서는 '표시'만 받는다 — 원문과 시간은 로컬 파일이 기준이다.
    # (구글시트가 값을 자동 변환하거나 왕복 중 한 글자만 어긋나도 자막에 박히므로.)
    marks = {}
    for r in rows[1:]:
        kind = r[cols["kind"]].strip()
        if kind not in ("카드", "문장"):
            continue                      # 구분선/빈 행
        marks[(kind, r[cols["no"]].strip())] = {
            "cut": is_cut(r[cols["cut"]]),
            "fix": r[cols["fix"]].strip(),
            "gold": r[cols["gold"]].strip(),
        }
    blank = {"cut": False, "fix": "", "gold": ""}

    local_cues = parse_srt(edit / f"cues_{name}.srt")
    cards, sents = [], []
    off = 0.0
    for r in edl["ranges"]:
        seg = float(r["end"]) - float(r["start"])
        src = str(r.get("source", ""))
        if src.startswith("card_") or r.get("beat") == "CARD":
            cname = src[len("card_"):]
            cards.append({"no": cname, "start": off, "end": off + seg,
                          **marks.get(("카드", cname), blank)})
        off += seg
    for i, (t0, t1, text) in enumerate(local_cues, 1):
        sents.append({"no": str(i), "start": t0, "end": t1, "text": text,
                      **marks.get(("문장", str(i)), blank)})

    unknown = set(marks) - {("카드", c["no"]) for c in cards} - {("문장", s["no"]) for s in sents}
    if unknown:
        print(f"  ! 시트에만 있고 로컬에 없는 행 {len(unknown)}개 무시: {sorted(unknown)[:5]}")
    print(f"입력 {source}")
    print(f"  시트 표시 {len(marks)}행 / 로컬 기준 카드 {len(cards)} + 문장 {len(sents)}")

    # --- 컷 구간 (출력 타임라인) ---
    raw_cuts = [(c["start"], c["end"]) for c in cards if c["cut"]]
    for i, s in enumerate(sents):
        if not s["cut"]:
            continue
        end = s["end"]
        # 뒤따르는 침묵도 같이 제거 — 안 그러면 앞뒤 정적이 겹쳐 어색해진다.
        nxt = sents[i + 1]["start"] if i + 1 < len(sents) else None
        if nxt is not None and 0 < nxt - end <= GAP_ABSORB:
            end = min(nxt, window_end(wins, s["start"]))
        raw_cuts.append((s["start"], end))
    cuts = merge_cuts(raw_cuts)
    cut_total = sum(e - s for s, e in cuts)
    print(f"  컷 표시 {len(raw_cuts)}건 -> 제거 구간 {len(cuts)}개, 총 {cut_total:.2f}s")

    # --- 자막 재구성 ---
    new_cues, dropped = [], 0
    for s in sents:
        if any(cs < s["end"] and ce > s["start"] for cs, ce in cuts):
            dropped += 1
            continue
        text = s["fix"] or s["text"]
        gold = [g.strip() for g in s["gold"].split(",") if g.strip()]
        new_cues.append({
            "start": s["start"] - removed_before(cuts, s["start"]),
            "end": s["end"] - removed_before(cuts, s["end"]),
            "text": text, "gold": gold,
        })
    new_cues.sort(key=lambda c: c["start"])
    for i in range(1, len(new_cues)):
        if new_cues[i]["start"] < new_cues[i - 1]["end"]:
            new_cues[i - 1]["end"] = new_cues[i]["start"]
    n_fix = sum(1 for s in sents if s["fix"])
    n_gold = sum(len(c["gold"]) for c in new_cues)
    print(f"  자막: 유지 {len(new_cues)} / 컷으로 제거 {dropped} / 문구수정 {n_fix} / 강조 {n_gold}건")

    srt_path = edit / f"cues_{name}_v2.srt"
    lines = []
    for i, c in enumerate(new_cues, 1):
        lines += [str(i), f"{fmt_ts(c['start'])} --> {fmt_ts(c['end'])}",
                  wrap2(c["text"]), ""]
    srt_path.write_text("\n".join(lines), encoding="utf-8")

    # --- EDL v2 ---
    new_ranges, new_total, dropped_ranges = rebuild_edl(edl, cuts)
    edl_v2 = {"sources": edl["sources"], "fps": edl.get("fps"),
              "ranges": new_ranges, "total_duration_s": round(new_total, 3),
              "subtitles": srt_path.name,
              "ws": {"cuts": cuts, "dropped_cues": dropped,
                     "dropped_ranges": dropped_ranges}}
    edl_v2_path = edit / f"edl_{name}_v2.json"
    edl_v2_path.write_text(json.dumps(edl_v2, ensure_ascii=False, indent=1),
                           encoding="utf-8")
    print(f"  EDL: {len(edl['ranges'])} -> {len(new_ranges)} 세그먼트, "
          f"{edl.get('total_duration_s')}s -> {new_total:.2f}s")

    # --- OM cues ---
    # OM 자막 밴드는 짧은 한 줄만 담는다. 2줄 큐는 순차 단일줄 큐로 쪼갠다
    # (시간은 글자수 비례) — _om_cues.py 와 같은 규칙.
    wins_card, total = card_windows(new_ranges)
    om_cues = []
    for c in new_cues:
        pieces = [l for l in wrap2(c["text"]).split("\n") if l.strip()]
        weights = [max(len(l), 1) for l in pieces]
        wsum, cur, dur = sum(weights), c["start"], c["end"] - c["start"]
        for j, (line, w) in enumerate(zip(pieces, weights)):
            end = c["end"] if j == len(pieces) - 1 else cur + dur * w / wsum
            om_cues.append({"text": line,
                            "gold": [g for g in c["gold"] if g in line],
                            "start": round(cur, 3), "end": round(end, 3)})
            cur = end
    too_long = [c["text"] for c in om_cues if len(c["text"]) > 28]
    if too_long:
        print(f"  ! 밴드 초과(28자) {len(too_long)}건: {too_long[:3]}")
    om_path = PUBLIC / f"bw_260827_비욘드캠퍼스_{name}_cues.json"
    om_path.write_text(json.dumps(
        {"cues": om_cues, "totalSeconds": total, "cardWindows": wins_card,
         "note": f"generated from worksheet {name}; cuts/fixes/gold by user"},
        ensure_ascii=False, indent=1), encoding="utf-8")

    # --- 카드 수정 ---
    changed_cards = []
    cards_path = edit / "cards.json"
    if cards_path.exists() and any(c["fix"] for c in cards):
        doc = json.loads(cards_path.read_text(encoding="utf-8"))
        by_name = {c["name"]: c for c in doc["cards"]}
        for c in cards:
            if not c["fix"] or c["no"] not in by_name:
                continue
            by_name[c["no"]] = apply_card_text(by_name[c["no"]], c["fix"])
            changed_cards.append(c["no"])
            (edit / "cards" / f"{c['no']}.props.v2.json").write_text(
                json.dumps(by_name[c["no"]]["props"], ensure_ascii=False, indent=1),
                encoding="utf-8")
        doc["cards"] = [by_name[c["name"]] for c in doc["cards"]]
        (edit / "cards_v2.json").write_text(
            json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")

    report = {
        "input": source, "cut_segments": cuts, "cut_seconds": round(cut_total, 3),
        "dropped_cues": dropped, "kept_cues": len(new_cues),
        "dropped_ranges": dropped_ranges, "text_fixes": n_fix, "gold_marked": n_gold,
        "changed_cards": changed_cards,
        "total_seconds_old": edl.get("total_duration_s"),
        "total_seconds_new": round(new_total, 3),
        "files": {"srt": srt_path.name, "edl": edl_v2_path.name,
                  "om_cues": om_path.name},
    }
    (edit / f"ws_report_{name}.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")

    if changed_cards:
        print(f"  카드 문구 변경: {', '.join(changed_cards)} -> cards_v2.json")
        print("  ** 해당 카드는 재렌더가 필요합니다:")
        for n in changed_cards:
            comp = "IntroVariant" if n == "intro" else "BwCard"
            print(f"     remotion render src/index.tsx {comp} "
                  f"{edit / 'cards' / (n + '.mp4')} "
                  f"--props={edit / 'cards' / (n + '.props.v2.json')}")
    print(">> 다음 단계: edl_%s_v2.json 으로 base 재빌드 -> OM 시퀀스 -> 합성" % name)


if __name__ == "__main__":
    main()
