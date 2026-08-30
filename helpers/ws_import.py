# -*- coding: utf-8 -*-
"""워크시트 반영기 v3: 사용자가 확정한 시트 -> 컷·수정·강조를 EDL/자막/OM cues에 반영.

시트는 '표시'만 받는다 (F 컷 / G 수정 / H 강조). 어절 텍스트·시간·문장 묶음은
전부 로컬 파일(EDL + 전사 + SRT)에서 다시 만든다. 구글시트가 값을 자동 변환하거나
왕복 중 한 글자만 어긋나도 자막에 박히기 때문이다.
행 순서는 ws_export.build_timeline() 이 만들고, 시트의 '번호'가 그 연번이다.

컷 단위:
  어절 X -> 그 어절의 발화 구간을 뺀다.
  공백 X -> 그 정적을 뺀다. 앞뒤 EDGE_PAD 는 남겨 말꼬리가 잘리지 않게 한다.
  카드 X -> 그 카드 구간을 통째로 뺀다.

입력: <edit>/worksheet_<name>.tsv  또는 --input <파일|구글시트 export URL>
출력: edl_<name>_v2.json / cues_<name>_v2.srt / OM cues / cards_v2.json / ws_report

Usage:
    python helpers/ws_import.py c0017 [--input <file_or_url>]
"""
from __future__ import annotations
import csv, io, json, sys, urllib.request
from pathlib import Path

from ws_export import build_timeline, discover_edit, is_card

MAX_LINE = 18          # 자막 한 줄 최대 글자수 (자막기준 v2)
MIN_DUR = 0.8          # 자막 최소 표시 시간
EDGE_PAD = 0.06        # 무음을 자를 때 앞뒤로 남기는 여유
MIN_KEEP = 0.05        # 컷 후 남는 조각이 이보다 짧으면 반올림 찌꺼기로 보고 버린다
PUBLIC = Path(r"C:\Users\DHMoon\video-openmontage\remotion-composer\public")


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
    # 어절(E)도 읽는다 — 사람은 수정 컬럼보다 그 자리에서 글자를 고치는 게 자연스럽다.
    want = {"kind": "구분", "no": "번호", "word": "어절",
            "cut": "컷", "fix": "수정", "gold": "강조"}
    idx = {h.strip(): i for i, h in enumerate(header)}
    missing = [v for v in want.values() if v not in idx]
    if missing:
        raise SystemExit(f"시트 헤더에서 컬럼을 찾지 못함: {missing}\n헤더: {header}")
    return {k: idx[v] for k, v in want.items()}


def truthy(v: str) -> bool:
    return v.strip().upper() in ("X", "O", "V", "★", "1", "Y", "TRUE")


def wrap2(t: str) -> str:
    """18자 기준 2줄로. 윗줄이 아랫줄보다 길지 않게 (방송 관례)."""
    if len(t) <= MAX_LINE:
        return t
    mid = (len(t) + 1) // 2
    spaces = [i for i, c in enumerate(t) if c == " "]
    if not spaces:
        return t
    cut = min(spaces, key=lambda i: (abs(i - mid), 0 if i <= mid else 1))
    l1, l2 = t[:cut].rstrip(), t[cut + 1:].lstrip()
    if len(l2) < len(l1):
        sp = [i for i, c in enumerate(l1) if c == " "]
        if sp:
            l1, l2 = l1[:sp[-1]].rstrip(), (l1[sp[-1] + 1:] + " " + l2).strip()
    return l1 + "\n" + l2


def fmt_ts(sec: float) -> str:
    ms = int(round(sec * 1000))
    h, rem = divmod(ms, 3600000)
    m, rem = divmod(rem, 60000)
    s, ms = divmod(rem, 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


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
    """컷 구간(출력 타임라인 초)을 EDL ranges 에 반영."""
    new_ranges, off, dropped = [], 0.0, 0
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
        keep = [(a, b) for a, b in keep if b - a > MIN_KEEP]
        off += seg
        if not keep:
            dropped += 1
            continue
        for ks, ke in keep:
            new_ranges.append({
                "source": r.get("source", ""),
                "start": round(rs + (ks - out_s), 4),
                "end": round(rs + (ke - out_s), 4),
                "beat": r.get("beat", "?"), "quote": r.get("quote", ""),
                "reason": r.get("reason", ""),
            })
    return new_ranges, sum(x["end"] - x["start"] for x in new_ranges), dropped


def card_windows(ranges: list):
    wins, off = [], 0.0
    for r in ranges:
        seg = r["end"] - r["start"]
        if is_card(r):
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
    inpath = sys.argv[sys.argv.index("--input") + 1] if "--input" in sys.argv else None
    edit = discover_edit()
    if edit is None:
        raise SystemExit("cannot discover 260827_*/edit under F:/")

    source = inpath or str(edit / f"worksheet_{name}.tsv")
    rows = load_table(source)
    cols = find_cols(rows[0])
    marks = {}
    for r in rows[1:]:
        kind = r[cols["kind"]].strip()
        if kind in ("카드", "어절", "공백"):
            marks[(kind, r[cols["no"]].strip())] = {
                "cut": truthy(r[cols["cut"]]),
                "fix": r[cols["fix"]].strip(),
                "sheet_word": r[cols["word"]].strip(),
                "gold": r[cols["gold"]].strip()}
    blank = {"cut": False, "fix": "", "sheet_word": "", "gold": ""}

    cards, items, cues, edl = build_timeline(edit, name)

    # 어절 칸을 그 자리에서 고친 경우도 수정으로 받는다. 다만 시트 텍스트를 말없이
    # 삼키면 왕복 중 깨진 글자가 자막에 박히므로, 바뀐 것은 전부 찍어서 보이게 한다.
    inline = []
    for it in items:
        if it["kind"] != "어절":
            continue
        m = marks.get(("어절", str(it["no"])))
        if not m or m["fix"] or not m["sheet_word"]:
            continue
        if m["sheet_word"] != it["text"]:
            m["fix"] = m["sheet_word"]
            inline.append((it["no"], it["text"], m["sheet_word"]))
    for c in cards:
        m = marks.get(("카드", c["name"]))
        if m and not m["fix"] and m["sheet_word"] and m["sheet_word"] != c.get("text", ""):
            m["fix"] = m["sheet_word"]
            inline.append((c["name"], c.get("text", ""), m["sheet_word"]))
    if inline:
        print(f"  어절 칸에서 직접 고치신 것 {len(inline)}건 — 수정으로 반영합니다:")
        for no, old, new in inline[:40]:
            print(f"     {no}: \"{old}\" -> \"{new}\"")
        if len(inline) > 40:
            print(f"     ... 외 {len(inline) - 40}건")
    known = {("카드", c["name"]) for c in cards} | {(i["kind"], str(i["no"])) for i in items}
    unknown = set(marks) - known
    if unknown:
        print(f"  ! 로컬에 없는 행 {len(unknown)}개 무시: {sorted(unknown)[:5]}")
    print(f"입력 {source}")
    print(f"  시트 표시 {len(marks)}행 / 로컬 카드 {len(cards)} + 어절·무음 {len(items)}")

    # --- 컷 구간 ---
    raw, n_word_cut, n_gap_cut, n_card_cut = [], 0, 0, 0
    for c in cards:
        if marks.get(("카드", c["name"]), blank)["cut"]:
            raw.append((c["out_start"], c["out_end"]))
            n_card_cut += 1
    for it in items:
        if not marks.get((it["kind"], str(it["no"])), blank)["cut"]:
            continue
        if it["kind"] == "공백":
            # 세그먼트 경계에 붙은 쪽은 어차피 하드컷이라 여백을 남기지 않는다.
            # 남기면 2프레임짜리 조각이 살아남아 화면에서 깜빡인다.
            a = it["out_start"] + (0.0 if it["edge_start"] else EDGE_PAD)
            b = it["out_end"] - (0.0 if it["edge_end"] else EDGE_PAD)
            if b - a > MIN_KEEP:
                raw.append((a, b))
                n_gap_cut += 1
        else:
            raw.append((it["out_start"], it["out_end"]))
            n_word_cut += 1
    cuts = merge_cuts(raw)
    cut_total = sum(e - s for s, e in cuts)
    print(f"  컷: 어절 {n_word_cut} + 무음 {n_gap_cut} + 카드 {n_card_cut} "
          f"-> 제거 구간 {len(cuts)}개, 총 {cut_total:.2f}s")

    # --- 자막 재구성: 살아남은 어절만 원래 문장 묶음대로 다시 잇는다 ---
    survivors = {}
    n_fix = n_gold = 0
    for it in items:
        if it["kind"] != "어절" or it["cue"] is None:
            continue
        if any(cs < it["out_end"] and ce > it["out_start"] for cs, ce in cuts):
            continue
        m = marks.get(("어절", str(it["no"])), blank)
        fix, gold = m["fix"], m["gold"]
        text = it["text"]
        if fix:
            n_fix += 1
            if fix == "-":
                text = ""
            else:
                text = fix
        # 강조: ★ 면 어절 전체, 글자를 적으면 그 부분만.
        # OM 의 goldSpan() 이 word.indexOf(gold) 로 부분문자열을 칠하므로
        # "경쟁력이라고" 어절에 "경쟁력" 을 적으면 그 세 글자만 금색이 된다.
        gold_txt = ""
        if gold:
            gold_txt = text if truthy(gold) else gold.strip()
            if gold_txt and gold_txt not in text:
                print(f"  ! 번호 {it['no']}: 강조 \"{gold_txt}\" 가 어절 "
                      f"\"{text}\" 안에 없습니다 — 강조되지 않습니다")
                gold_txt = ""
            elif gold_txt:
                n_gold += 1
        survivors.setdefault(it["cue"], []).append(
            {"t0": it["out_start"], "t1": it["out_end"], "text": text,
             "gold": gold_txt})

    new_cues = []
    for ci in sorted(survivors):
        ws = survivors[ci]
        text = " ".join(w["text"] for w in ws if w["text"]).strip()
        if not text:
            continue
        for a, b in ((" ,", ","), (" .", "."), (" ?", "?"), (" !", "!")):
            text = text.replace(a, b)
        new_cues.append({
            "start": ws[0]["t0"] - removed_before(cuts, ws[0]["t0"]),
            "end": ws[-1]["t1"] - removed_before(cuts, ws[-1]["t1"]),
            "text": text,
            "gold": [w["gold"] for w in ws if w["gold"]],
        })
    new_cues.sort(key=lambda c: c["start"])
    for i in range(1, len(new_cues)):
        if new_cues[i]["start"] < new_cues[i - 1]["end"]:
            new_cues[i - 1]["end"] = new_cues[i]["start"]
    for i, c in enumerate(new_cues):
        if c["end"] - c["start"] < MIN_DUR:
            room = new_cues[i + 1]["start"] - 0.05 if i + 1 < len(new_cues) else c["start"] + MIN_DUR
            c["end"] = max(c["end"], min(c["start"] + MIN_DUR, room))
    print(f"  자막: 큐 {len(cues)} -> {len(new_cues)} / 어절수정 {n_fix} / 강조 {n_gold}")

    srt_path = edit / f"cues_{name}_v2.srt"
    lines = []
    for i, c in enumerate(new_cues, 1):
        lines += [str(i), f"{fmt_ts(c['start'])} --> {fmt_ts(c['end'])}",
                  wrap2(c["text"]), ""]
    srt_path.write_text("\n".join(lines), encoding="utf-8")

    # --- EDL v2 ---
    new_ranges, new_total, dropped_ranges = rebuild_edl(edl, cuts)
    edl_v2_path = edit / f"edl_{name}_v2.json"
    edl_v2_path.write_text(json.dumps(
        {"sources": edl["sources"], "fps": edl.get("fps"), "ranges": new_ranges,
         "total_duration_s": round(new_total, 3), "subtitles": srt_path.name,
         "ws": {"cuts": cuts, "dropped_ranges": dropped_ranges}},
        ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"  EDL: {len(edl['ranges'])} -> {len(new_ranges)} 세그먼트, "
          f"{edl.get('total_duration_s')}s -> {new_total:.2f}s")
    slivers = [r for r in new_ranges if r["end"] - r["start"] < 0.3]
    if slivers:
        print(f"  ! 0.3초 미만으로 남은 조각 {len(slivers)}개 — 화면에서 깜빡입니다. "
              f"컷을 옆 경계까지 넓히세요:")
        for r in slivers[:5]:
            print(f"     {r['source']} {r['start']:.2f}-{r['end']:.2f} "
                  f"({r['end'] - r['start']:.2f}s)")

    # --- OM cues: 밴드는 짧은 한 줄만 담는다. 2줄 큐는 순차 단일줄로 쪼갠다 ---
    wins_card, total = card_windows(new_ranges)
    om_cues = []
    for c in new_cues:
        pieces = [l for l in wrap2(c["text"]).split("\n") if l.strip()]
        weights = [max(len(l), 1) for l in pieces]
        wsum, cur, dur = sum(weights), c["start"], c["end"] - c["start"]
        for j, (line, w) in enumerate(zip(pieces, weights)):
            end = c["end"] if j == len(pieces) - 1 else cur + dur * w / wsum
            om_cues.append({"text": line, "gold": [g for g in c["gold"] if g in line],
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

    # --- 카드 문구 수정 ---
    changed = []
    if (edit / "cards.json").exists():
        doc = json.loads((edit / "cards.json").read_text(encoding="utf-8"))
        by_name = {c["name"]: c for c in doc["cards"]}
        for c in cards:
            fix = marks.get(("카드", c["name"]), blank)["fix"]
            if not fix or c["name"] not in by_name:
                continue
            by_name[c["name"]] = apply_card_text(by_name[c["name"]], fix)
            changed.append(c["name"])
            (edit / "cards" / f"{c['name']}.props.v2.json").write_text(
                json.dumps(by_name[c["name"]]["props"], ensure_ascii=False, indent=1),
                encoding="utf-8")
        if changed:
            doc["cards"] = [by_name[c["name"]] for c in doc["cards"]]
            (edit / "cards_v2.json").write_text(
                json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")

    (edit / f"ws_report_{name}.json").write_text(json.dumps(
        {"input": source, "cut_segments": cuts, "cut_seconds": round(cut_total, 3),
         "cut_words": n_word_cut, "cut_gaps": n_gap_cut, "cut_cards": n_card_cut,
         "cues_before": len(cues), "cues_after": len(new_cues),
         "word_fixes": n_fix, "gold_marked": n_gold, "changed_cards": changed,
         "total_seconds_old": edl.get("total_duration_s"),
         "total_seconds_new": round(new_total, 3),
         "files": {"srt": srt_path.name, "edl": edl_v2_path.name,
                   "om_cues": om_path.name}},
        ensure_ascii=False, indent=1), encoding="utf-8")

    if changed:
        print(f"  카드 문구 변경: {', '.join(changed)} -> cards_v2.json")
        print("  ** 해당 카드는 재렌더가 필요합니다:")
        for n in changed:
            comp = "IntroVariant" if n == "intro" else "BwCard"
            print(f"     remotion render src/index.tsx {comp} "
                  f"{edit / 'cards' / (n + '.mp4')} "
                  f"--props={edit / 'cards' / (n + '.props.v2.json')}")
    print(f">> 다음 단계: {edl_v2_path.name} 으로 base 재빌드 -> OM 시퀀스 -> 합성")


if __name__ == "__main__":
    main()
