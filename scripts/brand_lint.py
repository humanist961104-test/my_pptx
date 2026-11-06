"""my_pptx 브랜드 규칙 검사기.

사용:  PYTHONUTF8=1 python brand_lint.py deck.pptx [--strict]

.pptx(XML)를 직접 열어 brandguideline.md 10장 체크리스트 중 자동으로 볼 수 있는 항목을 검사한다.
표준 라이브러리만 사용한다.

검사 항목
  ERROR  팔레트 밖 색상 (도형 채움·선·글자; 차트·이미지 내부는 제외)
  ERROR  빨강 배경/빨강 도형 위 흰색이 아닌 글자
  ERROR  허용 외 글꼴 (Malgun Gothic / 맑은 고딕 / Arial / Calibri)
  ERROR  11pt 미만 글자 (출처·캡션용 9–10pt는 WARN)
  ERROR  발표자 노트 없음
  WARN   표지·마무리가 전면 빨강이 아님
  WARN   5장 이상인데 흰 바탕/빨강 바탕이 한쪽만 있음 (유형 섞기 부족)
  WARN   시각 요소(그림·차트·도형) 없는 슬라이드
  WARN   제목 아래 얇은 가로 선(밑줄) 의심
종료 코드: ERROR가 있으면 1 (--strict 이면 WARN도 1)
"""
import json
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

# 팔레트·글꼴은 스킬 루트 theme.json에서 읽는다 (MY_PPTX_THEME 환경변수로 교체 가능)
_THEME_PATH = os.environ.get("MY_PPTX_THEME") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "theme.json")
with open(_THEME_PATH, encoding="utf-8") as _f:
    THEME = json.load(_f)
PALETTE = {v.upper() for v in THEME["colors"].values()}
RED_BG = {THEME["colors"]["RED"].upper(), THEME["colors"]["RED_DEEP"].upper()}
FONTS = {f.lower() for f in THEME["fonts"]["allowed"]}
MIN_PT = THEME.get("minFontPt", 11)
CAPTION_PT = THEME.get("captionFontPt", 10)
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
EMU = 914400


def q(tag):
    pre, name = tag.split(":")
    return "{%s}%s" % (NS[pre], name)


def solid(el):
    """요소 바로 아래 solidFill/srgbClr 값"""
    if el is None:
        return None
    sf = el.find("a:solidFill/a:srgbClr", NS)
    return sf.get("val").upper() if sf is not None else None


def slide_order(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rmap = {r.get("Id"): r.get("Target") for r in rels}
    out = []
    for sid in pres.findall("p:sldIdLst/p:sldId", NS):
        t = rmap[sid.get("{%s}id" % NS["r"])]
        out.append("ppt/" + t.lstrip("/").replace("ppt/", ""))
    return out


def has_notes(z, slide_path):
    rel = slide_path.replace("slides/", "slides/_rels/") + ".rels"
    if rel not in z.namelist():
        return False
    rels = ET.fromstring(z.read(rel))
    for r in rels:
        if r.get("Type", "").endswith("/notesSlide"):
            npath = "ppt/" + r.get("Target").replace("../", "")
            note = ET.fromstring(z.read(npath))
            for sp in note.iter(q("p:sp")):
                ph = sp.find(".//p:nvPr/p:ph", NS)
                if ph is not None and ph.get("type") == "body":
                    text = "".join(t.text or "" for t in sp.iter(q("a:t"))).strip()
                    return bool(text)
    return False


def box(sp):
    off = sp.find(".//a:xfrm/a:off", NS)
    ext = sp.find(".//a:xfrm/a:ext", NS)
    if off is None or ext is None:
        return None
    x, y = int(off.get("x")) / EMU, int(off.get("y")) / EMU
    w, h = int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU
    return x, y, w, h


def inside(inner, outer):
    ix, iy, iw, ih = inner
    ox, oy, ow, oh = outer
    cx, cy = ix + iw / 2, iy + ih / 2
    return ox <= cx <= ox + ow and oy <= cy <= oy + oh


def lint(path):
    errs, warns = [], []
    with zipfile.ZipFile(path) as z:
        slides = slide_order(z)
        bgs = []
        for n, sp_path in enumerate(slides, 1):
            root = ET.fromstring(z.read(sp_path))
            tag = f"슬라이드 {n}"
            bg = solid(root.find("p:cSld/p:bg/p:bgPr", NS)) or "FFFFFF"
            bgs.append(bg)
            if bg not in PALETTE:
                errs.append(f"{tag}: 배경색 {bg} 팔레트 밖")

            red_boxes = []
            visuals = 0
            for sp in root.iter(q("p:sp")):
                sppr = sp.find("p:spPr", NS)
                fill = solid(sppr)
                ln = sppr.find("a:ln", NS) if sppr is not None else None
                lncol = solid(ln)
                geom = sppr.find("a:prstGeom", NS) if sppr is not None else None
                b = box(sp)
                has_text = any((t.text or "").strip() for t in sp.iter(q("a:t")))
                for col, what in ((fill, "채움"), (lncol, "선")):
                    if col and col not in PALETTE:
                        errs.append(f"{tag}: 도형 {what}색 {col} 팔레트 밖")
                if fill in RED_BG and b:
                    red_boxes.append(b)
                if not has_text and (fill or geom is not None):
                    visuals += 1
                if geom is not None and geom.get("prst") == "line" and b and b[3] < 0.02 and b[1] < 1.4 and b[2] > 3:
                    warns.append(f"{tag}: 제목 아래 가로 선 의심 (y={b[1]:.2f}\", w={b[2]:.2f}\") — 제목 밑줄 금지")

                # 텍스트 런 검사
                for r in sp.iter(q("a:r")):
                    t = (r.findtext("a:t", default="", namespaces=NS) or "").strip()
                    if not t:
                        continue
                    rpr = r.find("a:rPr", NS)
                    col = solid(rpr) or "000000"
                    snippet = t[:15]
                    if col not in PALETTE:
                        errs.append(f"{tag}: 글자색 {col} 팔레트 밖 — \"{snippet}\"")
                    sz = rpr.get("sz") if rpr is not None else None
                    if sz:
                        pt = int(sz) / 100
                        if pt < CAPTION_PT:
                            errs.append(f"{tag}: {pt:g}pt 글자 (최소 캡션 {CAPTION_PT}pt) — \"{snippet}\"")
                        elif pt < MIN_PT:
                            warns.append(f"{tag}: {pt:g}pt 글자(캡션·출처만 허용) — \"{snippet}\"")
                    if rpr is not None:
                        for fe in ("a:latin", "a:ea"):
                            f = rpr.find(fe, NS)
                            if f is not None and f.get("typeface", "").lower() not in FONTS and not f.get("typeface", "").startswith("+"):
                                errs.append(f"{tag}: 글꼴 '{f.get('typeface')}' 허용 외 — \"{snippet}\"")
                    on_red = bg in RED_BG or (b and any(inside(b, rb) for rb in red_boxes if rb != b)) or fill in RED_BG
                    if on_red and col != "FFFFFF":
                        errs.append(f"{tag}: 빨강 위 글자색 {col} (흰색만 허용) — \"{snippet}\"")

            visuals += len(root.findall(".//p:pic", NS)) + len(root.findall(".//p:graphicFrame", NS))
            if visuals == 0:
                warns.append(f"{tag}: 시각 요소(그림·차트·도형) 없음")
            if not has_notes(z, sp_path):
                errs.append(f"{tag}: 발표자 노트 없음")

        if bgs:
            if bgs[0] not in RED_BG:
                warns.append("표지(1장)가 전면 빨강이 아님")
            if len(bgs) > 1 and bgs[-1] not in RED_BG:
                warns.append("마무리(마지막 장)가 전면 빨강이 아님")
            if len(bgs) >= 5 and (all(b in RED_BG for b in bgs[1:-1]) or all(b not in RED_BG for b in bgs[1:-1])):
                warns.append("내지 배경이 한 종류뿐 — 섹션(E)형 빨강 슬라이드나 흰 내지를 섞으세요")
    return len(slides), errs, warns


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(2)
    n, errs, warns = lint(args[0])
    print(f"검사: {args[0]} ({n}장)")
    for e in errs:
        print("  ERROR", e)
    for w in warns:
        print("  WARN ", w)
    print(f"결과: ERROR {len(errs)} / WARN {len(warns)}")
    fail = errs or ("--strict" in sys.argv and warns)
    sys.exit(1 if fail else 0)


if __name__ == "__main__":
    main()
