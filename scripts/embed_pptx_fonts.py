#!/usr/bin/env python3
"""
PPTX에 Pretendard 글꼴을 임베드하는 스크립트.

PowerPoint의 글꼴 임베드 규격(OOXML embeddedFontLst)을 직접 구현한다.
- PowerPoint는 TrueType 윤곽선 폰트만 임베드를 지원하므로,
  저장소의 Pretendard OTF(CFF)를 fontTools(cu2qu)로 TTF로 변환해 심는다.
- 변환된 TTF는 ~/.cache/pretendard-ttf/ 에 캐시한다.

사용법:
    python3 scripts/embed_pptx_fonts.py <발표자료.pptx> [출력.pptx]
    (출력 생략 시 <발표자료>-embedded.pptx 로 저장)

동작:
1. pptx 안의 모든 XML에서 typeface="Pretendard..." 를 수집
2. 각 패밀리의 TTF를 ppt/fonts/fontN.fntdata 로 추가
3. [Content_Types].xml, presentation.xml(.rels)에 임베드 선언 추가
   ("Pretendard" 패밀리는 Bold 슬롯에 Pretendard-Bold를 함께 연결)
"""
import io
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
OTF_DIR = REPO / "assets" / "fonts"
CACHE = Path.home() / ".cache" / "pretendard-ttf"

# PPT typeface 이름 → OTF 파일 (regular 슬롯, Pretendard는 bold 슬롯도)
FAMILY_TO_OTF = {
    "Pretendard": ("Pretendard-Regular.otf", "Pretendard-Bold.otf"),
    "Pretendard Thin": ("Pretendard-Thin.otf", None),
    "Pretendard ExtraLight": ("Pretendard-ExtraLight.otf", None),
    "Pretendard Light": ("Pretendard-Light.otf", None),
    "Pretendard Medium": ("Pretendard-Medium.otf", None),
    "Pretendard SemiBold": ("Pretendard-SemiBold.otf", None),
    "Pretendard Bold": ("Pretendard-Bold.otf", None),
    "Pretendard ExtraBold": ("Pretendard-ExtraBold.otf", None),
    "Pretendard Black": ("Pretendard-Black.otf", None),
}

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_FONT = NS_R + "/font"


def otf_to_ttf(otf_path: Path) -> Path:
    """OTF(CFF)를 TTF(TrueType 윤곽선)로 변환해 캐시에서 반환."""
    CACHE.mkdir(parents=True, exist_ok=True)
    out = CACHE / (otf_path.stem + ".ttf")
    if out.exists() and out.stat().st_mtime >= otf_path.stat().st_mtime:
        return out
    from fontTools.ttLib import TTFont
    from cu2qu.pens import Cu2QuPen
    from fontTools.pens.ttGlyphPen import TTGlyphPen

    font = TTFont(str(otf_path))
    glyph_order = font.getGlyphOrder()
    glyf = {}
    glyph_set = font.getGlyphSet()
    for name in glyph_order:
        tt_pen = TTGlyphPen(glyph_set)
        cu2qu_pen = Cu2QuPen(tt_pen, max_err=1.0, reverse_direction=True)
        glyph_set[name].draw(cu2qu_pen)
        glyf[name] = tt_pen.glyph()
    from fontTools.ttLib import newTable
    glyf_table = newTable("glyf")
    glyf_table.glyphOrder = glyph_order
    glyf_table.glyphs = glyf
    font["glyf"] = glyf_table
    loca = newTable("loca")
    font["loca"] = loca
    maxp = font["maxp"]
    maxp.tableVersion = 0x00010000
    for attr in (
        "maxZones", "maxTwilightPoints", "maxStorage", "maxFunctionDefs",
        "maxInstructionDefs", "maxStackElements", "maxSizeOfInstructions",
        "maxComponentElements", "maxComponentDepth",
    ):
        if not hasattr(maxp, attr):
            setattr(maxp, attr, 1 if attr == "maxZones" else 0)
    post = font["post"]
    post.formatType = 2.0
    post.extraNames = []
    post.mapping = {}
    post.glyphOrder = glyph_order
    for tag in ("CFF ", "VORG"):
        if tag in font:
            del font[tag]
    head = font["head"]
    head.indexToLocFormat = 1
    font.sfntVersion = "\x00\x01\x00\x00"  # OTTO → TrueType 서명 (PowerPoint 임베드 요건)
    font.save(str(out))
    return out


def collect_typefaces(z: zipfile.ZipFile) -> set:
    faces = set()
    for name in z.namelist():
        if name.endswith(".xml"):
            try:
                data = z.read(name).decode("utf-8", "ignore")
            except Exception:
                continue
            faces.update(re.findall(r'typeface="(Pretendard[^"]*)"', data))
    return faces


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = Path(sys.argv[1])
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_name(src.stem + "-embedded.pptx")

    with zipfile.ZipFile(src) as z:
        faces = sorted(f for f in collect_typefaces(z) if f in FAMILY_TO_OTF)
        if not faces:
            sys.exit("Pretendard typeface를 찾지 못했습니다 — 임베드할 것이 없습니다.")
        names = z.namelist()
        contents = {n: z.read(n) for n in names}

    # 1) 폰트 파일 준비 (family, slot(r=regular|b=bold), ttf bytes)
    entries = []
    for fam in faces:
        reg_otf, bold_otf = FAMILY_TO_OTF[fam]
        entries.append((fam, "regular", otf_to_ttf(OTF_DIR / reg_otf).read_bytes()))
        if bold_otf:
            entries.append((fam, "bold", otf_to_ttf(OTF_DIR / bold_otf).read_bytes()))

    # 2) [Content_Types].xml 에 fntdata 타입 추가
    ct = contents["[Content_Types].xml"].decode()
    if "fntdata" not in ct:
        ct = ct.replace(
            "<Default Extension=",
            '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default Extension=',
            1,
        )
    contents["[Content_Types].xml"] = ct.encode()

    # 3) presentation.xml.rels 에 font 관계 추가
    rels_name = "ppt/_rels/presentation.xml.rels"
    rels = contents[rels_name].decode()
    existing_ids = set(re.findall(r'Id="(rId\d+)"', rels))
    next_id = max((int(i[3:]) for i in existing_ids), default=0) + 1
    rel_tags, rid_of = [], {}
    for i, (fam, slot, data) in enumerate(entries, start=1):
        rid = f"rId{next_id}"
        next_id += 1
        rid_of[(fam, slot)] = rid
        rel_tags.append(
            f'<Relationship Id="{rid}" Type="{REL_FONT}" Target="fonts/font{i}.fntdata"/>'
        )
        contents[f"ppt/fonts/font{i}.fntdata"] = data
    rels = rels.replace("</Relationships>", "".join(rel_tags) + "</Relationships>")
    contents[rels_name] = rels.encode()

    # 4) presentation.xml 에 embedTrueTypeFonts + embeddedFontLst 추가
    pres = contents["ppt/presentation.xml"].decode()
    if "embedTrueTypeFonts" not in pres:
        pres = re.sub(r"<p:presentation\b", '<p:presentation embedTrueTypeFonts="1"', pres, count=1)
    font_xml = []
    for fam in faces:
        slots = f'<p:regular r:id="{rid_of[(fam, "regular")]}"/>'
        if (fam, "bold") in rid_of:
            slots += f'<p:bold r:id="{rid_of[(fam, "bold")]}"/>'
        font_xml.append(
            f'<p:embeddedFont><p:font typeface="{fam}"/>{slots}</p:embeddedFont>'
        )
    lst = "<p:embeddedFontLst>" + "".join(font_xml) + "</p:embeddedFontLst>"
    # 스키마 순서: notesSz 바로 뒤 (notesSz는 필수 요소)
    m = re.search(r"<p:notesSz[^/]*/>", pres)
    if not m:
        sys.exit("presentation.xml에서 notesSz를 찾지 못했습니다.")
    pres = pres[: m.end()] + lst + pres[m.end():]
    contents["ppt/presentation.xml"] = pres.encode()

    # 5) 저장
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])
        for name in contents:
            if name.startswith("ppt/fonts/"):
                zout.writestr(name, contents[name])

    total = sum(len(d) for _, _, d in entries)
    print(f"임베드 완료: {dst.name} — {len(faces)}개 패밀리, 폰트 {len(entries)}개({total//1024}KB)")


if __name__ == "__main__":
    main()
