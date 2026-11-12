"""Post-process pptxgenjs output for PowerPoint (desktop + mobile) compatibility.
- Give the notes master its own theme part (pptxgenjs shares theme1 between slide and notes masters).
- Renumber shape ids per part so none are duplicated.
- Refuse negative extents/offsets beyond the slide.
- Add a Fade transition to every slide (skip with --no-transitions).
"""
import sys, zipfile, shutil, re
src = sys.argv[1]; trans = "--no-transitions" not in sys.argv; tmp = src + ".tmp"
def renumber(x):
    n = [1]
    def sub(m):
        n[0] += 1
        return '<p:cNvPr id="%d"' % n[0]
    return re.sub(r'<p:cNvPr id="\d+"', sub, x)
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    names = zin.namelist()
    for item in zin.infolist():
        f = item.filename; data = zin.read(f)
        if f == "ppt/notesMasters/_rels/notesMaster1.xml.rels":
            data = data.replace(b"../theme/theme1.xml", b"../theme/theme2.xml")
        elif f == "[Content_Types].xml" and "ppt/theme/theme2.xml" not in names:
            data = data.replace(b'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
                b'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>')
            assert b"theme2.xml" in data
        elif re.match(r"ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster|notesSlides/notesSlide|notesMasters/notesMaster)\d+\.xml$", f):
            x = data.decode("utf-8")
            if re.search(r'\bc[xy]="-\d', x) or re.search(r'<a:off x="-\d', x):
                sys.exit("negative extent/offset in " + f)
            x = renumber(x)
            if trans and f.startswith("ppt/slides/") and "<p:transition" not in x:
                x = x.replace("</p:clrMapOvr>", '</p:clrMapOvr><p:transition spd="med"><p:fade/></p:transition>', 1)
            data = x.encode("utf-8")
        zout.writestr(item, data)
        if f == "ppt/theme/theme1.xml" and "ppt/theme/theme2.xml" not in names:
            zout.writestr("ppt/theme/theme2.xml", data)
shutil.move(tmp, src); print("post-processed", src)
