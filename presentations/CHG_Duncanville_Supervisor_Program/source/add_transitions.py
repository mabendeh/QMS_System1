"""Post-process pptxgenjs output for PowerPoint compatibility.
- Renumber shape ids per slide so none are duplicated (pptxgenjs reuses an id for the slide-number placeholder).
- Reject negative extents (PowerPoint refuses files with negative cx/cy).
- Add a subtle Fade transition to every slide.
"""
import sys, zipfile, shutil, re
src = sys.argv[1]; tmp = src + ".tmp"
def renumber(x):
    n = [1]
    def sub(m):
        n[0] += 1
        return '<p:cNvPr id="%d"' % n[0]
    return re.sub(r'<p:cNvPr id="\d+"', sub, x)
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if re.match(r"ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster)\d+\.xml$", item.filename):
            x = data.decode("utf-8")
            if re.search(r'\bc[xy]="-\d', x):
                sys.exit("negative extent in " + item.filename)
            x = renumber(x)
            if item.filename.startswith("ppt/slides/") and "<p:transition" not in x:
                x = x.replace("</p:clrMapOvr>", '</p:clrMapOvr><p:transition spd="med"><p:fade/></p:transition>', 1)
            data = x.encode("utf-8")
        zout.writestr(item, data)
shutil.move(tmp, src); print("post-processed")
