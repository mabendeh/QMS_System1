"""Add a subtle Fade slide transition to every slide (post-process for pptxgenjs output)."""
import sys, zipfile, shutil, re, os
src = sys.argv[1]; tmp = src + ".tmp"
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if re.match(r"ppt/slides/slide\d+\.xml$", item.filename):
            x = data.decode("utf-8")
            if "<p:transition" not in x:
                tr = '<p:transition spd="med"><p:fade/></p:transition>'
                if "</p:clrMapOvr>" in x:
                    x = x.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + tr, 1)
                else:
                    x = x.replace("</p:cSld>", "</p:cSld>" + tr, 1)
            data = x.encode("utf-8")
        zout.writestr(item, data)
shutil.move(tmp, src); print("transitions added")
