"""Build the single-file app from app/src.
Outputs:
  app/CHG_Process_Mapper.html   standalone page (open in any modern browser)
  app/dist/artifact_page.html   body fragment published as the claude.ai artifact
"""
import pathlib, re
root = pathlib.Path(__file__).resolve().parents[1]
src = root / "app" / "src"
js = "\n".join(p.read_text() for p in sorted((src / "js").glob("*.js")))
assert "</script" not in js.lower(), "JS must not contain a closing script tag"
shell = (src / "shell.html").read_text()
frag = shell.replace("/*__APP_JS__*/", js)
(root / "app" / "dist").mkdir(exist_ok=True)
(root / "app" / "dist" / "artifact_page.html").write_text(frag)
standalone = ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
              '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
              '<style>:root{color-scheme:light}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>'
              '</head><body>' + frag + '</body></html>')
(root / "app" / "CHG_Process_Mapper.html").write_text(standalone)
print("built", len(standalone), "bytes")
