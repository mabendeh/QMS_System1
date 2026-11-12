import sys, glob
from PIL import Image, ImageDraw, ImageFont
files = sorted(glob.glob(sys.argv[1] + "/slide-*.jpg"))
out = sys.argv[2]; cols = 4; tw = int(sys.argv[3]) if len(sys.argv) > 3 else 760
ims = [Image.open(f) for f in files]
th = int(tw * ims[0].height / ims[0].width)
pad, lab = 30, 40
rows = (len(ims) + cols - 1) // cols
W = cols * tw + (cols + 1) * pad; H = 140 + rows * (th + lab + pad) + pad
sheet = Image.new("RGB", (W, H), (238, 242, 246)); d = ImageDraw.Draw(sheet)
fb = ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf", 44)
fr = ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf", 26)
d.text((pad, 40), "CHG Duncanville | Supervisor Productivity & Engagement Program Kickoff | Contact sheet", font=fb, fill=(11, 37, 69))
for i, im in enumerate(ims):
    r, c = divmod(i, cols); x = pad + c * (tw + pad); y = 140 + r * (th + lab + pad)
    sheet.paste(im.resize((tw, th), Image.LANCZOS), (x, y + lab))
    d.text((x, y + 4), f"Slide {i+1}", font=fr, fill=(19, 49, 92))
sheet.save(out, quality=90); print(out)
