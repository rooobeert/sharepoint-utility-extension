# Draws the extension icon (stacked lists, SharePoint green) at 1024 px and
# writes the package icons and the 300x300 store logo.
# Usage: python3 scripts/draw-icon.py   (needs Pillow; not part of the package)
from PIL import Image, ImageDraw

S = 1024
GREEN = (3, 120, 124, 255)   # SharePoint theme colour #03787C
WHITE = (255, 255, 255, 255)
SOFT = (255, 255, 255, 150)

im = Image.new('RGBA', (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(im)
d.rounded_rectangle([40, 40, S - 40, S - 40], radius=200, fill=GREEN)
for i, y in enumerate([250, 450, 650]):
    d.rounded_rectangle([230, y, 794, y + 130], radius=40, fill=WHITE if i < 2 else SOFT)
    d.ellipse([270, y + 35, 330, y + 95], fill=GREEN)

for size in (16, 32, 48, 128):
    im.resize((size, size), Image.LANCZOS).save(f'extension/icons/{size}.png', optimize=True)
im.resize((300, 300), Image.LANCZOS).save('store/logo-300.png', optimize=True)
