from PIL import Image
import os

src = os.path.dirname(os.path.abspath(__file__))
icons = ['instagram', 'google', 'linkedin', 'linktree']

for name in icons:
    path = os.path.join(src, f'{name}.png')
    img = Image.open(path).convert('RGBA')
    r, g, b, a = img.split()
    white = Image.new('RGBA', img.size, (255, 255, 255, 0))
    white.putalpha(a)
    out = os.path.join(src, f'{name}_white.png')
    white.save(out)
    print(f'Saved {out}')
