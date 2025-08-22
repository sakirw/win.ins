
from PIL import Image
from pathlib import Path
png = Path("icon.png"); ico = Path("icon.ico")
if png.exists():
    img = Image.open(png).convert("RGBA")
    sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)]
    img.save(ico, sizes=sizes); print("icon.ico üretildi.")
else:
    print("icon.png yok; ikonsuz derlenecek.")
