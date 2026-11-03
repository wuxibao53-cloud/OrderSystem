from PIL import Image, ImageDraw
import os

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "webui")
S = 1024


def build():
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    grad = Image.new("RGBA", (S, S))
    gp = grad.load()
    for y in range(S):
        for x in range(S):
            t = (x + y) / (2 * S)
            gp[x, y] = (int(24 + 20 * t), int(112 - 30 * t), int(210 - 30 * t), 255)
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((40, 40, S - 40, S - 40), radius=210, fill=255)
    img.paste(grad, (0, 0), mask)
    d = ImageDraw.Draw(img)
    # 剪贴板板身
    d.rounded_rectangle((250, 215, 774, 835), radius=70, fill=(255, 255, 255, 255))
    # 夹子
    d.rounded_rectangle((380, 160, 644, 300), radius=45, fill=(255, 183, 43, 255))
    d.rounded_rectangle((440, 195, 584, 235), radius=20, fill=(255, 255, 255, 255))
    # 单据行
    for i, y in enumerate((400, 520, 640)):
        d.ellipse((310, y - 30, 370, y + 30), fill=(22, 140, 90, 255) if i < 2 else (200, 205, 215, 255))
        d.rounded_rectangle((405, y - 22, 700 - i * 60, y + 22), radius=22, fill=(205, 214, 230, 255))
    # 勾
    d.line([(318, 400), (338, 424), (366, 380)], fill=(255, 255, 255, 255), width=12, joint="curve")
    d.line([(318, 520), (338, 544), (366, 500)], fill=(255, 255, 255, 255), width=12, joint="curve")
    return img


img = build()
sizes = [16, 24, 32, 48, 64, 128, 256]
img.resize((256, 256), Image.LANCZOS).save(
    os.path.join(OUT, "favicon.ico"), format="ICO",
    sizes=[(s, s) for s in sizes])
img.resize((512, 512), Image.LANCZOS).save(os.path.join(OUT, "app-icon.png"))
print("ok")
