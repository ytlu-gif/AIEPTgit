from PIL import Image, ImageDraw, ImageFont

def make_icon():
    size = (256, 256)
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 繪製質感深藍背景圓形
    draw.ellipse([8, 8, 248, 248], fill=(26, 82, 118))
    # 繪製內圈白色圓框
    draw.ellipse([24, 24, 232, 232], outline=(255, 255, 255), width=8)

    try:
        font = ImageFont.truetype("arial.ttf", 130)
    except IOError:
        font = ImageFont.load_default()

    # 繪製字母 A
    draw.text((85, 40), "A", fill=(255, 255, 255), font=font)
    # 繪製音波弧線
    draw.arc([150, 150, 210, 210], start=270, end=90, fill=(241, 196, 15), width=12)

    # 儲存為 PNG 與 ICO 圖檔
    img.save("app_icon.png")
    img.save("app_icon.ico", format="ICO", sizes=[(256, 256), (64, 64), (32, 32), (16, 16)])
    print("成功生成 app_icon.ico 與 app_icon.png！")

if __name__ == "__main__":
    make_icon()