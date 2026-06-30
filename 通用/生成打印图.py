# -*- coding: utf-8 -*-
"""
【固化版】把 `收据图片/` 下的 jpg 拼成 A4 纵向打印图，每页 3 张垂直均匀分布。

用法：
    python 每月租金生成/通用/生成打印图.py --year 2026 --month 6
    python 每月租金生成/通用/生成打印图.py --year 2026 --month 6 --village 变电站

输入：
    {ROOT}/{village}/{year}年{month}月/收据图片/*.jpg

输出：
    {ROOT}/{village}/{year}年{month}月/打印/打印_第N页.jpg
    {ROOT}/{village}/{year}年{month}月/打印/{year}年{month}月打印.pdf   # 全部页面合并版，方便一次性打印

规格（铁律，禁止自由发挥）：
- 纸张 A4 纵向 @ 300dpi（2480 × 3508 px）
- 页边距 100 px
- 每页固定 3 张，垂直均匀分布；最后一页不足 3 张就实际放几张（不补白卡）
- 排序：按文件名前缀房号升序
- JPEG 质量 92；PDF 由 PIL 直接多页保存（A4 原始像素，无重压缩）
"""
import argparse
import os
import sys
from PIL import Image


A4_W, A4_H = 2480, 3508
MARGIN = 100
USABLE_W = A4_W - MARGIN * 2
USABLE_H = A4_H - MARGIN * 2
PER_PAGE = 3


def sort_key(fn):
    try:
        return int(fn.split("房")[0])
    except Exception:
        return 9999


def main():
    ap = argparse.ArgumentParser(description="把收据 jpg 拼成 A4 打印图（每页 3 张）")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--month", type=int, required=True)
    ap.add_argument("--village", default="官田村")
    ap.add_argument("--root", default=None)
    args = ap.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 脚本在 {root}/通用/ 下，root 默认 = 父目录
    root = args.root or os.path.dirname(script_dir)

    month_dir = os.path.join(root, args.village, f"{args.year}年{args.month}月")
    src = os.path.join(month_dir, "收据图片")
    dst = os.path.join(month_dir, "打印")

    if not os.path.isdir(src):
        print(f"❌ 找不到收据图片目录：{src}（请先跑 通用/导出收据图片.py）", file=sys.stderr)
        sys.exit(1)

    files = sorted(
        [f for f in os.listdir(src) if f.lower().endswith((".jpg", ".jpeg", ".png"))],
        key=sort_key,
    )
    if not files:
        print(f"❌ {src} 下没有图片", file=sys.stderr)
        sys.exit(1)

    os.makedirs(dst, exist_ok=True)

    # 用第一张图算等比缩放后的高度
    sample = Image.open(os.path.join(src, files[0]))
    ratio = sample.size[0] / sample.size[1]
    target_w = USABLE_W
    target_h = int(target_w / ratio)

    pages = [files[i:i + PER_PAGE] for i in range(0, len(files), PER_PAGE)]

    print(f"📥 收据图片：{src}（{len(files)} 张）")
    print(f"📤 打印图：{dst}")
    print(f"📐 A4@300dpi {A4_W}×{A4_H}，每页 {PER_PAGE} 张，共 {len(pages)} 页")
    print()

    canvases = []  # 累积所有页面的 RGB Image，最后一次性打成 PDF
    for idx, page_files in enumerate(pages, 1):
        canvas = Image.new("RGB", (A4_W, A4_H), "white")
        n = len(page_files)
        gap = (USABLE_H - target_h * n) // (n + 1) if n > 0 else 0
        y = MARGIN + gap
        for fn in page_files:
            im = Image.open(os.path.join(src, fn)).convert("RGB").resize((target_w, target_h), Image.LANCZOS)
            canvas.paste(im, (MARGIN, y))
            y += target_h + gap
        out = os.path.join(dst, f"打印_第{idx}页.jpg")
        canvas.save(out, "JPEG", quality=92)
        canvases.append(canvas)
        rooms = ', '.join(f.split('_')[0] for f in page_files)
        print(f"  ✅ 第{idx}页 ({n} 张：{rooms})")

    # 合并 PDF：A4@300dpi → 物理尺寸 8.27×11.69 英寸；一份 PDF 全部打印
    pdf_path = os.path.join(dst, f"{args.year}年{args.month}月打印.pdf")
    canvases[0].save(
        pdf_path,
        "PDF",
        resolution=300.0,
        save_all=True,
        append_images=canvases[1:],
    )
    print(f"  📄 合并 PDF：{pdf_path}")

    print(f"\n🎉 共 {len(pages)} 页打印图（jpg + pdf），源图 {len(files)} 张")


if __name__ == "__main__":
    main()
