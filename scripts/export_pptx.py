from pathlib import Path
import re

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "ppt-outline" / "slides_full_content.md"
OUTPUT = ROOT / "deliverables" / "创策云UniPlanAI课堂汇报.pptx"
SCRIPT_OUTPUT = ROOT / "deliverables" / "创策云UniPlanAI汇报讲稿.md"

BLUE = RGBColor(23, 111, 193)
DARK = RGBColor(18, 36, 60)
MID = RGBColor(75, 97, 120)
LIGHT_BG = RGBColor(239, 246, 252)
WHITE = RGBColor(255, 255, 255)


def parse_slides(markdown):
    pattern = re.compile(r"^## 第\s*(\d+)\s*页[：:](.+)$", re.MULTILINE)
    matches = list(pattern.finditer(markdown))
    slides = []
    for idx, match in enumerate(matches):
        start = match.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(markdown)
        block = markdown[start:end].strip()
        title = match.group(2).strip()
        bullets = []
        visual = ""
        speaker = ""
        mode = None
        for raw in block.splitlines():
            line = raw.strip()
            if not line:
                continue
            if line.startswith("页面文字"):
                mode = "bullets"
                continue
            if line.startswith("图示建议"):
                mode = "visual"
                visual = line.split("：", 1)[-1].strip() if "：" in line else ""
                continue
            if line.startswith("汇报讲稿"):
                mode = "speaker"
                speaker = line.split("：", 1)[-1].strip() if "：" in line else ""
                continue
            if mode == "bullets" and line.startswith("- "):
                bullets.append(line[2:].strip())
            elif mode == "visual":
                visual = (visual + " " + line).strip()
            elif mode == "speaker":
                speaker = (speaker + " " + line).strip()
        slides.append(
            {
                "number": int(match.group(1)),
                "title": title,
                "bullets": bullets[:5],
                "visual": visual,
                "speaker": speaker,
            }
        )
    return slides


def add_textbox(slide, left, top, width, height, text, size=18, bold=False, color=DARK, align=None):
    shape = slide.shapes.add_textbox(left, top, width, height)
    frame = shape.text_frame
    frame.clear()
    paragraph = frame.paragraphs[0]
    if align is not None:
        paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return shape


def add_bullets(slide, bullets):
    box = slide.shapes.add_textbox(Inches(0.9), Inches(1.65), Inches(7.2), Inches(4.35))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    for index, bullet in enumerate(bullets):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = bullet
        paragraph.level = 0
        paragraph.space_after = Pt(8)
        for run in paragraph.runs:
            run.font.name = "Microsoft YaHei"
            run.font.size = Pt(21)
            run.font.color.rgb = DARK


def add_visual_card(slide, text):
    left, top, width, height = Inches(8.45), Inches(1.5), Inches(3.9), Inches(4.45)
    card = slide.shapes.add_shape(1, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_BG
    card.line.color.rgb = RGBColor(196, 218, 238)
    add_textbox(slide, left + Inches(0.25), top + Inches(0.35), width - Inches(0.5), Inches(0.4), "图示建议", 16, True, BLUE)
    add_textbox(slide, left + Inches(0.25), top + Inches(0.9), width - Inches(0.5), height - Inches(1.2), text or "可放置原型截图或流程图", 15, False, MID)


def apply_background(slide):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = WHITE
    top_bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.22))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = BLUE
    top_bar.line.fill.background()
    accent = slide.shapes.add_shape(1, Inches(0), Inches(6.95), Inches(13.333), Inches(0.25))
    accent.fill.solid()
    accent.fill.fore_color.rgb = LIGHT_BG
    accent.line.fill.background()


def build_ppt(slides_data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for index, item in enumerate(slides_data):
        slide = prs.slides.add_slide(blank)
        apply_background(slide)
        if index == 0:
            add_textbox(slide, Inches(0.85), Inches(1.35), Inches(11.6), Inches(0.65), "创策云 UniPlan AI", 36, True, BLUE)
            add_textbox(slide, Inches(0.9), Inches(2.25), Inches(10.8), Inches(0.5), "面向大学生创新创业项目的 AI 协作与商业计划书平台", 24, False, DARK)
            for i, bullet in enumerate(item["bullets"][2:5]):
                add_textbox(slide, Inches(0.95), Inches(3.25 + i * 0.5), Inches(9.8), Inches(0.35), bullet, 18, False, MID)
            add_visual_card(slide, item["visual"])
        else:
            add_textbox(slide, Inches(0.65), Inches(0.55), Inches(11.7), Inches(0.65), item["title"], 28, True, DARK)
            add_bullets(slide, item["bullets"])
            add_visual_card(slide, item["visual"])
            add_textbox(slide, Inches(11.65), Inches(6.9), Inches(1.1), Inches(0.3), f"{item['number']:02d}", 12, True, BLUE, PP_ALIGN.RIGHT)
    return prs


def export_speaker_notes(slides_data):
    lines = ["# 创策云 UniPlan AI 汇报讲稿", "", "建议总时长：6—8 分钟。", ""]
    for item in slides_data:
        lines.append(f"## 第 {item['number']} 页：{item['title']}")
        lines.append("")
        lines.append(item["speaker"] or "【待补充讲稿】")
        lines.append("")
    SCRIPT_OUTPUT.write_text("\n".join(lines), encoding="utf-8")


def main():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    markdown = SOURCE.read_text(encoding="utf-8")
    slides_data = parse_slides(markdown)
    if not slides_data:
        raise RuntimeError("未能从 slides_full_content.md 解析到 PPT 页面")
    prs = build_ppt(slides_data)
    prs.save(OUTPUT)
    export_speaker_notes(slides_data)
    print(f"exported {OUTPUT}")
    print(f"exported {SCRIPT_OUTPUT}")


if __name__ == "__main__":
    main()
