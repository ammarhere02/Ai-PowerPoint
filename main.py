import os
import requests
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from dotenv import load_dotenv
from openai import OpenAI
import argparse
import logging
import time

# === Hardcoded Template ===
template_path = "BNW.pptx"                # Your template file
output_path = "output_bnw_result.pptx"    # Output file

prs = Presentation(template_path)

# Choose a layout from your template (0=Title Slide, 1=Title+Content, etc.)
slide_layout = prs.slide_layouts[1]

slide = prs.slides.add_slide(slide_layout)
slide.shapes.title.text = "This uses the full template design!"
slide.placeholders[1].text = "No code is overriding the template's style."

# Add more slides as needed, each time choosing `prs.slide_layouts[x]`
prs.save(output_path)


# Setup
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

load_dotenv()
OPEN_AI_KEY = os.getenv("OPEN_AI")
UNSPLASH_API_KEY = os.getenv("UNSPLASH_API_KEY")
if not OPEN_AI_KEY:
    raise ValueError("Missing OpenAI Key")
client = OpenAI(api_key=OPEN_AI_KEY)


def prettify_text(text, max_bullet_len=80):
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if len(lines) == 1:
        return 'title', lines[0]
    elif all(len(l) > max_bullet_len for l in lines):
        return 'body', [lines[0]] + [l for l in lines[1:]]
    else:
        return 'bullets', lines

def set_slide_background(slide, rgb=(255, 253, 208)):  # Example: cream color
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb)

def create_slide(new_prs, kind, content, force_bg_solid=False):
    # Use layouts from the CREAM.pptx template.
    if kind == 'title':
        layout = new_prs.slide_layouts[0]  # Usually the title slide
        slide = new_prs.slides.add_slide(layout)
        title = slide.shapes.title
        title.text = content
        title.text_frame.paragraphs[0].font.size = Pt(38)
        title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        title.text_frame.paragraphs[0].font.bold = True
        if force_bg_solid:
            set_slide_background(slide)
    elif kind == 'body':
        layout = new_prs.slide_layouts[5] if len(new_prs.slide_layouts) > 5 else new_prs.slide_layouts[1]
        slide = new_prs.slides.add_slide(layout)
        left, top, width, height = Inches(0.7), Inches(1.3), Inches(6), Inches(4.2)
        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.word_wrap = True
        p = tf.add_paragraph()
        p.text = content[0]
        p.font.size = Pt(32)
        p.font.bold = True
        for item in content[1:]:
            para = tf.add_paragraph()
            para.text = item
            para.level = 1
            para.font.size = Pt(24)
        if force_bg_solid:
            set_slide_background(slide)
    elif kind == 'bullets':
        layout = new_prs.slide_layouts[1]
        slide = new_prs.slides.add_slide(layout)
        title = slide.shapes.title
        title.text = "Key Points"
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame
        tf.clear()
        for line in content:
            p = tf.add_paragraph()
            p.text = line
            p.level = 0
            p.font.size = Pt(24)
        if force_bg_solid:
            set_slide_background(slide)
    else:
        layout = new_prs.slide_layouts[0]
        slide = new_prs.slides.add_slide(layout)
        if force_bg_solid:
            set_slide_background(slide)
    return slide

def extract_slide_texts(prs):
    texts = []
    for slide in prs.slides:
        text = "\n".join(
            shape.text.strip() for shape in slide.shapes if hasattr(shape, "text") and shape.text.strip()
        ) or "No text content"
        texts.append(text)
    return texts

def get_slide_scores(slide_texts):
    scores = []
    for i in range(0, len(slide_texts), 3):
        batch = slide_texts[i:i+3]
        prompt = "\n\n".join([f"Slide {i + j + 1}: {text}" for j, text in enumerate(batch)])
        try:
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{
                    "role": "user",
                    "content": "Rate each slide from 1 to 10 for importance. Respond with comma-separated numbers only.\n\n" + prompt
                }],
                temperature=0.2
            )
            score_str = response.choices[0].message.content.strip()
            batch_scores = [int(s.strip()) for s in score_str.replace('.', '').split(",") if s.strip().isdigit()]
            scores.extend(batch_scores + [5] * (len(batch) - len(batch_scores)))
        except Exception as e:
            logger.error(f"Error getting slide scores: {e}")
            scores.extend([5] * len(batch))
        time.sleep(2)
    return scores[:len(slide_texts)]

def get_unsplash_image_url(query):
    if not UNSPLASH_API_KEY:
        logger.warning("No Unsplash API key provided")
        return None
    headers = {"Authorization": f"Client-ID {UNSPLASH_API_KEY}"}
    params = {"query": query, "per_page": 1, "orientation": "landscape"}
    try:
        r = requests.get("https://api.unsplash.com/search/photos", headers=headers, params=params, timeout=9)
        r.raise_for_status()
        data = r.json()
        if data["results"]:
            return data["results"][0]["urls"]["regular"]
    except Exception as e:
        logger.error(f"Unsplash fetch error: {e}")
    return None

def add_image_no_overlap(slide, url, layout_type="bullets"):
    try:
        if not url:
            return
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        img_file = BytesIO(resp.content)
        if layout_type == "bullets":
            left = Inches(5.8)
            top = Inches(1.2)
            width = Inches(3.1)
            height = Inches(2.0)
        elif layout_type == "body":
            left = Inches(6.4)
            top = Inches(1.8)
            width = Inches(2.2)
            height = Inches(1.7)
        else:  # title or fallback
            left = Inches(3.4)
            top = Inches(3.4)
            width = Inches(3.5)
            height = Inches(2)
        slide.shapes.add_picture(img_file, left, top, width, height)
        logger.info("Image added with no overlap")
        img_file.close()
    except Exception as e:
        logger.error(f"Error adding image to slide: {e}")

def build_neat_pptx(input_path, output_path, keep_indices, slide_texts, force_bg_solid=False):
    original = Presentation(input_path)
    # Use CREAM.pptx as the theme/template for *all* new slides
    new_prs = Presentation(template_path)
    # Remove any template starter slides (usually 1 in a blank template)
    while len(new_prs.slides):
        rId = new_prs.slides._sldIdLst[0].rId
        new_prs.part.drop_rel(rId)
        new_prs.slides._sldIdLst.remove(new_prs.slides._sldIdLst[0])

    valid_max = len(original.slides) - 1
    filtered_indices = [idx for idx in keep_indices if 0 <= idx <= valid_max]

    main_query = "presentation"
    if slide_texts:
        first_slide = slide_texts[0]
        main_query = " ".join(first_slide.split()[:6]) or "presentation"

    img_url = get_unsplash_image_url(main_query)

    for i, idx in enumerate(filtered_indices):
        text = slide_texts[idx]
        kind, content = prettify_text(text)
        slide = create_slide(new_prs, kind, content, force_bg_solid=force_bg_solid)
        # Add the relevant image (from first slide) at every 3rd slide, and never overlapping text
        if img_url and (i % 3 == 0 or i == 0):
            add_image_no_overlap(slide, img_url, layout_type=kind)
    new_prs.save(output_path)
    logger.info(f"Created neat slides in {output_path}")

def process_pptx(input_path, output_path, score_threshold=6, force_bg_solid=False):
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file '{input_path}' does not exist")
    prs = Presentation(input_path)
    if len(prs.slides) == 0:
        raise ValueError("Input presentation contains no slides")
    slide_texts = extract_slide_texts(prs)
    scores = get_slide_scores(slide_texts)
    keep_indices = [i for i, s in enumerate(scores) if s >= score_threshold]
    if not keep_indices:
        keep_indices = list(range(len(prs.slides)))
    build_neat_pptx(input_path, output_path, keep_indices, slide_texts, force_bg_solid=force_bg_solid)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Neaten and enhance PowerPoint presentation.")
    parser.add_argument("input_file", help="Path to the input PowerPoint file")
    parser.add_argument("--output_file", help="Path to the output PowerPoint file", default=None)
    parser.add_argument("--threshold", type=int, default=6, help="Importance threshold (1-10)")
    parser.add_argument("--solid_bg", action='store_true', help="Enforce a solid cream color background on all slides")
    args = parser.parse_args()
    if not args.output_file:
        args.output_file = os.path.splitext(args.input_file)[0] + "_neat.pptx"
    try:
        process_pptx(args.input_file, args.output_file, args.threshold, force_bg_solid=args.solid_bg)
        logger.info(f"Neat slides generated at {args.output_file}")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")