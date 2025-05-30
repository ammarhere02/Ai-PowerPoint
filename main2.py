import os
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from openai import OpenAI
from PIL import Image, ImageDraw, ImageFont
import requests
from io import BytesIO
import json
import random
from dotenv import load_dotenv  # Add this import
import shutil  # Add this import at the top with the other imports
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
# Load environment variables
load_dotenv()

class PowerPointProcessor:
    def __init__(self, openai_api_key=None, unsplash_api_key=None):
        """Initialize the processor with optional API keys"""
        self.openai_client = None
        self.unsplash_api_key = unsplash_api_key or os.getenv("UNSPLASH_API_KEY")

        # Use environment variable if no key provided
        openai_api_key = openai_api_key or os.getenv("OPEN_AI")
        if openai_api_key:
            self.openai_client = OpenAI(api_key=openai_api_key)

    def extract_text_from_pptx(self, file_path):
        """Extract text content from PowerPoint slides"""
        prs = Presentation(file_path)
        slides_content = []

        for i, slide in enumerate(prs.slides):
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text.append(shape.text.strip())

            slides_content.append({
                'slide_number': i + 1,
                'content': slide_text
            })

        return slides_content

    def _extract_slide_texts(self, prs):
        """Extract text content from PowerPoint slides with better handling of long text"""
        texts = []
        for i, slide in enumerate(prs.slides):
            slide_text = ""
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    # Append text with a newline separator
                    slide_text += shape.text.strip() + "\n\n"
            texts.append({
                'slide_number': i + 1,
                'content': slide_text.strip()
            })
        return texts

    def structure_content(self, slides_content):
        """Convert slides to properly structured format with headings, tables, and bullet points"""
        structured_slides = []

        for slide in slides_content:
            slide_num = slide['slide_number']
            content = slide['content']

            structured_content = self._analyze_and_structure_text(content)

            structured_slides.append({
                'slide_number': slide_num,
                'structured_content': structured_content,
                'needs_image': self._determine_image_need(structured_content)
            })

        return structured_slides

    def _analyze_and_structure_text(self, content):
        """Analyze text content and structure it with appropriate formatting"""
        if not content:
            return {'type': 'empty', 'title': 'Empty Slide', 'data': []}

        combined_text = '\n'.join(content)

        # Detect if content looks like tabular data
        if self._is_tabular_data(combined_text):
            return {
                'type': 'table',
                'title': self._extract_title(content[0] if content else 'Data Table'),
                'data': self._parse_table_data(combined_text)
            }

        # Detect if content has bullet points or lists
        elif self._has_bullet_points(combined_text):
            return {
                'type': 'bullet_list',
                'title': self._extract_title(content[0] if content else 'Key Points'),
                'points': self._extract_bullet_points(combined_text)
            }

        # Default to structured text with heading
        else:
            return {
                'type': 'structured_text',
                'title': self._extract_title(content[0] if content else 'Content'),
                'content': self._structure_paragraphs(combined_text)
            }

    def _is_tabular_data(self, text):
        """Check if text contains tabular data patterns"""
        patterns = [
            r'\t.*\t',  # Tab-separated
            r'\|.*\|',  # Pipe-separated
            r':\s*\d+',  # Colon-number pairs
            r'\d+\.\d+%',  # Percentages
        ]
        return any(re.search(pattern, text) for pattern in patterns)

    def _has_bullet_points(self, text):
        """Check if text has bullet point patterns"""
        patterns = [
            r'^[\•\-\*]\s',  # Bullet symbols
            r'^\d+\.\s',     # Numbered lists
            r'^[a-zA-Z]\.\s' # Letter lists
        ]
        lines = text.split('\n')
        if not lines:
            return False
        bullet_lines = sum(1 for line in lines if any(re.match(pattern, line.strip()) for pattern in patterns))
        return bullet_lines > len(lines) * 0.3  # 30% of lines are bullet points

    def _extract_title(self, first_line):
        """Extract title from first line"""
        if not first_line:
            return "Slide Content"
        if len(first_line) < 80 and not first_line.endswith('.'):
            return first_line
        return "Key Points"

    def _parse_table_data(self, text):
        """Parse text into table format"""
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        table_data = []

        for line in lines[1:]:  # Skip title line
            # Try different separators
            if '\t' in line:
                row = [cell.strip() for cell in line.split('\t')]
            elif '|' in line:
                row = [cell.strip() for cell in line.split('|')]
            else:
                # Try to split on multiple spaces or colons
                row = re.split(r'\s{2,}|:', line)
                row = [cell.strip() for cell in row if cell.strip()]

            if len(row) > 1:
                table_data.append(row)

        return table_data

    def _extract_bullet_points(self, text):
        """Extract bullet points from text"""
        lines = text.split('\n')
        points = []

        for line in lines:
            line = line.strip()
            if re.match(r'^[\•\-\*]\s', line):
                points.append(re.sub(r'^[\•\-\*]\s', '', line))
            elif re.match(r'^\d+\.\s', line):
                points.append(re.sub(r'^\d+\.\s', '', line))
            elif re.match(r'^[a-zA-Z]\.\s', line):
                points.append(re.sub(r'^[a-zA-Z]\.\s', '', line))
            elif line and not any(re.match(r'^[\•\-\*]\s|^\d+\.\s|^[a-zA-Z]\.\s', l.strip()) for l in lines[:max(0, lines.index(line))]):
                # This might be a title or header - skip
                continue
            elif line:
                points.append(line)

        return points if points else ["Content point"]

    def _structure_paragraphs(self, text):
        """Structure text into paragraphs with better handling of long text"""
        # First try to split by double newlines (paragraph breaks)
        paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]

        # If no paragraph breaks found, try single newlines
        if not paragraphs:
            paragraphs = [p.strip() for p in text.split('\n') if p.strip()]

        # If still no breaks, handle as a single paragraph
        if not paragraphs:
            paragraphs = [text]

        # Ensure we don't lose any content by truncating
        return paragraphs

    def _determine_image_need(self, structured_content):
        """Determine if a slide would benefit from an image"""
        image_keywords = [
            'process', 'workflow', 'diagram', 'chart', 'graph', 'visual',
            'example', 'comparison', 'analysis', 'data', 'statistics',
            'architecture', 'design', 'model', 'framework', 'structure',
            'timeline', 'roadmap', 'overview', 'summary', 'business',
            'strategy', 'growth', 'performance', 'results', 'metrics'
        ]

        content_text = str(structured_content).lower()
        keyword_matches = sum(1 for keyword in image_keywords if keyword in content_text)

        # Also consider content type
        if structured_content['type'] in ['table', 'bullet_list']:
            return True

        return keyword_matches >= 2

    def generate_images_for_slides(self, structured_slides):
        """Generate images for slides that need them"""
        for slide in structured_slides:
            if slide['needs_image']:
                image_prompt = self._create_image_prompt(slide['structured_content'])
                image_path = self._generate_image(image_prompt, slide['slide_number'])
                slide['image_path'] = image_path
            else:
                slide['image_path'] = None

        return structured_slides

    def _create_image_prompt(self, content):
        """Create a prompt for image generation based on content"""
        content_type = content['type']
        title = content.get('title', 'Slide Content')

        if content_type == 'table':
            return "business data analytics dashboard charts graphs professional"
        elif content_type == 'bullet_list':
            return "business presentation infographic professional corporate strategy"
        else:
            return "business professional office meeting presentation corporate"

    def _generate_image(self, prompt, slide_number):
        """Generate image using available APIs or create a placeholder"""
        # Create pics directory if it doesn't exist
        pics_dir = "pics"
        if not os.path.exists(pics_dir):
            os.makedirs(pics_dir)

        try:
            # Try Unsplash first
            if self.unsplash_api_key:
                return self._get_unsplash_image(prompt, slide_number, pics_dir)

            # Try OpenAI DALL-E (new API)
            elif self.openai_client:
                return self._generate_openai_image(prompt, slide_number, pics_dir)

            else:
                # Create placeholder image
                return self._create_placeholder_image(prompt, slide_number, pics_dir)

        except Exception as e:
            print(f"Error generating image for slide {slide_number}: {e}")
            return self._create_placeholder_image(prompt, slide_number, pics_dir)

    def _get_unsplash_image(self, prompt, slide_number, pics_dir):
        """Get image from Unsplash API"""
        try:
            url = "https://api.unsplash.com/search/photos"
            headers = {"Authorization": f"Client-ID {self.unsplash_api_key}"}
            params = {
                "query": prompt,
                "per_page": 10,
                "orientation": "landscape"
            }

            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()

            data = response.json()

            if data['results']:
                # Get a random image from results
                image_data = random.choice(data['results'])
                image_url = image_data['urls']['regular']

                # Download and save image
                img_response = requests.get(image_url)
                img_response.raise_for_status()

                image = Image.open(BytesIO(img_response.content))

                # Resize image for presentation
                image = image.resize((800, 600), Image.Resampling.LANCZOS)

                image_path = os.path.join(pics_dir, f"unsplash_image_slide_{slide_number}.jpg")
                image.save(image_path, "JPEG", quality=85)
                return image_path

            else:
                return self._create_placeholder_image(prompt, slide_number, pics_dir)

        except Exception as e:
            print(f"Unsplash API error: {e}")
            return self._create_placeholder_image(prompt, slide_number, pics_dir)

    def _generate_openai_image(self, prompt, slide_number, pics_dir):
        """Generate image using OpenAI DALL-E (new API)"""
        try:
            full_prompt = f"Create a professional business presentation image for: {prompt}. Modern, clean design suitable for corporate presentation. High quality, professional photography style."

            response = self.openai_client.images.generate(
                model="dall-e-3",
                prompt=full_prompt,
                size="1024x1024",
                quality="standard",
                n=1,
            )

            image_url = response.data[0].url

            # Download and save image
            img_response = requests.get(image_url)
            img_response.raise_for_status()

            image = Image.open(BytesIO(img_response.content))

            # Resize for presentation
            image = image.resize((800, 600), Image.Resampling.LANCZOS)

            image_path = os.path.join(pics_dir, f"openai_image_slide_{slide_number}.jpg")
            image.save(image_path, "JPEG", quality=85)
            return image_path

        except Exception as e:
            print(f"OpenAI API error: {e}")
            return self._create_placeholder_image(prompt, slide_number, pics_dir)

    def _create_placeholder_image(self, prompt, slide_number, pics_dir):
        """Create a placeholder image when APIs are not available"""
        img = Image.new('RGB', (800, 600), color='#f8f9fa')
        draw = ImageDraw.Draw(img)

        # Draw a simple border
        draw.rectangle([10, 10, 790, 590], outline='#dee2e6', width=2)

        try:
            font_large = ImageFont.truetype("arial.ttf", 24)
            font_small = ImageFont.truetype("arial.ttf", 16)
        except:
            font_large = ImageFont.load_default()
            font_small = ImageFont.load_default()

        # Add title
        title_text = f"Slide {slide_number} Image"

        # Calculate text position for centering
        bbox = draw.textbbox((0, 0), title_text, font=font_large)
        text_width = bbox[2] - bbox[0]
        x = (800 - text_width) // 2

        draw.text((x, 250), title_text, fill='#495057', font=font_large)

        # Add description
        desc_text = f"Image placeholder for: {prompt[:50]}..."
        bbox = draw.textbbox((0, 0), desc_text, font=font_small)
        text_width = bbox[2] - bbox[0]
        x = (800 - text_width) // 2

        draw.text((x, 300), desc_text, fill='#6c757d', font=font_small)

        # Add simple graphic element
        draw.ellipse([350, 350, 450, 450], fill='#007bff', outline='#0056b3', width=2)

        image_path = os.path.join(pics_dir, f"placeholder_image_slide_{slide_number}.png")
        img.save(image_path)
        return image_path

    def create_beautiful_presentation(self, structured_slides, output_path="beautiful_presentation.pptx", template_style="modern_blue"):
        """Create a beautiful presentation with structured content and images"""
        prs = Presentation()

        # Remove default slide if it exists
        if len(prs.slides) > 0:
            xml_slides = prs.slides._sldIdLst
            if len(xml_slides) > 0:
                xml_slides.remove(xml_slides[0])

        # Add title slide
        self._create_title_slide(prs, structured_slides, template_style)

        for slide_data in structured_slides:
            slide = self._create_styled_slide(prs, slide_data, template_style)

        prs.save(output_path)
        return output_path

    def _get_template_colors(self, template_style):
        """Get color scheme for different templates"""
        templates = {
            "modern_blue": {
                "primary": RGBColor(31, 73, 125),      # Dark blue
                "secondary": RGBColor(79, 129, 189),    # Medium blue
                "accent": RGBColor(149, 179, 215),      # Light blue
                "background": RGBColor(248, 249, 250),   # Light gray
                "text": RGBColor(64, 64, 64),           # Dark gray
                "white": RGBColor(255, 255, 255)        # White
            },
            "corporate_green": {
                "primary": RGBColor(34, 139, 34),       # Forest green
                "secondary": RGBColor(60, 179, 113),     # Medium sea green
                "accent": RGBColor(144, 238, 144),       # Light green
                "background": RGBColor(248, 255, 248),   # Very light green
                "text": RGBColor(64, 64, 64),           # Dark gray
                "white": RGBColor(255, 255, 255)        # White
            },
            "professional_gray": {
                "primary": RGBColor(64, 64, 64),        # Dark gray
                "secondary": RGBColor(128, 128, 128),    # Medium gray
                "accent": RGBColor(192, 192, 192),       # Light gray
                "background": RGBColor(248, 248, 248),   # Very light gray
                "text": RGBColor(32, 32, 32),           # Very dark gray
                "white": RGBColor(255, 255, 255)        # White
            }
        }
        return templates.get(template_style, templates["modern_blue"])

    def _create_title_slide(self, prs, structured_slides, template_style):
        """Create a professional title slide"""
        colors = self._get_template_colors(template_style)
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)

        # Background gradient effect
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = colors["background"]

        # Add decorative header bar
        header_shape = slide.shapes.add_shape(
            1, Inches(0), Inches(0), Inches(10), Inches(1.5)  # Rectangle shape
        )
        header_fill = header_shape.fill
        header_fill.solid()
        header_fill.fore_color.rgb = colors["primary"]

        # Main title
        title_shape = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
        title_frame = title_shape.text_frame
        title_frame.text = "Enhanced Business Presentation"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(36)
        title_para.font.bold = True
        title_para.font.color.rgb = colors["primary"]
        title_para.alignment = PP_ALIGN.CENTER

        # Subtitle
        subtitle_shape = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(8), Inches(1))
        subtitle_frame = subtitle_shape.text_frame
        subtitle_frame.text = f"Professional Analysis • {len(structured_slides)} Key Insights"
        subtitle_para = subtitle_frame.paragraphs[0]
        subtitle_para.font.size = Pt(18)
        subtitle_para.font.color.rgb = colors["secondary"]
        subtitle_para.alignment = PP_ALIGN.CENTER
    def _create_styled_slide(self, prs, slide_data, template_style="modern_blue"):
        """Create a styled slide based on content type"""
        colors = self._get_template_colors(template_style)
        content = slide_data['structured_content']
        has_image = slide_data['image_path'] is not None

        # Use blank layout for full control
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)

        # Set slide background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = colors["background"]

        # Add decorative header bar
        header_shape = slide.shapes.add_shape(
            1, Inches(0), Inches(0), Inches(10), Inches(0.3)  # Rectangle
        )
        header_fill = header_shape.fill
        header_fill.solid()
        header_fill.fore_color.rgb = colors["primary"]

        if content['type'] == 'table':
            self._add_table_slide_template(slide, content, slide_data['image_path'], colors)
        elif content['type'] == 'bullet_list':
            self._add_bullet_slide_template(slide, content, slide_data['image_path'], colors)
        else:
            self._add_text_slide_template(slide, content, slide_data['image_path'], colors)

        return slide

    def _add_table_slide_template(self, slide, content, image_path, colors):
        """Add table content to slide with proper template layout"""
        # Title with background
        title_bg = slide.shapes.add_shape(
            1, Inches(0.5), Inches(0.8), Inches(9), Inches(0.8)  # Rectangle
        )
        title_bg_fill = title_bg.fill
        title_bg_fill.solid()
        title_bg_fill.fore_color.rgb = colors["white"]

        # Add subtle border to title
        title_bg.line.color.rgb = colors["accent"]
        title_bg.line.width = Pt(1)

        # Title text
        title_shape = slide.shapes.add_textbox(Inches(0.7), Inches(1), Inches(8.6), Inches(0.6))
        title_frame = title_shape.text_frame
        title_frame.text = content['title']
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(20)
        title_para.font.bold = True
        title_para.font.color.rgb = colors["primary"]

        # Determine layout based on image presence
        if image_path and os.path.exists(image_path):
            # Two-column layout: table on left, image on right
            table_left = Inches(0.5)
            table_width = Inches(4.5)
            image_left = Inches(5.5)
            image_width = Inches(3.5)
            image_height = Inches(3.5)
        else:
            # Full-width table
            table_left = Inches(0.5)
            table_width = Inches(9)

        # Add table if data exists
        if content['data'] and len(content['data']) > 0:
            rows = min(len(content['data']), 8)  # Limit rows to fit screen
            cols = min(max(len(row) for row in content['data']) if content['data'] else 2, 5)  # Limit columns

            table_shape = slide.shapes.add_table(rows, cols, table_left, Inches(2), table_width, Inches(4))
            table = table_shape.table

            # Style table with template colors
            for i, row_data in enumerate(content['data'][:rows]):
                for j, cell_data in enumerate(row_data[:cols]):
                    if j < cols:
                        cell = table.cell(i, j)
                        cell.text = str(cell_data)[:50]  # Limit text length

                        # Style cell
                        cell_para = cell.text_frame.paragraphs[0]
                        cell_para.font.size = Pt(10)

                        # Header row styling
                        if i == 0:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = colors["primary"]
                            cell_para.font.color.rgb = colors["white"]
                            cell_para.font.bold = True
                        else:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = colors["white"]
                            cell_para.font.color.rgb = colors["text"]

        # Add image if available
        if image_path and os.path.exists(image_path):
            try:
                slide.shapes.add_picture(image_path, image_left, Inches(2), image_width, image_height)
            except Exception as e:
                print(f"Error adding image: {e}")

    def _add_bullet_slide_template(self, slide, content, image_path, colors):
        """Add bullet point content to slide with proper template layout"""
        # Title with background
        title_bg = slide.shapes.add_shape(
            1, Inches(0.5), Inches(0.8), Inches(9), Inches(0.8)  # Rectangle
        )
        title_bg_fill = title_bg.fill
        title_bg_fill.solid()
        title_bg_fill.fore_color.rgb = colors["white"]

        # Add subtle border to title
        title_bg.line.color.rgb = colors["accent"]
        title_bg.line.width = Pt(1)

        # Title text
        title_shape = slide.shapes.add_textbox(Inches(0.7), Inches(1), Inches(8.6), Inches(0.6))
        title_frame = title_shape.text_frame
        title_frame.text = content['title']
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(20)
        title_para.font.bold = True
        title_para.font.color.rgb = colors["primary"]

        # Determine layout based on image presence
        if image_path and os.path.exists(image_path):
            # Two-column layout: text on left, image on right
            text_left = Inches(0.5)
            text_width = Inches(4.5)
            image_left = Inches(5.5)
            image_width = Inches(3.5)
            image_height = Inches(3.5)
        else:
            # Full-width text
            text_left = Inches(0.5)
            text_width = Inches(9)

        # Content background
        content_bg = slide.shapes.add_shape(
            1, text_left, Inches(2), text_width, Inches(4)  # Rectangle
        )
        content_bg_fill = content_bg.fill
        content_bg_fill.solid()
        content_bg_fill.fore_color.rgb = colors["white"]
        content_bg.line.color.rgb = colors["accent"]
        content_bg.line.width = Pt(1)

        text_shape = slide.shapes.add_textbox(text_left + Inches(0.2), Inches(2.2), text_width - Inches(0.4), Inches(3.6))
        text_frame = text_shape.text_frame
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_right = Inches(0.1)
        text_frame.word_wrap = True
        text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        points = content.get('points', ['No content available'])
        # Don't limit to a specific number of points - use auto-sizing instead
        for i, point in enumerate(points):
            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()

            # Don't arbitrarily truncate text
            p.text = f"• {str(point)}"
            p.font.size = Pt(11) if not image_path else Pt(10)  # Smaller text if image present
            p.font.color.rgb = colors["text"]
            p.space_after = Pt(6) if not image_path else Pt(4)  # Less space if image present


        # Add image if available
        if image_path and os.path.exists(image_path):
            try:
                slide.shapes.add_picture(image_path, image_left, Inches(2), image_width, image_height)
            except Exception as e:
                print(f"Error adding image: {e}")

    def _add_text_slide_template(self, slide, content, image_path, colors):
        """Add structured text content to slide with proper template layout"""
        # Title with background
        title_bg = slide.shapes.add_shape(
            1, Inches(0.5), Inches(0.8), Inches(9), Inches(0.8)  # Rectangle
        )
        title_bg_fill = title_bg.fill
        title_bg_fill.solid()
        title_bg_fill.fore_color.rgb = colors["white"]

        # Add subtle border to title
        title_bg.line.color.rgb = colors["accent"]
        title_bg.line.width = Pt(1)

        # Title text
        title_shape = slide.shapes.add_textbox(Inches(0.7), Inches(1), Inches(8.6), Inches(0.6))
        title_frame = title_shape.text_frame
        title_frame.text = content['title']
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(20)
        title_para.font.bold = True
        title_para.font.color.rgb = colors["primary"]

        # Determine layout based on image presence

        if image_path and os.path.exists(image_path):
            # Two-column layout: text on left, image on right
            text_left = Inches(0.5)
            text_width = Inches(4.5)
            image_left = Inches(5.5)
            image_width = Inches(3.5)
            image_height = Inches(3.5)
        else:
            # Full-width text
            text_left = Inches(0.5)
            text_width = Inches(9)

        # Content background
        content_bg = slide.shapes.add_shape(
            1, text_left, Inches(2), text_width, Inches(4)  # Rectangle
        )
        content_bg_fill = content_bg.fill
        content_bg_fill.solid()
        content_bg_fill.fore_color.rgb = colors["white"]
        content_bg.line.color.rgb = colors["accent"]
        content_bg.line.width = Pt(1)

        # Add content paragraphs - UPDATED CODE HERE
        text_shape = slide.shapes.add_textbox(text_left + Inches(0.2), Inches(2.2), text_width - Inches(0.4), Inches(3.6))
        text_frame = text_shape.text_frame
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_right = Inches(0.1)
        text_frame.word_wrap = True

        # Use automatic fitting to shape
        text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        paragraphs = content.get('content', ['No content available'])

        # Join paragraphs if image is present to save space
        if image_path and os.path.exists(image_path) and len(paragraphs) > 2:
            # Use first paragraph as is
            p = text_frame.paragraphs[0]
            p.text = str(paragraphs[0])
            p.font.size = Pt(11)
            p.font.color.rgb = colors["text"]
            p.space_after = Pt(8)

            # Join remaining paragraphs with bullets
            for i in range(1, len(paragraphs)):
                p = text_frame.add_paragraph()
                p.text = "• " + str(paragraphs[i])
                p.font.size = Pt(10)  # Slightly smaller font
                p.font.color.rgb = colors["text"]
                p.space_after = Pt(4)  # Less space between items
        else:
            # Use standard paragraph layout for slides without images or with few paragraphs
            for i, paragraph in enumerate(paragraphs):
                if i == 0:
                    p = text_frame.paragraphs[0]
                else:
                    p = text_frame.add_paragraph()

                p.text = str(paragraph)
                p.font.size = Pt(11)
                p.font.color.rgb = colors["text"]
                p.space_after = Pt(8)

        # Add image if available
        if image_path and os.path.exists(image_path):
            try:
                slide.shapes.add_picture(image_path, image_left, Inches(2), image_width, image_height)
            except Exception as e:
                print(f"Error adding image: {e}")

    def process_presentation(self, input_pptx_path, output_pptx_path="enhanced_presentation.pptx", template_style="modern_blue"):
        """Main method to process the entire presentation"""
        print("Step 1: Extracting content from PowerPoint...")
        slides_content = self.extract_text_from_pptx(input_pptx_path)

        print("Step 2: Structuring content...")
        structured_slides = self.structure_content(slides_content)

        print("Step 3: Generating images for relevant slides...")
        slides_with_images = self.generate_images_for_slides(structured_slides)

        print("Step 4: Creating beautiful presentation...")
        final_presentation_path = self.create_beautiful_presentation(slides_with_images, output_pptx_path, template_style)

        print(f"✅ Enhanced presentation created: {final_presentation_path}")

        # Print summary
        image_slides = sum(1 for slide in slides_with_images if slide['needs_image'])
        print(f"📊 Processed {len(slides_content)} slides")
        print(f"🖼️  Generated images for {image_slides} slides")
        print(f"🎨 Applied template: {template_style}")

        # Delete the pics folder after processing is complete
        if os.path.exists("pics"):
            try:
                shutil.rmtree("pics")
                print("🗑️  Temporary image files cleaned up")
            except Exception as e:
                print(f"⚠️ Warning: Could not remove temporary images: {e}")

        return final_presentation_path


# Usage example
if __name__ == "__main__":
    # Initialize processor with API keys from environment variables
    processor = PowerPointProcessor()

    # Process presentation
    input_file = "orignal_trimmed.pptx"  # Your input file
    output_file = "enhanced_presentation.pptx"

    if os.path.exists(input_file):
        try:
            result = processor.process_presentation(input_file, output_file)
            print(f"\n🎉 Success! Enhanced presentation saved as: {result}")
        except Exception as e:
            print(f"❌ Error processing presentation: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"❌ Input file '{input_file}' not found!")
        print("Creating a sample presentation for testing...")

        # Create a sample presentation for testing
        sample_prs = Presentation()
        slide = sample_prs.slides[0]
        slide.shapes.title.text = "Sample Business Report"
        content = slide.placeholders[1].text_frame
        content.text = "Key Performance Indicators:\n• Revenue increased by 25%\n• Customer satisfaction: 92%\n• Market share growth: 15%"

        sample_prs.save("sample_presentation.pptx")
        print("Sample presentation created: sample_presentation.pptx")

        # Process the sample
        try:
            result = processor.process_presentation("sample_presentation.pptx", "enhanced_sample.pptx")
            print(f"\n🎉 Sample processed successfully: {result}")
        except Exception as e:
            print(f"❌ Error processing sample: {e}")
            import traceback
            traceback.print_exc()