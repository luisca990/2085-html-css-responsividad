from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import os

# Create presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

def add_title_slide(prs, title, subtitle=""):
    """Add a title slide with modern design"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    
    # Background gradient effect using shapes
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0,
        prs.slide_width,
        prs.slide_height
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = RGBColor(15, 32, 60)
    bg_shape.line.fill.background()
    
    # Accent shape
    accent = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(2.5),
        Inches(9), Inches(2.5)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = RGBColor(41, 128, 185)
    accent.line.fill.background()
    accent.shadow.inherit = False
    
    # Title
    title_box = slide.shapes.add_textbox(
        Inches(1), Inches(2.8),
        Inches(8), Inches(1.5)
    )
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].font.size = Pt(54)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Subtitle
    if subtitle:
        subtitle_box = slide.shapes.add_textbox(
            Inches(1), Inches(4.5),
            Inches(8), Inches(0.8)
        )
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = subtitle
        subtitle_frame.paragraphs[0].font.size = Pt(24)
        subtitle_frame.paragraphs[0].font.color.rgb = RGBColor(236, 240, 241)
        subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    return slide

def add_content_slide(prs, title, content_items, layout_type="bullet"):
    """Add a content slide with visual elements"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0,
        prs.slide_width,
        prs.slide_height
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = RGBColor(236, 240, 241)
    bg_shape.line.fill.background()
    
    # Header bar
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0,
        prs.slide_width,
        Inches(1.2)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(41, 128, 185)
    header.line.fill.background()
    
    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(9), Inches(0.8)
    )
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].font.size = Pt(40)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    
    if layout_type == "bullet":
        # Content area with bullets
        content_box = slide.shapes.add_textbox(
            Inches(1), Inches(1.8),
            Inches(8), Inches(5)
        )
        text_frame = content_box.text_frame
        text_frame.word_wrap = True
        
        for i, item in enumerate(content_items):
            if i > 0:
                p = text_frame.add_paragraph()
            else:
                p = text_frame.paragraphs[0]
            
            p.text = item
            p.font.size = Pt(20)
            p.font.color.rgb = RGBColor(44, 62, 80)
            p.level = 0
            p.space_before = Pt(12)
    
    elif layout_type == "boxes":
        # Create visual boxes for content
        box_width = Inches(3.8)
        box_height = Inches(2.2)
        start_x = Inches(0.8)
        start_y = Inches(1.8)
        spacing = Inches(0.4)
        
        colors = [
            RGBColor(52, 152, 219),
            RGBColor(46, 204, 113),
            RGBColor(155, 89, 182),
            RGBColor(241, 196, 15)
        ]
        
        for i, item in enumerate(content_items[:4]):
            row = i // 2
            col = i % 2
            
            x = start_x + col * (box_width + spacing)
            y = start_y + row * (box_height + spacing)
            
            # Box shape
            box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                x, y,
                box_width, box_height
            )
            box.fill.solid()
            box.fill.fore_color.rgb = colors[i % len(colors)]
            box.line.color.rgb = RGBColor(255, 255, 255)
            box.line.width = Pt(3)
            box.shadow.inherit = False
            
            # Text in box
            text_box = slide.shapes.add_textbox(
                x + Inches(0.2), y + Inches(0.3),
                box_width - Inches(0.4), box_height - Inches(0.6)
            )
            text_frame = text_box.text_frame
            text_frame.text = item
            text_frame.word_wrap = True
            text_frame.paragraphs[0].font.size = Pt(16)
            text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            text_frame.paragraphs[0].font.bold = True
            text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    return slide

def add_timeline_slide(prs, title, timeline_items):
    """Add a timeline visualization slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0,
        prs.slide_width,
        prs.slide_height
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = RGBColor(236, 240, 241)
    bg_shape.line.fill.background()
    
    # Header
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0,
        prs.slide_width,
        Inches(1.2)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(41, 128, 185)
    header.line.fill.background()
    
    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(9), Inches(0.8)
    )
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].font.size = Pt(40)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    
    # Timeline line
    line = slide.shapes.add_connector(
        1,  # Straight connector
        Inches(1), Inches(4),
        Inches(9), Inches(4)
    )
    line.line.color.rgb = RGBColor(41, 128, 185)
    line.line.width = Pt(4)
    
    # Timeline points
    num_items = len(timeline_items)
    spacing = Inches(8) / (num_items - 1) if num_items > 1 else 0
    
    colors = [RGBColor(231, 76, 60), RGBColor(230, 126, 34), 
              RGBColor(241, 196, 15), RGBColor(46, 204, 113)]
    
    for i, (year, description) in enumerate(timeline_items):
        x = Inches(1) + (spacing * i)
        y = Inches(4)
        
        # Circle for timeline point
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            x - Inches(0.15), y - Inches(0.15),
            Inches(0.3), Inches(0.3)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = colors[i % len(colors)]
        circle.line.color.rgb = RGBColor(255, 255, 255)
        circle.line.width = Pt(3)
        
        # Year label
        year_box = slide.shapes.add_textbox(
            x - Inches(0.4), y - Inches(0.8),
            Inches(0.8), Inches(0.4)
        )
        year_frame = year_box.text_frame
        year_frame.text = year
        year_frame.paragraphs[0].font.size = Pt(18)
        year_frame.paragraphs[0].font.bold = True
        year_frame.paragraphs[0].font.color.rgb = colors[i % len(colors)]
        year_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Description
        desc_box = slide.shapes.add_textbox(
            x - Inches(0.6), y + Inches(0.3),
            Inches(1.2), Inches(1.5)
        )
        desc_frame = desc_box.text_frame
        desc_frame.text = description
        desc_frame.word_wrap = True
        desc_frame.paragraphs[0].font.size = Pt(12)
        desc_frame.paragraphs[0].font.color.rgb = RGBColor(44, 62, 80)
        desc_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    return slide

def add_diagram_slide(prs, title, center_text, surrounding_items):
    """Add a circular diagram slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0,
        prs.slide_width,
        prs.slide_height
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = RGBColor(236, 240, 241)
    bg_shape.line.fill.background()
    
    # Header
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0,
        prs.slide_width,
        Inches(1.2)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(41, 128, 185)
    header.line.fill.background()
    
    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(9), Inches(0.8)
    )
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].font.size = Pt(40)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    
    # Center circle
    center_x = Inches(5)
    center_y = Inches(4)
    center_radius = Inches(1)
    
    center_circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        center_x - center_radius, center_y - center_radius,
        center_radius * 2, center_radius * 2
    )
    center_circle.fill.solid()
    center_circle.fill.fore_color.rgb = RGBColor(41, 128, 185)
    center_circle.line.color.rgb = RGBColor(255, 255, 255)
    center_circle.line.width = Pt(4)
    
    # Center text
    center_text_box = slide.shapes.add_textbox(
        center_x - center_radius, center_y - Inches(0.3),
        center_radius * 2, Inches(0.6)
    )
    center_text_frame = center_text_box.text_frame
    center_text_frame.text = center_text
    center_text_frame.paragraphs[0].font.size = Pt(20)
    center_text_frame.paragraphs[0].font.bold = True
    center_text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    center_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    center_text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    # Surrounding items
    import math
    num_items = len(surrounding_items)
    angle_step = 2 * math.pi / num_items
    orbit_radius = Inches(2.5)
    
    colors = [
        RGBColor(52, 152, 219),
        RGBColor(46, 204, 113),
        RGBColor(155, 89, 182),
        RGBColor(241, 196, 15),
        RGBColor(231, 76, 60),
        RGBColor(26, 188, 156)
    ]
    
    for i, item in enumerate(surrounding_items):
        angle = i * angle_step - math.pi / 2
        x = center_x + orbit_radius * math.cos(angle)
        y = center_y + orbit_radius * math.sin(angle)
        
        item_radius = Inches(0.8)
        
        # Connecting line
        connector = slide.shapes.add_connector(
            1,
            int(center_x), int(center_y),
            int(x), int(y)
        )
        connector.line.color.rgb = colors[i % len(colors)]
        connector.line.width = Pt(2)
        
        # Item circle
        item_circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            x - item_radius, y - item_radius,
            item_radius * 2, item_radius * 2
        )
        item_circle.fill.solid()
        item_circle.fill.fore_color.rgb = colors[i % len(colors)]
        item_circle.line.color.rgb = RGBColor(255, 255, 255)
        item_circle.line.width = Pt(3)
        
        # Item text
        item_text_box = slide.shapes.add_textbox(
            x - item_radius, y - Inches(0.25),
            item_radius * 2, Inches(0.5)
        )
        item_text_frame = item_text_box.text_frame
        item_text_frame.text = item
        item_text_frame.word_wrap = True
        item_text_frame.paragraphs[0].font.size = Pt(14)
        item_text_frame.paragraphs[0].font.bold = True
        item_text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        item_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        item_text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    return slide

# Slide 1: Title Slide
add_title_slide(prs, "ARTIFICIAL INTELLIGENCE", "The Future of Technology")

# Slide 2: Introduction
slide2 = add_content_slide(prs, "Introduction", [
    "• AI: Simulation of human intelligence by machines",
    "• Systems that can learn, reason, and self-correct",
    "• Revolutionizing industries and daily life",
    "• From virtual assistants to autonomous vehicles",
    "• Transforming how we work, communicate, and solve problems"
], "bullet")

# Slide 3: History Timeline
add_timeline_slide(prs, "History of AI", [
    ("1950s", "Birth of AI\nTuring Test"),
    ("1960s-70s", "Early AI\nPrograms"),
    ("1980s-90s", "Expert\nSystems"),
    ("2000s-Now", "Machine\nLearning Era")
])

# Slide 4: Types of AI - Diagram
add_diagram_slide(prs, "Types of AI", "AI\nTypes", [
    "Narrow AI",
    "General AI",
    "Super\nIntelligence"
])

# Slide 5: Types Details
slide5 = add_content_slide(prs, "AI Classification", [
    "Narrow AI (ANI)\nSpecialized in specific tasks\nCurrent AI systems",
    "General AI (AGI)\nHuman-level intelligence\nTheoretical concept",
    "Superintelligence (ASI)\nSurpasses human intelligence\nFuture possibility",
    "Each level represents\nincreasing capability\nand complexity"
], "boxes")

# Slide 6: Applications - Diagram
add_diagram_slide(prs, "AI Applications", "AI in\nAction", [
    "Healthcare",
    "Finance",
    "Transport",
    "Entertainment",
    "Education",
    "Manufacturing"
])

# Slide 7: Healthcare Applications
slide7 = add_content_slide(prs, "AI in Healthcare", [
    "• Disease Diagnosis: AI analyzes medical images and patient data",
    "• Drug Discovery: Accelerates development of new medications",
    "• Personalized Treatment: Tailored healthcare plans",
    "• Predictive Analytics: Early disease detection",
    "• Robot-Assisted Surgery: Enhanced precision and outcomes"
], "bullet")

# Slide 8: Finance & Transportation
slide8 = add_content_slide(prs, "AI in Finance & Transportation", [
    "Finance:\n• Fraud detection\n• Algorithmic trading\n• Risk assessment",
    "Transportation:\n• Autonomous vehicles\n• Traffic optimization\n• Route planning",
    "Entertainment:\n• Content recommendations\n• Personalized experiences\n• Content creation",
    "Manufacturing:\n• Quality control\n• Predictive maintenance\n• Process optimization"
], "boxes")

# Slide 9: Challenges - Diagram
add_diagram_slide(prs, "Challenges & Ethics", "AI\nChallenges", [
    "Bias",
    "Privacy",
    "Job Loss",
    "Accountability",
    "Security"
])

# Slide 10: Ethical Considerations Details
slide10 = add_content_slide(prs, "Ethical Considerations", [
    "• Bias & Fairness: AI systems can perpetuate existing biases",
    "• Privacy Concerns: Data collection and surveillance issues",
    "• Job Displacement: Automation affecting employment",
    "• Accountability: Who is responsible for AI decisions?",
    "• Transparency: Understanding AI decision-making processes"
], "bullet")

# Slide 11: Future of AI
slide11 = add_content_slide(prs, "Future of AI", [
    "• Continued advancement in machine learning algorithms",
    "• Integration of AI in everyday devices and services",
    "• Development toward Artificial General Intelligence (AGI)",
    "• Enhanced human-AI collaboration and augmentation",
    "• Ethical frameworks and regulations for responsible AI"
], "bullet")

# Slide 12: Conclusion
slide12 = add_content_slide(prs, "Conclusion", [
    "AI is transforming\nour world",
    "Immense potential\nfor innovation",
    "Addressing challenges\nis crucial",
    "Responsible development\nand deployment"
], "boxes")

# Slide 13: Thank You
add_title_slide(prs, "THANK YOU", "Questions & Discussion")

# Save presentation
output_file = "/vercel/sandbox/AI_Enhanced_3D_Presentation.pptx"
prs.save(output_file)
print(f"✓ Enhanced 3D presentation created: {output_file}")
print(f"✓ Total slides: {len(prs.slides)}")
print("✓ Features: Timeline diagrams, circular diagrams, visual boxes, modern design")
