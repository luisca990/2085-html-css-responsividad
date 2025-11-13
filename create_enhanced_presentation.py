from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image, ImageDraw, ImageFont
import io

# Create presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

def add_gradient_background(slide, color1, color2):
    """Add a gradient background to slide"""
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 90.0
    fill.gradient_stops[0].color.rgb = RGBColor(*color1)
    fill.gradient_stops[1].color.rgb = RGBColor(*color2)

def add_3d_shape(slide, shape_type, left, top, width, height, color, text="", depth=True):
    """Add a 3D shape with shadow and depth effects"""
    shape = slide.shapes.add_shape(
        shape_type,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height)
    )
    
    # Fill color
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*color)
    
    # Add 3D effects
    shape.shadow.inherit = False
    shape.shadow.shadow_type = 'outer'
    shape.shadow.distance = Pt(5)
    shape.shadow.angle = 45
    shape.shadow.blur_radius = Pt(8)
    shape.shadow.transparency = 0.5
    
    # Line style
    shape.line.color.rgb = RGBColor(255, 255, 255)
    shape.line.width = Pt(2)
    
    if text:
        text_frame = shape.text_frame
        text_frame.text = text
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        for paragraph in text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
    
    return shape

def add_title_slide():
    """Slide 1: Title Slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    add_gradient_background(slide, (15, 32, 64), (44, 62, 80))
    
    # Main title with 3D effect
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = "ARTIFICIAL INTELLIGENCE\nIN CYBERSECURITY"
    title_frame.word_wrap = True
    
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(54)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Add shadow to title
    title_box.shadow.inherit = False
    title_box.shadow.shadow_type = 'outer'
    title_box.shadow.distance = Pt(8)
    title_box.shadow.blur_radius = Pt(12)
    
    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(2), Inches(4.5), Inches(6), Inches(0.8))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Protecting Digital Assets with Intelligent Systems"
    
    for paragraph in subtitle_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(24)
            run.font.color.rgb = RGBColor(100, 200, 255)
            run.font.italic = True
    
    # Decorative 3D shapes
    add_3d_shape(slide, MSO_SHAPE.HEXAGON, 0.5, 0.5, 1, 1, (52, 152, 219))
    add_3d_shape(slide, MSO_SHAPE.HEXAGON, 8.5, 6, 1, 1, (46, 204, 113))
    add_3d_shape(slide, MSO_SHAPE.HEXAGON, 8.5, 0.5, 1, 1, (155, 89, 182))

def add_agenda_slide():
    """Slide 2: Agenda"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, (26, 42, 74), (52, 73, 94))
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "AGENDA"
    
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(44)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Agenda items with 3D boxes
    agenda_items = [
        ("1", "AI Applications in Cybersecurity", (52, 152, 219)),
        ("2", "Machine Learning in Threat Detection", (46, 204, 113)),
        ("3", "Challenges & Limitations", (231, 76, 60)),
        ("4", "Future Trends", (155, 89, 182))
    ]
    
    start_top = 2
    for i, (num, text, color) in enumerate(agenda_items):
        # Number circle
        circle = add_3d_shape(slide, MSO_SHAPE.OVAL, 1.5, start_top + i * 1.1, 0.6, 0.6, color, num)
        
        # Text box
        text_box = slide.shapes.add_textbox(Inches(2.5), Inches(start_top + i * 1.1), Inches(6), Inches(0.6))
        text_frame = text_box.text_frame
        text_frame.text = text
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(22)
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.bold = True

def add_ai_applications_slide():
    """Slide 3: AI Applications in Cybersecurity"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, (26, 42, 74), (52, 73, 94))
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "AI APPLICATIONS IN CYBERSECURITY"
    
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(36)
            run.font.bold = True
            run.font.color.rgb = RGBColor(52, 152, 219)
    
    # Central AI hub
    add_3d_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 4, 2.5, 2, 1.2, (52, 152, 219), "AI\nCORE")
    
    # Applications around the center
    applications = [
        ("Threat\nDetection", 1.5, 1.5, (46, 204, 113)),
        ("Anomaly\nDetection", 7.5, 1.5, (155, 89, 182)),
        ("Automated\nResponse", 1.5, 4.5, (230, 126, 34)),
        ("Predictive\nAnalysis", 7.5, 4.5, (231, 76, 60))
    ]
    
    for text, left, top, color in applications:
        add_3d_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, 1.8, 1, color, text)
        
        # Add connecting lines (arrows)
        connector = slide.shapes.add_connector(
            1,  # Straight connector
            Inches(left + 0.9 if left < 5 else left),
            Inches(top + 0.5),
            Inches(5 if left < 5 else 5),
            Inches(3.1)
        )
        connector.line.color.rgb = RGBColor(255, 255, 255)
        connector.line.width = Pt(2)

def add_ml_threat_detection_slide():
    """Slide 4: Machine Learning in Threat Detection"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, (26, 42, 74), (52, 73, 94))
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "MACHINE LEARNING IN THREAT DETECTION"
    
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(32)
            run.font.bold = True
            run.font.color.rgb = RGBColor(46, 204, 113)
    
    # Process flow diagram
    steps = [
        ("Data\nCollection", 1, 2, (52, 152, 219)),
        ("Feature\nExtraction", 3.5, 2, (46, 204, 113)),
        ("Model\nTraining", 6, 2, (155, 89, 182)),
        ("Threat\nPrediction", 8.5, 2, (231, 76, 60))
    ]
    
    for i, (text, left, top, color) in enumerate(steps):
        add_3d_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, 1.8, 1.2, color, text)
        
        # Add arrow to next step
        if i < len(steps) - 1:
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                Inches(left + 1.9),
                Inches(top + 0.4),
                Inches(0.5),
                Inches(0.4)
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = RGBColor(255, 255, 255)
            arrow.line.color.rgb = RGBColor(255, 255, 255)
    
    # Key benefits
    benefits_box = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(8), Inches(2.5))
    benefits_frame = benefits_box.text_frame
    benefits_frame.text = "KEY BENEFITS:\n\n• Real-time pattern recognition\n• Adaptive learning from new threats\n• Reduced false positives\n• Automated threat classification"
    
    for paragraph in benefits_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(18)
            run.font.color.rgb = RGBColor(255, 255, 255)
            if "KEY BENEFITS:" in run.text:
                run.font.bold = True
                run.font.size = Pt(22)

def add_ml_techniques_slide():
    """Slide 5: ML Techniques"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, (26, 42, 74), (52, 73, 94))
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "MACHINE LEARNING TECHNIQUES"
    
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(36)
            run.font.bold = True
            run.font.color.rgb = RGBColor(46, 204, 113)
    
    # Three columns for techniques
    techniques = [
        ("Supervised\nLearning", "Classification\nof known threats", (52, 152, 219)),
        ("Unsupervised\nLearning", "Anomaly\ndetection", (155, 89, 182)),
        ("Deep\nLearning", "Complex pattern\nrecognition", (46, 204, 113))
    ]
    
    start_left = 1.5
    for i, (title, desc, color) in enumerate(techniques):
        # Main box
        add_3d_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, start_left + i * 2.8, 2, 2.3, 1.5, color, title)
        
        # Description box
        desc_box = slide.shapes.add_textbox(
            Inches(start_left + i * 2.8),
            Inches(3.8),
            Inches(2.3),
            Inches(1)
        )
        desc_frame = desc_box.text_frame
        desc_frame.text = desc
        desc_frame.word_wrap = True
        
        for paragraph in desc_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Bottom note
    note_box = slide.shapes.add_textbox(Inches(1), Inches(5.5), Inches(8), Inches(1.5))
    note_frame = note_box.text_frame
    note_frame.text = "These techniques work together to create robust, adaptive security systems capable of identifying both known and emerging threats."
    
    for paragraph in note_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(200, 200, 200)
            run.font.italic = True

def add_challenges_slide():
    """Slide 6: Challenges & Limitations"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, (26, 42, 74), (52, 73, 94))
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "CHALLENGES & LIMITATIONS"
    
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(36)
            run.font.bold = True
            run.font.color.rgb = RGBColor(231, 76, 60)
    
    # Challenge boxes in grid
    challenges = [
        ("Data Quality", "Requires large,\nlabeled datasets", 1.5, 1.8, (231, 76, 60)),
        ("Adversarial\nAttacks", "AI systems can\nbe manipulated", 5.5, 1.8, (230, 126, 34)),
        ("False\nPositives", "Balance between\nsecurity & usability", 1.5, 4, (192, 57, 43)),
        ("Explainability", "Black box\ndecision making", 5.5, 4, (211, 84, 0))
    ]
    
    for title, desc, left, top, color in challenges:
        # Challenge title
        add_3d_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, 2.5, 0.8, color, title)
        
        # Description
        desc_box = slide.shapes.add_textbox(Inches(left), Inches(top + 0.9), Inches(2.5), Inches(0.8))
        desc_frame = desc_box.text_frame
        desc_frame.text = desc
        desc_frame.word_wrap = True
        
        for paragraph in desc_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(255, 255, 255)

def add_future_trends_slide():
    """Slide 7: Future Trends"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, (26, 42, 74), (52, 73, 94))
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "FUTURE TRENDS IN AI CYBERSECURITY"
    
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(34)
            run.font.bold = True
            run.font.color.rgb = RGBColor(155, 89, 182)
    
    # Timeline/roadmap style
    trends = [
        ("Quantum-Safe\nAI", "2025-2027", (52, 152, 219)),
        ("Autonomous\nSecurity", "2026-2028", (46, 204, 113)),
        ("Federated\nLearning", "2027-2029", (155, 89, 182)),
        ("AI-Powered\nZero Trust", "2028-2030", (230, 126, 34))
    ]
    
    # Draw timeline
    timeline = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1),
        Inches(3.5),
        Inches(8),
        Inches(0.1)
    )
    timeline.fill.solid()
    timeline.fill.fore_color.rgb = RGBColor(255, 255, 255)
    timeline.line.color.rgb = RGBColor(255, 255, 255)
    
    for i, (title, year, color) in enumerate(trends):
        left = 1.5 + i * 2
        
        # Trend box
        add_3d_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, 1.8, 1.8, 1, color, title)
        
        # Year label
        year_box = slide.shapes.add_textbox(Inches(left), Inches(4.2), Inches(1.8), Inches(0.5))
        year_frame = year_box.text_frame
        year_frame.text = year
        
        for paragraph in year_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.bold = True
        
        # Connection dot
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(left + 0.8),
            Inches(3.4),
            Inches(0.2),
            Inches(0.2)
        )
        dot.fill.solid()
        dot.fill.fore_color.rgb = RGBColor(*color)
        dot.line.color.rgb = RGBColor(255, 255, 255)

def add_conclusion_slide():
    """Slide 8: Conclusion"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, (26, 42, 74), (52, 73, 94))
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "CONCLUSION"
    
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(44)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Key takeaways
    takeaways = [
        ("AI transforms cybersecurity through intelligent automation", (52, 152, 219)),
        ("Machine learning enables proactive threat detection", (46, 204, 113)),
        ("Challenges remain but solutions are evolving", (230, 126, 34)),
        ("The future is autonomous, adaptive security", (155, 89, 182))
    ]
    
    start_top = 2.2
    for i, (text, color) in enumerate(takeaways):
        # Icon/number
        add_3d_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.5, start_top + i * 1.1, 0.5, 0.5, color, str(i+1))
        
        # Text
        text_box = slide.shapes.add_textbox(Inches(2.3), Inches(start_top + i * 1.1), Inches(6.5), Inches(0.7))
        text_frame = text_box.text_frame
        text_frame.text = text
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(18)
                run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Final message
    final_box = slide.shapes.add_textbox(Inches(2), Inches(6.3), Inches(6), Inches(0.8))
    final_frame = final_box.text_frame
    final_frame.text = "AI is not just the future of cybersecurity—\nit's the present necessity"
    
    for paragraph in final_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(20)
            run.font.color.rgb = RGBColor(100, 200, 255)
            run.font.italic = True
            run.font.bold = True

def add_thank_you_slide():
    """Slide 9: Thank You"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_background(slide, (15, 32, 64), (44, 62, 80))
    
    # Thank you text
    title_box = slide.shapes.add_textbox(Inches(2), Inches(2.5), Inches(6), Inches(2))
    title_frame = title_box.text_frame
    title_frame.text = "THANK YOU"
    
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(64)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Add shadow
    title_box.shadow.inherit = False
    title_box.shadow.shadow_type = 'outer'
    title_box.shadow.distance = Pt(10)
    title_box.shadow.blur_radius = Pt(15)
    
    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(2), Inches(4.5), Inches(6), Inches(0.8))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Questions & Discussion"
    
    for paragraph in subtitle_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(28)
            run.font.color.rgb = RGBColor(100, 200, 255)
    
    # Decorative elements
    add_3d_shape(slide, MSO_SHAPE.HEXAGON, 1, 1, 0.8, 0.8, (52, 152, 219))
    add_3d_shape(slide, MSO_SHAPE.HEXAGON, 8.2, 1, 0.8, 0.8, (46, 204, 113))
    add_3d_shape(slide, MSO_SHAPE.HEXAGON, 1, 5.7, 0.8, 0.8, (155, 89, 182))
    add_3d_shape(slide, MSO_SHAPE.HEXAGON, 8.2, 5.7, 0.8, 0.8, (230, 126, 34))

# Create all slides
print("Creating enhanced 3D presentation...")
add_title_slide()
print("✓ Title slide created")
add_agenda_slide()
print("✓ Agenda slide created")
add_ai_applications_slide()
print("✓ AI Applications slide created")
add_ml_threat_detection_slide()
print("✓ ML Threat Detection slide created")
add_ml_techniques_slide()
print("✓ ML Techniques slide created")
add_challenges_slide()
print("✓ Challenges slide created")
add_future_trends_slide()
print("✓ Future Trends slide created")
add_conclusion_slide()
print("✓ Conclusion slide created")
add_thank_you_slide()
print("✓ Thank You slide created")

# Save presentation
output_file = '/vercel/sandbox/AI_Cybersecurity_Enhanced_3D.pptx'
prs.save(output_file)
print(f"\n✅ Presentation saved successfully: {output_file}")
print(f"Total slides: {len(prs.slides)}")
