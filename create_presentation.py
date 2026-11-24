import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_deck(filename="ASD_EyeTracking_Presentation.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    DARK_BG = RGBColor(15, 23, 42)       # Slate 900
    NAVY = RGBColor(30, 27, 75)          # Indigo 950
    PRIMARY = RGBColor(99, 102, 241)     # Indigo 500
    ACCENT = RGBColor(139, 92, 246)      # Violet 500
    LIGHT_BG = RGBColor(248, 250, 252)   # Slate 50
    CARD_BG = RGBColor(255, 255, 255)
    CARD_BORDER = RGBColor(226, 232, 240)
    TEXT_DARK = RGBColor(30, 41, 59)     # Slate 800
    TEXT_MUTED = RGBColor(100, 116, 139) # Slate 500
    TEXT_LIGHT = RGBColor(255, 255, 255)
    SUCCESS_COLOR = RGBColor(16, 185, 129)

    def add_header(slide, title_text, category="EYE-TRACKING RESEARCH PORTAL"):
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11), Inches(0.35))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = PRIMARY
        p_cat.font.name = "Arial"

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.5), Inches(0.8))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(26)
        p_t.font.bold = True
        p_t.font.color.rgb = DARK_BG
        p_t.font.name = "Arial"

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        if border_color:
            shape.line.color.rgb = border_color
            shape.line.width = Pt(1.5)
        else:
            shape.line.fill.background()
        return shape

    # ==========================================
    # SLIDE 1: Title (Dark Theme)
    # ==========================================
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = DARK_BG
    bg1.line.fill.background()

    pill = add_card(slide1, Inches(1.0), Inches(1.3), Inches(3.8), Inches(0.45), bg_color=NAVY, border_color=PRIMARY)
    tb_pill = slide1.shapes.add_textbox(Inches(1.0), Inches(1.3), Inches(3.8), Inches(0.45))
    p_pill = tb_pill.text_frame.paragraphs[0]
    p_pill.text = "AI CLINICAL SCREENING PLATFORM"
    p_pill.font.size = Pt(11)
    p_pill.font.bold = True
    p_pill.font.color.rgb = PRIMARY
    p_pill.alignment = PP_ALIGN.CENTER

    tb_title = slide1.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(2.2))
    tf = tb_title.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Autism (ASD) Detection via Eye-Tracking"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT

    p2 = tf.add_paragraph()
    p2.text = "An Accurate, Explainable AI Screening Tool & Web Application"
    p2.font.size = Pt(22)
    p2.font.color.rgb = ACCENT

    tb_sub = slide1.shapes.add_textbox(Inches(1.0), Inches(4.3), Inches(11.3), Inches(1.2))
    tf_sub = tb_sub.text_frame
    tf_sub.word_wrap = True
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = "Using Machine Learning to analyze where children look when viewing images, achieving ~98% diagnostic accuracy with a simple, clinician-friendly web portal."
    p_sub.font.size = Pt(15)
    p_sub.font.color.rgb = RGBColor(203, 213, 225)

    badges = [
        ("Architecture", "MobileNet + PCA + Stacking"),
        ("Accuracy", "~98% on Benchmark Data"),
        ("Explainability", "Clear 'Why This Patient?' Cards"),
        ("Deployment", "Live Web Portal on Cloud")
    ]
    for i, (b_title, b_val) in enumerate(badges):
        b_x = Inches(1.0 + i * 2.85)
        add_card(slide1, b_x, Inches(5.7), Inches(2.7), Inches(1.1), bg_color=RGBColor(30, 41, 59), border_color=RGBColor(51, 65, 85))
        tb_b = slide1.shapes.add_textbox(b_x + Inches(0.15), Inches(5.8), Inches(2.4), Inches(0.9))
        tf_b = tb_b.text_frame
        p_bt = tf_b.paragraphs[0]
        p_bt.text = b_title.upper()
        p_bt.font.size = Pt(10)
        p_bt.font.bold = True
        p_bt.font.color.rgb = PRIMARY
        p_bv = tf_b.add_paragraph()
        p_bv.text = b_val
        p_bv.font.size = Pt(13)
        p_bv.font.bold = True
        p_bv.font.color.rgb = TEXT_LIGHT

    # ==========================================
    # SLIDE 2: Problem & Motivation (Simple English)
    # ==========================================
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "The Real-World Problem: Why We Built This")

    col_width = Inches(3.65)
    gap = Inches(0.35)
    cards_data = [
        ("The Current Problem", "Diagnosis Takes Too Long", 
         ["Autism affects 1 in 36 children worldwide.",
          "Traditional tests require specialized hospital visits and lengthy behavioral exams.",
          "Families often wait 2 to 3 years for an appointment, missing the most critical early development years."]),
        ("The Simple Idea", "Eyes Reveal Attention", 
         ["When looking at people, neurotypical children naturally look at eyes, faces, and smiles.",
          "Children with Autism often avoid direct eye contact and focus on background shapes or objects instead.",
          "Eye-tracking creates a 'map' of their gaze—giving us an objective, non-invasive digital clue."]),
        ("Our Solution", "An Instant AI Assistant", 
         ["We created an AI system that looks at that gaze map and makes an instant, 98% accurate prediction.",
          "Doctors do not need to understand code: they simply upload the image to our web interface.",
          "The tool explains exactly WHY it made the decision in simple, everyday language."])
    ]

    for i, (tag, heading, points) in enumerate(cards_data):
        cx = Inches(0.8 + i * (col_width + gap))
        add_card(slide2, cx, Inches(1.8), col_width, Inches(5.0))
        tb = slide2.shapes.add_textbox(cx + Inches(0.25), Inches(2.0), col_width - Inches(0.5), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p_tag = tf.paragraphs[0]
        p_tag.text = tag.upper()
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = PRIMARY
        
        p_head = tf.add_paragraph()
        p_head.text = heading
        p_head.font.size = Pt(18)
        p_head.font.bold = True
        p_head.font.color.rgb = DARK_BG
        p_head.space_after = Pt(14)

        for pt in points:
            p_pt = tf.add_paragraph()
            p_pt.text = "• " + pt
            p_pt.font.size = Pt(13)
            p_pt.font.color.rgb = TEXT_DARK
            p_pt.space_after = Pt(10)

    # ==========================================
    # SLIDE 3: How It Works (Simple 4 Steps)
    # ==========================================
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "How Our AI Works (4 Simple Steps)")

    stages = [
        ("Step 1", "Clean the Image", 
         ["Resize image to standard 224x224 pixels.", "Light blur to smooth jagged computer lines.", "Balance brightness so all images look uniform."]),
        ("Step 2", "Find Key Patterns", 
         ["We use MobileNet, a proven vision AI.", "It examines lines, curves, and gaze dots.", "Converts the image into 1,024 descriptive numbers."]),
        ("Step 3", "Keep What Matters", 
         ["1,024 numbers is too many and can cause confusion.", "PCA algorithm selects the 250 best signals.", "Retains over 95% of the information while making it 4x faster."]),
        ("Step 4", "Team of Classifiers", 
         ["Two smart models (SVM + KNN) cast their votes.", "A 'manager' model (Logistic Regression) makes the final call.", "Outputs the exact percentage chance of ASD."])
    ]

    s_width = Inches(2.7)
    s_gap = Inches(0.3)
    for i, (num, sub, points) in enumerate(stages):
        sx = Inches(0.8 + i * (s_width + s_gap))
        add_card(slide3, sx, Inches(1.8), s_width, Inches(5.0))
        tb = slide3.shapes.add_textbox(sx + Inches(0.2), Inches(2.0), s_width - Inches(0.4), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p_n = tf.paragraphs[0]
        p_n.text = num
        p_n.font.size = Pt(14)
        p_n.font.bold = True
        p_n.font.color.rgb = PRIMARY

        p_s = tf.add_paragraph()
        p_s.text = sub
        p_s.font.size = Pt(13)
        p_s.font.bold = True
        p_s.font.color.rgb = DARK_BG
        p_s.space_after = Pt(12)

        for pt in points:
            p_pt = tf.add_paragraph()
            p_pt.text = "• " + pt
            p_pt.font.size = Pt(12)
            p_pt.font.color.rgb = TEXT_DARK
            p_pt.space_after = Pt(8)

    # ==========================================
    # SLIDE 4: Results & Validation
    # ==========================================
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Results: Proving the System Works")

    add_card(slide4, Inches(0.8), Inches(1.8), Inches(6.5), Inches(5.0))
    tb_table = slide4.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.9), Inches(4.5))
    tf_t = tb_table.text_frame
    tf_t.word_wrap = True
    p = tf_t.paragraphs[0]
    p.text = "TEST RESULTS ON BENCHMARK DATASET"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(12)

    table_lines = [
        ("Metric", "Published Paper", "Our Pipeline", "Result"),
        ("Accuracy", "98.0%", "97.8%", "Matched!"),
        ("Precision", "98.4%", "98.2%", "Matched!"),
        ("Sensitivity", "98.2%", "97.9%", "Matched!"),
        ("Specificity", "97.8%", "97.7%", "Matched!")
    ]
    for row in table_lines:
        p_r = tf_t.add_paragraph()
        p_r.text = f"{row[0]:<18} | {row[1]:<16} | {row[2]:<16} | {row[3]}"
        p_r.font.size = Pt(12)
        p_r.font.name = "Courier New"
        if row[0] == "Metric":
            p_r.font.bold = True
            p_r.font.color.rgb = DARK_BG
        else:
            p_r.font.color.rgb = TEXT_DARK
        p_r.space_after = Pt(8)

    add_card(slide4, Inches(7.6), Inches(1.8), Inches(4.9), Inches(5.0))
    tb_take = slide4.shapes.add_textbox(Inches(7.9), Inches(2.0), Inches(4.3), Inches(4.5))
    tf_k = tb_take.text_frame
    tf_k.word_wrap = True
    p_k = tf_k.paragraphs[0]
    p_k.text = "WHY OUR RESULTS ARE STRONG"
    p_k.font.size = Pt(12)
    p_k.font.bold = True
    p_k.font.color.rgb = PRIMARY
    p_k.space_after = Pt(10)

    takeaways = [
        ("Two Models are Better Than One", "By combining SVM and KNN, the two algorithms check each other's work. If one model is unsure, the other provides balance."),
        ("No Overfitting", "Using PCA to reduce features prevented the model from memorizing the test images, meaning it performs reliably on completely new patients."),
        ("Fast Response", "Predictions take less than half a second (400 milliseconds), making it practical for real clinics.")
    ]
    for h_txt, b_txt in takeaways:
        p_th = tf_k.add_paragraph()
        p_th.text = h_txt
        p_th.font.size = Pt(14)
        p_th.font.bold = True
        p_th.font.color.rgb = DARK_BG
        p_tb = tf_k.add_paragraph()
        p_tb.text = b_txt
        p_tb.font.size = Pt(12)
        p_tb.font.color.rgb = TEXT_DARK
        p_tb.space_after = Pt(10)

    # ==========================================
    # SLIDE 5: The Web Application
    # ==========================================
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "The Frontend: Designed for Real Doctors & Clinics")

    add_card(slide5, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.0))
    tb_app = slide5.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.0), Inches(4.5))
    tf_app = tb_app.text_frame
    tf_app.word_wrap = True
    p_a = tf_app.paragraphs[0]
    p_a.text = "SIMPLE & USER-FRIENDLY"
    p_a.font.size = Pt(12)
    p_a.font.bold = True
    p_a.font.color.rgb = PRIMARY
    p_a.space_after = Pt(10)

    app_features = [
        ("No Coding Needed", "Clinicians do not need to touch Python or Jupyter notebooks. Everything runs through a clean website."),
        ("Drag-and-Drop Upload", "Drag any eye-tracking scanpath image or click 'Browse files' to upload."),
        ("Instant Image Preview", "Displays the uploaded file instantly so the user can verify they chose the right image."),
        ("Instant Results & Progress Bar", "Clicking 'Run Analysis' animates a clear percentage score with color badges (Green for Neurotypical, Purple for ASD).")
    ]
    for af_h, af_b in app_features:
        p_h = tf_app.add_paragraph()
        p_h.text = "• " + af_h
        p_h.font.size = Pt(13)
        p_h.font.bold = True
        p_h.font.color.rgb = DARK_BG
        p_b = tf_app.add_paragraph()
        p_b.text = "   " + af_b
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = TEXT_MUTED
        p_b.space_after = Pt(8)

    add_card(slide5, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.0), bg_color=LIGHT_BG)
    tb_flow = slide5.shapes.add_textbox(Inches(7.1), Inches(2.0), Inches(5.1), Inches(4.5))
    tf_fl = tb_flow.text_frame
    tf_fl.word_wrap = True
    p_fl = tf_fl.paragraphs[0]
    p_fl.text = "STEP-BY-STEP USER WORKFLOW"
    p_fl.font.size = Pt(12)
    p_fl.font.bold = True
    p_fl.font.color.rgb = PRIMARY
    p_fl.space_after = Pt(12)

    flow_steps = [
        ("1. Drop Image", "Upload the patient's eye-tracking gaze plot."),
        ("2. Backend Processes", "The Flask server applies the 4 notebook steps behind the scenes."),
        ("3. Clear Status Badge", "See 'ASD Indicated' or 'Neurotypical Pattern' immediately."),
        ("4. Understand Why", "Read the 3 visual explanation cards."),
        ("5. Next Step Guidance", "Provides recommended standard clinical evaluations (e.g. ADOS-2).")
    ]
    for st_h, st_b in flow_steps:
        p_st = tf_fl.add_paragraph()
        p_st.text = st_h
        p_st.font.size = Pt(13)
        p_st.font.bold = True
        p_st.font.color.rgb = DARK_BG
        p_sb = tf_fl.add_paragraph()
        p_sb.text = st_b
        p_sb.font.size = Pt(11)
        p_sb.font.color.rgb = TEXT_DARK
        p_sb.space_after = Pt(6)

    # ==========================================
    # SLIDE 6: Explainability in Simple English (Replaced Jargon)
    # ==========================================
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Explainable AI: Why Did the System Make This Decision?")

    bio_cards = [
        ("CHECK 1", "Where Did They Look?", "Peripheral vs Center",
         "Children with Autism often fixate on the outer borders or empty background of an image rather than looking across the whole scene. Our system measures if their gaze is clustered away from central objects."),
        ("CHECK 2", "Eye Contact & Social Faces", "Avoiding Faces",
         "When shown images of people, neurotypical children spend ~68% of their time looking at eyes and mouths. In ASD scanpaths, this drops significantly (often below 30%), showing clear face avoidance."),
        ("CHECK 3", "Smooth vs. Jerky Eye Jumps", "Restless Movements",
         "Instead of following calm, smooth lines between objects, children with Autism often make short, scattered, zigzag eye movements. The system flags this high path jitter as a key indicator.")
    ]
    for i, (b_tag, b_name, b_stat, b_desc) in enumerate(bio_cards):
        bx = Inches(0.8 + i * 3.9)
        add_card(slide6, bx, Inches(1.8), Inches(3.6), Inches(5.0))
        tb_b = slide6.shapes.add_textbox(bx + Inches(0.2), Inches(2.0), Inches(3.2), Inches(4.5))
        tf_bio = tb_b.text_frame
        tf_bio.word_wrap = True
        p_btag = tf_bio.paragraphs[0]
        p_btag.text = b_tag
        p_btag.font.size = Pt(11)
        p_btag.font.bold = True
        p_btag.font.color.rgb = PRIMARY

        p_bn = tf_bio.add_paragraph()
        p_bn.text = b_name
        p_bn.font.size = Pt(16)
        p_bn.font.bold = True
        p_bn.font.color.rgb = DARK_BG

        p_bs = tf_bio.add_paragraph()
        p_bs.text = b_stat
        p_bs.font.size = Pt(12)
        p_bs.font.bold = True
        p_bs.font.color.rgb = ACCENT
        p_bs.space_after = Pt(10)

        p_bd = tf_bio.add_paragraph()
        p_bd.text = b_desc
        p_bd.font.size = Pt(12)
        p_bd.font.color.rgb = TEXT_DARK
        p_bd.space_after = Pt(12)

        p_cl = tf_bio.add_paragraph()
        p_cl.text = "Why this matters: Doctors trust an AI that explains its thinking, rather than a mystery score."
        p_cl.font.size = Pt(11)
        p_cl.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 7: Handbook & User Guide (New Slide)
    # ==========================================
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "System Handbook: Instructions for Users & Clinic Staff")

    hb_cards = [
        ("IMAGE REQUIREMENTS", "What Images Can Be Uploaded?", 
         ["Supports standard PNG, JPG, and JPEG formats.",
          "Compatible with raw gaze plots, line scanpaths, and colored heatmaps.",
          "Any screen resolution works—the server automatically standardizes sizing to 224x224."]),
        ("BUILT-IN SAFETY", "Simulation & Fallback Mode", 
         ["If trained model files are not installed, the app NEVER crashes.",
          "It automatically switches to a 'contour-simulation mode' to allow offline UI demonstrations.",
          "A clear banner informs the doctor whether live weights or simulation mode is active."]),
        ("CLINICAL RESPONSIBILITY", "Screening Aid vs Final Diagnosis", 
         ["The portal explicitly states it is a preliminary screening assistant, not a replacement for a doctor.",
          "Provides a quantitative confidence score so clinicians can see if a result is borderline or strong.",
          "Recommends formal ADOS-2 testing when a high-risk pattern is detected."])
    ]
    for i, (h_tag, h_title, h_points) in enumerate(hb_cards):
        hx = Inches(0.8 + i * 3.9)
        add_card(slide7, hx, Inches(1.8), Inches(3.6), Inches(5.0))
        tb_h = slide7.shapes.add_textbox(hx + Inches(0.2), Inches(2.0), Inches(3.2), Inches(4.5))
        tf_h = tb_h.text_frame
        tf_h.word_wrap = True
        p_ht = tf_h.paragraphs[0]
        p_ht.text = h_tag
        p_ht.font.size = Pt(11)
        p_ht.font.bold = True
        p_ht.font.color.rgb = PRIMARY

        p_hh = tf_h.add_paragraph()
        p_hh.text = h_title
        p_hh.font.size = Pt(16)
        p_hh.font.bold = True
        p_hh.font.color.rgb = DARK_BG
        p_hh.space_after = Pt(10)

        for pt in h_points:
            p_p = tf_h.add_paragraph()
            p_p.text = "• " + pt
            p_p.font.size = Pt(12)
            p_p.font.color.rgb = TEXT_DARK
            p_p.space_after = Pt(10)

    # ==========================================
    # SLIDE 8: Summary (Dark Theme)
    # ==========================================
    slide8 = prs.slides.add_slide(blank_layout)
    bg8 = slide8.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg8.fill.solid()
    bg8.fill.fore_color.rgb = DARK_BG
    bg8.line.fill.background()

    tb_end = slide8.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.3), Inches(4.5))
    tf_e = tb_end.text_frame
    tf_e.word_wrap = True
    pe = tf_e.paragraphs[0]
    pe.text = "SUMMARY & KEY TAKEAWAYS"
    pe.font.size = Pt(14)
    pe.font.bold = True
    pe.font.color.rgb = PRIMARY
    pe.space_after = Pt(10)

    pe2 = tf_e.add_paragraph()
    pe2.text = "What We Have Delivered:"
    pe2.font.size = Pt(28)
    pe2.font.bold = True
    pe2.font.color.rgb = TEXT_LIGHT
    pe2.space_after = Pt(14)

    recap_points = [
        "1. High Diagnostic Accuracy: Achieved ~98% accuracy on benchmark eye-tracking data using our MobileNet + PCA + Stacking pipeline.",
        "2. Understandable Explanations: Replaced complex math terms with 3 simple visual checks (where they look, face avoidance, and eye jump jitter).",
        "3. Real Working Website: Built a clean, modern web application where any user can drag, drop, and test scanpaths.",
        "4. Ready for Cloud & Open Source: Code is organized on GitHub with production Gunicorn servers and pre-configured for free Render cloud deployment."
    ]
    for rpt in recap_points:
        pr = tf_e.add_paragraph()
        pr.text = rpt
        pr.font.size = Pt(15)
        pr.font.color.rgb = RGBColor(226, 232, 240)
        pr.space_after = Pt(12)

    prs.save(filename)
    print(f"Clean presentation created at {filename}")

if __name__ == "__main__":
    create_deck(sys.argv[1] if len(sys.argv) > 1 else "ASD_EyeTracking_Presentation.pptx")
