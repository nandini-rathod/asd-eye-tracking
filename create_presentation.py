import os
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

    app_dir = os.path.dirname(os.path.abspath(__file__))
    img_upload_path = os.path.join(app_dir, "screenshot_upload.png")
    img_output_path = os.path.join(app_dir, "screenshot_output.jpg")

    def add_header(slide, title_text, category="EYE-TRACKING RESEARCH PORTAL"):
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11), Inches(0.35))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = PRIMARY
        p_cat.font.name = "Arial"

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.5), Inches(0.8))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(25)
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
    # SLIDE 1: EXACT FIRST PAGE (Title + Team + Supervisor)
    # ==========================================
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = DARK_BG
    bg1.line.fill.background()

    pill = add_card(slide1, Inches(1.0), Inches(1.1), Inches(3.8), Inches(0.45), bg_color=NAVY, border_color=PRIMARY)
    tb_pill = slide1.shapes.add_textbox(Inches(1.0), Inches(1.1), Inches(3.8), Inches(0.45))
    p_pill = tb_pill.text_frame.paragraphs[0]
    p_pill.text = "FINAL YEAR RESEARCH PROJECT"
    p_pill.font.size = Pt(11)
    p_pill.font.bold = True
    p_pill.font.color.rgb = PRIMARY
    p_pill.alignment = PP_ALIGN.CENTER

    tb_title = slide1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(2.0))
    tf = tb_title.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Autism Spectrum Disorder (ASD) Prediction"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = TEXT_LIGHT

    p2 = tf.add_paragraph()
    p2.text = "using Deep Learning & Feature Reduction"
    p2.font.size = Pt(26)
    p2.font.bold = True
    p2.font.color.rgb = ACCENT

    add_card(slide1, Inches(1.0), Inches(4.2), Inches(6.2), Inches(2.6), bg_color=RGBColor(30, 41, 59), border_color=RGBColor(51, 65, 85))
    tb_team = slide1.shapes.add_textbox(Inches(1.3), Inches(4.4), Inches(5.6), Inches(2.2))
    tf_team = tb_team.text_frame
    tf_team.word_wrap = True
    p_th = tf_team.paragraphs[0]
    p_th.text = "TEAM MEMBERS"
    p_th.font.size = Pt(12)
    p_th.font.bold = True
    p_th.font.color.rgb = PRIMARY
    p_th.space_after = Pt(8)

    members = [
        ("Pranav Priy", "2023UGCS046"),
        ("Tanisha Jain", "2023UGCS122"),
        ("Rathod Nandini", "2023UGCS079")
    ]
    for m_name, m_roll in members:
        pm = tf_team.add_paragraph()
        pm.text = f"{m_name:<20} •  {m_roll}"
        pm.font.size = Pt(14)
        pm.font.bold = True
        pm.font.color.rgb = TEXT_LIGHT
        pm.space_after = Pt(4)

    add_card(slide1, Inches(7.5), Inches(4.2), Inches(4.8), Inches(2.6), bg_color=RGBColor(30, 41, 59), border_color=RGBColor(51, 65, 85))
    tb_sup = slide1.shapes.add_textbox(Inches(7.8), Inches(4.4), Inches(4.2), Inches(2.2))
    tf_sup = tb_sup.text_frame
    tf_sup.word_wrap = True
    p_sh = tf_sup.paragraphs[0]
    p_sh.text = "SUPERVISOR"
    p_sh.font.size = Pt(12)
    p_sh.font.bold = True
    p_sh.font.color.rgb = PRIMARY
    p_sh.space_after = Pt(12)

    ps1 = tf_sup.add_paragraph()
    ps1.text = "Dr. Nidhi Kushwaha"
    ps1.font.size = Pt(18)
    ps1.font.bold = True
    ps1.font.color.rgb = TEXT_LIGHT
    ps1.space_after = Pt(6)

    ps2 = tf_sup.add_paragraph()
    ps2.text = "Department of Computer Science & Engineering"
    ps2.font.size = Pt(13)
    ps2.font.color.rgb = RGBColor(203, 213, 225)

    # ==========================================
    # SLIDE 2: Problem Statement & Objectives
    # ==========================================
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Problem Statement & Project Objectives")

    col_width = Inches(5.6)
    add_card(slide2, Inches(0.8), Inches(1.6), col_width, Inches(5.3))
    tb_prob = slide2.shapes.add_textbox(Inches(1.1), Inches(1.8), col_width - Inches(0.6), Inches(4.8))
    tf_p = tb_prob.text_frame
    tf_p.word_wrap = True
    p_ph = tf_p.paragraphs[0]
    p_ph.text = "TRADITIONAL DIAGNOSIS BOTTLENECKS"
    p_ph.font.size = Pt(12)
    p_ph.font.bold = True
    p_ph.font.color.rgb = PRIMARY
    p_ph.space_after = Pt(10)

    prob_items = [
        ("Delayed Diagnosis (4–5 Years)", "Diagnosis typically occurs around age 4–5, meaning children miss the critical early developmental intervention window."),
        ("High Cost & Specialist Scarcity", "Requires multiple lengthy behavioral evaluations with specialized clinicians, leading to long waiting lists."),
        ("Subjectivity in Observation", "Behavioral observations vary from clinician to clinician with no objective digital biomarkers.")
    ]
    for p_head, p_desc in prob_items:
        ph = tf_p.add_paragraph()
        ph.text = "• " + p_head
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = DARK_BG
        pd = tf_p.add_paragraph()
        pd.text = "   " + p_desc
        pd.font.size = Pt(11)
        pd.font.color.rgb = TEXT_MUTED
        pd.space_after = Pt(8)

    add_card(slide2, Inches(6.8), Inches(1.6), col_width, Inches(5.3), bg_color=LIGHT_BG)
    tb_obj = slide2.shapes.add_textbox(Inches(7.1), Inches(1.8), col_width - Inches(0.6), Inches(4.8))
    tf_o = tb_obj.text_frame
    tf_o.word_wrap = True
    p_oh = tf_o.paragraphs[0]
    p_oh.text = "PROJECT OBJECTIVES"
    p_oh.font.size = Pt(12)
    p_oh.font.bold = True
    p_oh.font.color.rgb = PRIMARY
    p_oh.space_after = Pt(10)

    obj_items = [
        ("01. ASD Classification", "Accurately classify subjects as ASD or Non-ASD using eye-tracking scanpaths."),
        ("02. Deep Feature Extraction", "Extract meaningful visual gaze representations using MobileNet."),
        ("03. Dimensionality Reduction", "Apply PCA to reduce features to 250 components, removing noise and redundancy."),
        ("04. Stacking Ensemble Learning", "Maximize accuracy via an SVM + KNN -> Logistic Regression meta-learner."),
        ("05. Clinical Web Tool & Explainability", "Deliver an accessible web portal with clear 'Why This Patient?' visual explanations.")
    ]
    for o_head, o_desc in obj_items:
        oh = tf_o.add_paragraph()
        oh.text = o_head
        oh.font.size = Pt(13)
        oh.font.bold = True
        oh.font.color.rgb = DARK_BG
        od = tf_o.add_paragraph()
        od.text = o_desc
        od.font.size = Pt(11)
        od.font.color.rgb = TEXT_DARK
        od.space_after = Pt(6)

    # ==========================================
    # SLIDE 3: Dataset & Preprocessing Pipeline
    # ==========================================
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Dataset Description & Preprocessing Pipeline")

    add_card(slide3, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.3))
    tb_data = slide3.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.0), Inches(4.8))
    tf_d = tb_data.text_frame
    tf_d.word_wrap = True
    p_dh = tf_d.paragraphs[0]
    p_dh.text = "ETSDS EYE TRACKING DATASET"
    p_dh.font.size = Pt(12)
    p_dh.font.bold = True
    p_dh.font.color.rgb = PRIMARY
    p_dh.space_after = Pt(10)

    data_points = [
        ("Modality", "High-speed eye-tracking devices recording how individuals look at social visual images."),
        ("Classes", "ASD (219 images) vs. Non-ASD / Neurotypical (328 images). Total ~547 raw samples."),
        ("Data Split", "Stratified 80% Training / 20% Testing split to preserve class ratios."),
        ("Data Augmentation", "Rotation (+/-15°), horizontal and vertical flips to expand training set to ~1750 images, preventing overfitting.")
    ]
    for d_h, d_b in data_points:
        pdh = tf_d.add_paragraph()
        pdh.text = "• " + d_h + ":"
        pdh.font.size = Pt(13)
        pdh.font.bold = True
        pdh.font.color.rgb = DARK_BG
        pdb = tf_d.add_paragraph()
        pdb.text = "   " + d_b
        pdb.font.size = Pt(11)
        pdb.font.color.rgb = TEXT_MUTED
        pdb.space_after = Pt(8)

    add_card(slide3, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.3), bg_color=LIGHT_BG)
    tb_prep = slide3.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(4.8))
    tf_pr = tb_prep.text_frame
    tf_pr.word_wrap = True
    p_prh = tf_pr.paragraphs[0]
    p_prh.text = "PREPROCESSING PIPELINE"
    p_prh.font.size = Pt(12)
    p_prh.font.bold = True
    p_prh.font.color.rgb = PRIMARY
    p_prh.space_after = Pt(10)

    prep_steps = [
        ("1. Sizing", "All images resized to standard 224 x 224 pixels."),
        ("2. Gaussian Blur", "Light 5x5 smoothing to remove line artifacts without altering path shape."),
        ("3. RGB Format", "Consistent 3-channel color format across all input images."),
        ("4. Normalization", "Pixel values scaled to [0, 1] range for stable gradient flow."),
        ("5. Color Gradient Preservation", "Gaze heatmap gradients preserved to retain spatial attention intensity.")
    ]
    for pr_h, pr_b in prep_steps:
        pph = tf_pr.add_paragraph()
        pph.text = pr_h
        pph.font.size = Pt(13)
        pph.font.bold = True
        pph.font.color.rgb = DARK_BG
        ppb = tf_pr.add_paragraph()
        ppb.text = pr_b
        ppb.font.size = Pt(11)
        ppb.font.color.rgb = TEXT_DARK
        ppb.space_after = Pt(6)

    # ==========================================
    # SLIDE 4: NEW SLIDE - FRONTEND WORKING & LIVE DEMONSTRATION WITH PHOTOS
    # ==========================================
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Live System Demonstration: From Scanpath Input to Diagnostic Output", category="APPLICATION IN ACTION")

    # Left Container: Upload Stage
    add_card(slide4, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.4))
    tb_up = slide4.shapes.add_textbox(Inches(1.0), Inches(1.75), Inches(5.3), Inches(0.6))
    tf_up = tb_up.text_frame
    p_uph = tf_up.paragraphs[0]
    p_uph.text = "STEP 1: SCANPATH INTAKE & PREVIEW"
    p_uph.font.size = Pt(12)
    p_uph.font.bold = True
    p_uph.font.color.rgb = PRIMARY

    if os.path.exists(img_upload_path):
        slide4.shapes.add_picture(img_upload_path, Inches(1.0), Inches(2.2), width=Inches(5.3))

    tb_up_txt = slide4.shapes.add_textbox(Inches(1.0), Inches(5.2), Inches(5.3), Inches(1.6))
    tf_ut = tb_up_txt.text_frame
    tf_ut.word_wrap = True
    pu1 = tf_ut.paragraphs[0]
    pu1.text = "• Clinician drags any scanpath image (PNG/JPG) or clicks 'Browse files'."
    pu1.font.size = Pt(11)
    pu1.font.color.rgb = TEXT_DARK
    pu2 = tf_ut.add_paragraph()
    pu2.text = "• The portal validates file size (28.4 KB) and generates an instant live thumbnail."
    pu2.font.size = Pt(11)
    pu2.font.color.rgb = TEXT_MUTED
    pu3 = tf_ut.add_paragraph()
    pu3.text = "• Clicking 'Run Analysis' triggers the backend ML inference pipeline asynchronously."
    pu3.font.size = Pt(11)
    pu3.font.color.rgb = TEXT_MUTED

    # Right Container: Output Stage
    add_card(slide4, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.4), bg_color=CARD_BG, border_color=PRIMARY)
    tb_out = slide4.shapes.add_textbox(Inches(7.0), Inches(1.75), Inches(5.3), Inches(0.6))
    tf_out = tb_out.text_frame
    p_outh = tf_out.paragraphs[0]
    p_outh.text = "STEP 2: REAL-TIME DIAGNOSTIC DASHBOARD"
    p_outh.font.size = Pt(12)
    p_outh.font.bold = True
    p_outh.font.color.rgb = PRIMARY

    if os.path.exists(img_output_path):
        slide4.shapes.add_picture(img_output_path, Inches(7.0), Inches(2.2), width=Inches(5.3))

    # ==========================================
    # SLIDE 5: Stacking Ensemble Architecture
    # ==========================================
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Stacking Ensemble Learning Architecture")

    ens_cards = [
        ("BASE MODEL 1", "Support Vector Machine (SVM)", "RBF Kernel",
         ["Constructs non-linear decision boundaries in high-dimensional feature space.",
          "Highly effective with PCA-reduced dimensions.",
          "Learns global separation margin between ASD and Non-ASD scanpaths."]),
        ("BASE MODEL 2", "K-Nearest Neighbors (KNN)", "k = 5 Neighbors",
         ["Instance-based lazy learner modeling localized density in feature space.",
          "Optimized at k=5 neighbors.",
          "Complements SVM by focusing on local cluster similarities."]),
        ("META MODEL", "Logistic Regression", "Calibrated Output",
         ["Trained on predictions of SVM and KNN using 5-Fold cross validation.",
          "Learns optimal weighting to correct individual model errors.",
          "Outputs calibrated confidence probability scores for the doctor."])
    ]
    for i, (m_tag, m_name, m_tech, m_pts) in enumerate(ens_cards):
        ex = Inches(0.8 + i * 3.9)
        add_card(slide5, ex, Inches(1.6), Inches(3.6), Inches(5.3))
        tb_e = slide5.shapes.add_textbox(ex + Inches(0.2), Inches(1.8), Inches(3.2), Inches(4.8))
        tf_e = tb_e.text_frame
        tf_e.word_wrap = True
        p_et = tf_e.paragraphs[0]
        p_et.text = m_tag
        p_et.font.size = Pt(11)
        p_et.font.bold = True
        p_et.font.color.rgb = PRIMARY

        p_en = tf_e.add_paragraph()
        p_en.text = m_name
        p_en.font.size = Pt(16)
        p_en.font.bold = True
        p_en.font.color.rgb = DARK_BG

        p_ec = tf_e.add_paragraph()
        p_ec.text = m_tech
        p_ec.font.size = Pt(12)
        p_ec.font.bold = True
        p_ec.font.color.rgb = ACCENT
        p_ec.space_after = Pt(10)

        for pt in m_pts:
            pe_pt = tf_e.add_paragraph()
            pe_pt.text = "• " + pt
            pe_pt.font.size = Pt(12)
            pe_pt.font.color.rgb = TEXT_DARK
            pe_pt.space_after = Pt(8)

    # ==========================================
    # SLIDE 6: EXACT RESULTS & METRICS TABLE
    # ==========================================
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Experimental Results & Performance Comparison")

    add_card(slide6, Inches(0.8), Inches(1.6), Inches(7.2), Inches(5.3))
    tb_res = slide6.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(6.8), Inches(4.8))
    tf_r = tb_res.text_frame
    tf_r.word_wrap = True
    p_rh = tf_r.paragraphs[0]
    p_rh.text = "PERFORMANCE METRICS TABLE"
    p_rh.font.size = Pt(12)
    p_rh.font.bold = True
    p_rh.font.color.rgb = PRIMARY
    p_rh.space_after = Pt(10)

    rows = [
        ("Model", "Accuracy", "Precision", "Recall", "F1-Score"),
        ("MobileNet + PCA + Stacking", "92.69%", "93.1%", "92.4%", "92.7%"),
        ("ResNet50 + PCA + Stacking", "90.12%", "90.8%", "89.6%", "90.2%"),
        ("MobileNet + SVM (no stack)", "87.34%", "87.9%", "86.8%", "87.3%"),
        ("ResNet50 + KNN", "84.21%", "84.5%", "83.9%", "84.2%"),
        ("Baseline (no PCA)", "78.56%", "79.1%", "77.8%", "78.4%")
    ]
    for row in rows:
        pr = tf_r.add_paragraph()
        pr.text = f"{row[0]:<28} | {row[1]:<10} | {row[2]:<10} | {row[3]:<8} | {row[4]}"
        pr.font.size = Pt(11.5)
        pr.font.name = "Courier New"
        if row[0] == "Model":
            pr.font.bold = True
            pr.font.color.rgb = DARK_BG
        elif "92.69%" in row[1]:
            pr.font.bold = True
            pr.font.color.rgb = PRIMARY
        else:
            pr.font.color.rgb = TEXT_DARK
        pr.space_after = Pt(8)

    add_card(slide6, Inches(8.3), Inches(1.6), Inches(4.2), Inches(5.3), bg_color=CARD_BG, border_color=PRIMARY)
    tb_side = slide6.shapes.add_textbox(Inches(8.5), Inches(1.8), Inches(3.8), Inches(4.8))
    tf_s = tb_side.text_frame
    tf_s.word_wrap = True
    p_sh = tf_s.paragraphs[0]
    p_sh.text = "KEY TAKEAWAYS"
    p_sh.font.size = Pt(12)
    p_sh.font.bold = True
    p_sh.font.color.rgb = PRIMARY
    p_sh.space_after = Pt(10)

    side_pts = [
        ("92.69% Overall Accuracy", "Stacking ensemble delivered the highest accuracy and balanced precision-recall."),
        ("PCA Impact", "Jumped from 78.56% (Baseline) to 92.69% by filtering noise and redundant dimensions."),
        ("Ensemble Advantage", "Stacking outperformed single SVM by +5.35% and single KNN by +8.48%."),
        ("Speed-Accuracy Winner", "MobileNet surpassed ResNet50 while being 6x smaller in model parameters.")
    ]
    for sh, sb in side_pts:
        psh = tf_s.add_paragraph()
        psh.text = sh
        psh.font.size = Pt(13)
        psh.font.bold = True
        psh.font.color.rgb = DARK_BG
        psb = tf_s.add_paragraph()
        psb.text = sb
        psb.font.size = Pt(11)
        psb.font.color.rgb = TEXT_MUTED
        psb.space_after = Pt(8)

    # ==========================================
    # SLIDE 7: Web Application Features
    # ==========================================
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Clinical Web Application: Translating Research to Practice")

    add_card(slide7, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.3))
    tb_web = slide7.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.0), Inches(4.8))
    tf_w = tb_web.text_frame
    tf_w.word_wrap = True
    p_wh = tf_w.paragraphs[0]
    p_wh.text = "WHAT WE BUILT BEYOND THE NOTEBOOK"
    p_wh.font.size = Pt(12)
    p_wh.font.bold = True
    p_wh.font.color.rgb = PRIMARY
    p_wh.space_after = Pt(10)

    web_pts = [
        ("From Code Cells to Web Interface", "Clinicians do not need to install Python, Colab, or run code. They interact through a clean web portal."),
        ("Drag-and-Drop Scanpath Intake", "Uploads any gaze trace image with instant client-side preview."),
        ("Real-time Inference Engine", "Flask REST backend computes predictions in <400ms using the exact 4-step pipeline."),
        ("Production Server Architecture", "Configured with Gunicorn WSGI server and pushed to GitHub for free cloud hosting.")
    ]
    for wh, wb in web_pts:
        pwh = tf_w.add_paragraph()
        pwh.text = "• " + wh + ":"
        pwh.font.size = Pt(13)
        pwh.font.bold = True
        pwh.font.color.rgb = DARK_BG
        pwb = tf_w.add_paragraph()
        pwb.text = "   " + wb
        pwb.font.size = Pt(11)
        pwb.font.color.rgb = TEXT_MUTED
        pwb.space_after = Pt(8)

    add_card(slide7, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.3), bg_color=LIGHT_BG)
    tb_flow = slide7.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(4.8))
    tf_fl = tb_flow.text_frame
    tf_fl.word_wrap = True
    p_flh = tf_fl.paragraphs[0]
    p_flh.text = "STEP-BY-STEP USER WORKFLOW"
    p_flh.font.size = Pt(12)
    p_flh.font.bold = True
    p_flh.font.color.rgb = PRIMARY
    p_flh.space_after = Pt(10)

    steps = [
        ("1. Drop Scanpath", "Clinician uploads gaze image into the dropzone."),
        ("2. Backend Execution", "Backend applies MobileNet -> PCA -> Stacking Ensemble."),
        ("3. Instant Status Badge", "Displays 'ASD Pattern Indicated' or 'Neurotypical Pattern'."),
        ("4. Confidence Gauge", "Visual animated meter showing exact probability percentage."),
        ("5. Explainability Breakdown", "Presents 3 clear visual reasons explaining the diagnosis.")
    ]
    for s_num, s_desc in steps:
        psh = tf_fl.add_paragraph()
        psh.text = s_num
        psh.font.size = Pt(13)
        psh.font.bold = True
        psh.font.color.rgb = DARK_BG
        psb = tf_fl.add_paragraph()
        psb.text = s_desc
        psb.font.size = Pt(11)
        psb.font.color.rgb = TEXT_DARK
        psb.space_after = Pt(6)

    # ==========================================
    # SLIDE 8: Explainability (Simple Words)
    # ==========================================
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "Explainability: 'Why This Patient?' (Simple English)")

    exp_cards = [
        ("CHECK 1", "Where Did They Look?", "Peripheral vs Center",
         "Children with Autism often fixate on the outer borders or background of an image rather than central objects. Our system measures if gaze points are clustered away from primary visual anchors."),
        ("CHECK 2", "Eye Contact & Faces", "Face Avoidance (<30% Dwell)",
         "When shown images of people, neurotypical children naturally look at eyes and mouths (~68% dwell). In ASD scanpaths, this drops below 30%, showing clear avoidance of social facial cues."),
        ("CHECK 3", "Smooth vs. Jerky Jumps", "Restless Eye Movements",
         "Instead of following smooth, calm lines, children with Autism often make short, scattered, zigzag eye movements. The system flags this high path jitter as a key indicator.")
    ]
    for i, (b_tag, b_name, b_stat, b_desc) in enumerate(exp_cards):
        bx = Inches(0.8 + i * 3.9)
        add_card(slide8, bx, Inches(1.6), Inches(3.6), Inches(5.3))
        tb_b = slide8.shapes.add_textbox(bx + Inches(0.2), Inches(1.8), Inches(3.2), Inches(4.8))
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
        p_cl.text = "Why this matters: Doctors trust an AI that explains its reasoning, rather than a mystery score."
        p_cl.font.size = Pt(11)
        p_cl.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 9: System Handbook
    # ==========================================
    slide9 = prs.slides.add_slide(blank_layout)
    add_header(slide9, "System Handbook: Instructions for Users & Clinic Staff")

    hb_cards = [
        ("FILE FORMATS", "What Images Can Be Uploaded?", 
         ["Supports standard PNG, JPG, and JPEG files.",
          "Compatible with raw gaze plots, line scanpaths, and colored heatmaps.",
          "Any resolution works—the server automatically standardizes sizing to 224x224."]),
        ("RELIABILITY", "Built-In Safety & Fallback", 
         ["If trained model weights are missing, the app NEVER crashes.",
          "It automatically runs in 'simulation mode' using contour analysis to allow offline demonstrations.",
          "A clear banner informs the doctor whether live weights or simulation mode is active."]),
        ("RESPONSIBILITY", "Screening Aid vs Final Diagnosis", 
         ["The portal explicitly states it is a screening assistant, not a replacement for a doctor.",
          "Provides a confidence score so clinicians can see if a result is borderline or strong.",
          "Recommends formal ADOS-2 clinical testing when a high-risk pattern is detected."])
    ]
    for i, (h_tag, h_title, h_points) in enumerate(hb_cards):
        hx = Inches(0.8 + i * 3.9)
        add_card(slide9, hx, Inches(1.6), Inches(3.6), Inches(5.3))
        tb_h = slide9.shapes.add_textbox(hx + Inches(0.2), Inches(1.8), Inches(3.2), Inches(4.8))
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
    # SLIDE 10: Conclusion (Dark Theme)
    # ==========================================
    slide10 = prs.slides.add_slide(blank_layout)
    bg10 = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg10.fill.solid()
    bg10.fill.fore_color.rgb = DARK_BG
    bg10.line.fill.background()

    tb_end = slide10.shapes.add_textbox(Inches(1.0), Inches(1.3), Inches(11.3), Inches(5.0))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True
    pe = tf_end.paragraphs[0]
    pe.text = "CONCLUSION & FUTURE WORK"
    pe.font.size = Pt(14)
    pe.font.bold = True
    pe.font.color.rgb = PRIMARY
    pe.space_after = Pt(8)

    pe2 = tf_end.add_paragraph()
    pe2.text = "Project Summary:"
    pe2.font.size = Pt(26)
    pe2.font.bold = True
    pe2.font.color.rgb = TEXT_LIGHT
    pe2.space_after = Pt(10)

    recap_points = [
        "1. High Diagnostic Accuracy: Achieved 92.69% accuracy on benchmark eye-tracking data using MobileNet + PCA + Stacking Ensemble.",
        "2. Understandable Explanations: Replaced heavy jargon with 3 intuitive visual checks (where they look, face avoidance, and jerky jumps).",
        "3. Real Working Web Portal: Built an interactive, full-stack application for doctors with instant drag-and-drop inference.",
        "4. Deployment Ready: Fully tracked on GitHub (nandini-rathod/asd-eye-tracking) with Gunicorn and configured for Render cloud hosting.",
        "5. Future Expansion: Real-time eye-tracker integration, Vision Transformers (ViT), and multimodal fusion with EEG data."
    ]
    for rpt in recap_points:
        pr = tf_end.add_paragraph()
        pr.text = rpt
        pr.font.size = Pt(14)
        pr.font.color.rgb = RGBColor(226, 232, 240)
        pr.space_after = Pt(8)

    prs.save(filename)
    print(f"Presentation successfully updated with photo demonstration slide at {filename}")

if __name__ == "__main__":
    create_deck(sys.argv[1] if len(sys.argv) > 1 else "ASD_EyeTracking_Presentation.pptx")
