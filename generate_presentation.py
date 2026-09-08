"""
Generates a polished PowerPoint Presentation (PPTX) for Olist E-Commerce Analytics.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_presentation(output_path):
    prs = Presentation()
    # Set 16:9 widescreen slides
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette
    COLOR_PRIMARY = RGBColor(15, 23, 42)      # Deep Navy #0F172A
    COLOR_ACCENT = RGBColor(2, 132, 199)      # Sky Blue #0284C7
    COLOR_SUCCESS = RGBColor(16, 185, 129)    # Emerald #10B981
    COLOR_WARNING = RGBColor(245, 158, 11)    # Amber #F59E0B
    COLOR_DANGER = RGBColor(239, 68, 68)      # Red #EF4444
    COLOR_CARD_BG = RGBColor(248, 250, 252)   # Slate 50
    COLOR_CARD_BORDER = RGBColor(203, 213, 225)# Slate 300
    COLOR_TEXT_DARK = RGBColor(30, 41, 59)    # Slate 800
    COLOR_TEXT_MUTED = RGBColor(100, 116, 139)# Slate 500
    COLOR_WHITE = RGBColor(255, 255, 255)

    def add_header(slide, title_text, subtitle_text):
        # Top header bar
        header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        header_box.fill.solid()
        header_box.fill.fore_color.rgb = COLOR_PRIMARY
        header_box.line.color.rgb = COLOR_PRIMARY
        
        # Title text
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.6)
        tf.margin_top = Inches(0.15)
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        
        p2 = tf.add_paragraph()
        p2.text = subtitle_text
        p2.font.size = Pt(11)
        p2.font.color.rgb = RGBColor(148, 163, 184)

    def add_card(slide, left, top, width, height, title, value, subtitle="", val_color=COLOR_ACCENT):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.15)

        p1 = tf.paragraphs[0]
        p1.text = title.upper()
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_TEXT_MUTED
        p1.alignment = PP_ALIGN.CENTER

        p2 = tf.add_paragraph()
        p2.text = value
        p2.font.size = Pt(20)
        p2.font.bold = True
        p2.font.color.rgb = val_color
        p2.alignment = PP_ALIGN.CENTER

        if subtitle:
            p3 = tf.add_paragraph()
            p3.text = subtitle
            p3.font.size = Pt(9.5)
            p3.font.color.rgb = COLOR_TEXT_DARK
            p3.alignment = PP_ALIGN.CENTER

    # -------------------------------------------------------------
    # SLIDE 1: TITLE SLIDE
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_PRIMARY
    bg1.line.color.rgb = COLOR_PRIMARY

    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(3.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_badge = tf1.paragraphs[0]
    p_badge.text = "EXECUTIVE BUSINESS INTELLIGENCE REPORT"
    p_badge.font.size = Pt(13)
    p_badge.font.bold = True
    p_badge.font.color.rgb = COLOR_ACCENT

    p_main = tf1.add_paragraph()
    p_main.text = "Brazilian E-Commerce (Olist)\nData Analytics & Strategic Insights"
    p_main.font.size = Pt(36)
    p_main.font.bold = True
    p_main.font.color.rgb = COLOR_WHITE

    p_sub = tf1.add_paragraph()
    p_sub.text = "تحليل شامل للبيانات: تنظيف البيانات، مؤشرات الأداء الرئيسية، وسلوك العملاء، وتوصيات النمو"
    p_sub.font.size = Pt(16)
    p_sub.font.color.rgb = RGBColor(203, 213, 225)

    # -------------------------------------------------------------
    # SLIDE 2: DATASET OVERVIEW (قصة البيانات ببساطة)
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "1. What is the Data? | فهم طبيعة وقصة البيانات", "Overview of Olist Brazilian Marketplace & Dataset Architecture")

    # 3 Content Blocks
    box_w = Inches(3.7)
    box_h = Inches(5.6)
    
    # Block 1: Who is Olist?
    b1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4), box_w, box_h)
    b1.fill.solid()
    b1.fill.fore_color.rgb = COLOR_CARD_BG
    b1.line.color.rgb = COLOR_CARD_BORDER
    tf_b1 = b1.text_frame
    tf_b1.word_wrap = True
    tf_b1.margin_left = Inches(0.25)
    tf_b1.margin_top = Inches(0.25)
    
    p = tf_b1.paragraphs[0]
    p.text = "🛒 ما هي منصة Olist؟"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY
    
    bullets1 = [
        "منصة تجارة إلكترونية برازيلية رائدة تربط بين آلاف البائعين المستقلين وملايين المشترين في البرازيل.",
        "تتولى المنصة إدارة الطلبات، عمليات الدفع، متابعة الشحن، وتقييمات العملاء.",
        "البيانات تغطي الفترة من 2016 إلى 2018 بإجمالي أكثر من 100 ألف طلب تجاري."
    ]
    for b in bullets1:
        p = tf_b1.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Block 2: 9 Relational Tables
    b2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.7), Inches(1.4), box_w, box_h)
    b2.fill.solid()
    b2.fill.fore_color.rgb = COLOR_CARD_BG
    b2.line.color.rgb = COLOR_CARD_BORDER
    tf_b2 = b2.text_frame
    tf_b2.word_wrap = True
    tf_b2.margin_left = Inches(0.25)
    tf_b2.margin_top = Inches(0.25)

    p = tf_b2.paragraphs[0]
    p.text = "🔗 الجداول الـ 9 المترابطة"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    bullets2 = [
        "1. Orders (الطلبات وحالات الشحن ومواعيد التوصيل)",
        "2. Order Items (المنتجات وسعر كل منتج ومصاريف الشحن)",
        "3. Products (كتالوج المنتجات وتصنيفاتها وأبعادها)",
        "4. Payments (طرق الدفع والأقساط وقيم المبيعات)",
        "5. Reviews (تقييمات العملاء 1-5 نجوم وتعليقاتهم)",
        "6. Customers & Sellers (بيانات المشترين والبائعين والمدن)",
        "7. Geolocation & Translation (الترجمة الإنجليزية والإحداثيات)"
    ]
    for b in bullets2:
        p = tf_b2.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Block 3: Cleaning Operations
    b3 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.8), Inches(1.4), box_w, box_h)
    b3.fill.solid()
    b3.fill.fore_color.rgb = COLOR_CARD_BG
    b3.line.color.rgb = COLOR_CARD_BORDER
    tf_b3 = b3.text_frame
    tf_b3.word_wrap = True
    tf_b3.margin_left = Inches(0.25)
    tf_b3.margin_top = Inches(0.25)

    p = tf_b3.paragraphs[0]
    p.text = "🧹 ما قمنا بتنظيفه بالبايثون"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    bullets3 = [
        "تحويل جميع التواريخ ومعالجة القيم المفقودة واستخراج شهور وسنوات الشراء.",
        "حساب سرعة الشحن الفعلية وتحديد الطلبات المتأخرة عن موعدها بدقة.",
        "ترجمة كافة تصنيفات المنتجات من البرتغالية للإنجليزية ومعالجة غير المصنف.",
        "دمج وتجميع طرق الدفع والأقساط المتعددة على مستوى كل طلب.",
        "دمج الجداول في جدول تحليلي رئيسي موحد Master Dataset."
    ]
    for b in bullets3:
        p = tf_b3.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # -------------------------------------------------------------
    # SLIDE 3: EXECUTIVE KPI DASHBOARD
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "2. Executive KPI Dashboard | لوحة المؤشرات الرئيسية", "High-Level Business Performance & Vital Health Metrics")

    # 6 Top KPI Cards
    card_w = Inches(3.8)
    card_h = Inches(1.5)

    add_card(s3, Inches(0.6), Inches(1.4), card_w, card_h, "إجمالي الإيرادات (Total GMV)", "R$ 15.42 Million", "إجمالي قيمة المبيعات المكتملة", COLOR_ACCENT)
    add_card(s3, Inches(4.75), Inches(1.4), card_w, card_h, "الطلبات المسلمة بنجاح", "96,478 Orders", "97.0% نسبة نجاح توصيل الطلبات", COLOR_PRIMARY)
    add_card(s3, Inches(8.9), Inches(1.4), card_w, card_h, "عدد العملاء الفريدين", "96,096 Customers", "قاعدة عملاء واسعة ومتنوعة", COLOR_PRIMARY)

    add_card(s3, Inches(0.6), Inches(3.1), card_w, card_h, "متوسط قيمة الطلب (AOV)", "R$ 159.83", "Average Order Value للسلة", COLOR_ACCENT)
    add_card(s3, Inches(4.75), Inches(3.1), card_w, card_h, "متوسط سرعة التوصيل", "12.6 Days", "زمن وصول الشحنة من الشراء للباب", COLOR_WARNING)
    add_card(s3, Inches(8.9), Inches(3.1), card_w, card_h, "نسبة التوصيل في الموعد", "91.9% On-Time", "التزام بجدول التوصيل التقديري", COLOR_SUCCESS)

    # Key Takeaways Box
    summary_box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(4.8), Inches(12.1), Inches(2.2))
    summary_box.fill.solid()
    summary_box.fill.fore_color.rgb = COLOR_CARD_BG
    summary_box.line.color.rgb = COLOR_CARD_BORDER
    tf_s = summary_box.text_frame
    tf_s.word_wrap = True
    tf_s.margin_left = Inches(0.3)
    tf_s.margin_top = Inches(0.2)

    p = tf_s.paragraphs[0]
    p.text = "📌 ملخص الأداء التنفيذي السريع:"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    bullets_kpi = [
        "نمو قوي ومستمر في حجم المبيعات والإيرادات الشهرية خلال 2017 و2018.",
        "أداء لوجستي جيد بنسبة التزام 91.9%، لكن الطلبات المتأخرة (8.1%) تمثل المصدر الأكبر لشكاوى وتقييمات العملاء السلبية.",
        "رضا عام إيجابي بمعدل 4.16 / 5.0 مع تركز 77.1% من التقييمات في فئتي 4 و5 نجوم."
    ]
    for b in bullets_kpi:
        p = tf_s.add_paragraph()
        p.text = "✔ " + b
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_TEXT_DARK

    # -------------------------------------------------------------
    # SLIDE 4: REPORT 1 - SALES & REVENUE GROWTH
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "3. Sales & Revenue Trajectory | اتجاهات المبيعات والنمو", "Report 1: Monthly Trends, Seasonality Spikes & Basket Dynamics")

    # Left Column: Insights
    b_l = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4), Inches(5.8), Inches(5.6))
    b_l.fill.solid()
    b_l.fill.fore_color.rgb = COLOR_CARD_BG
    b_l.line.color.rgb = COLOR_CARD_BORDER
    tf_l = b_l.text_frame
    tf_l.word_wrap = True
    tf_l.margin_left = Inches(0.3)
    tf_l.margin_top = Inches(0.25)

    p = tf_l.paragraphs[0]
    p.text = "📈 قراءة حركة المبيعات والإيرادات:"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    bullets_r1 = [
        "البداية (يناير 2017): انطلقت المبيعات بـ 750 طلب بإيراد 120 ألف ريال برازيلي.",
        "الذروة والنمو السريع: تضاعفت المبيعات تدريجياً حتى وصلت لأكثر من 7,000 طلب شهرياً بإيراد تخطى 1.1 مليون ريال شهرياً في 2018.",
        "طفرة البلاك فرايداي (نوفمبر 2017): قفزة قياسية بأكثر من 7,200 طلب وإيراد 1.15 مليون ريال نتيجة العروض الموسمية.",
        "حصة الشحن (Freight Share): تمثل مصاريف الشحن ما بين 15% إلى 17% من إجمالي قيمة ما يدفعه العميل."
    ]
    for b in bullets_r1:
        p = tf_l.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Right Column: Key Data Table
    table_shape = s4.shapes.add_table(6, 4, Inches(6.7), Inches(1.4), Inches(6.0), Inches(5.0))
    table = table_shape.table
    table.columns[0].width = Inches(1.5)
    table.columns[1].width = Inches(1.3)
    table.columns[2].width = Inches(1.7)
    table.columns[3].width = Inches(1.5)

    headers = ["الفترة", "الطلبات", "الإيرادات", "AOV"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

    sample_rows = [
        ["2017-01", "750", "R$ 120,312", "R$ 160.42"],
        ["2017-07", "3,872", "R$ 592,302", "R$ 152.97"],
        ["2017-11 (BF)", "7,289", "R$ 1,153,393", "R$ 158.24"],
        ["2018-03", "7,003", "R$ 1,119,654", "R$ 159.88"],
        ["2018-08", "6,351", "R$ 985,414", "R$ 155.16"],
    ]
    for r_idx, row_data in enumerate(sample_rows):
        for c_idx, val in enumerate(row_data):
            cell = table.cell(r_idx + 1, c_idx)
            cell.text = val
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(11)
            p.alignment = PP_ALIGN.CENTER

    # -------------------------------------------------------------
    # SLIDE 5: REPORT 2 - PRODUCT CATEGORIES
    # -------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "4. Product Category Merchandising | أداء فئات المنتجات", "Report 2: Top Revenue Drivers, Units Sold & Category Ratings")

    # Table of Top 5 Categories
    t5_shape = s5.shapes.add_table(6, 6, Inches(0.6), Inches(1.4), Inches(12.1), Inches(3.2))
    t5 = t5_shape.table
    t5.columns[0].width = Inches(3.0)
    t5.columns[1].width = Inches(1.6)
    t5.columns[2].width = Inches(2.2)
    t5.columns[3].width = Inches(1.6)
    t5.columns[4].width = Inches(1.8)
    t5.columns[5].width = Inches(1.9)

    c_headers = ["تصنيف المنتج (Category)", "القطع المباعة", "إجمالي الإيراد", "الحصة من المبيعات", "متوسط السعر", "تقييم العملاء"]
    for i, h in enumerate(c_headers):
        cell = t5.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

    cat_rows = [
        ["Health & Beauty (الصحة والجمال)", "9,670", "R$ 1,258,681", "9.3%", "R$ 130.16", "4.14 ⭐"],
        ["Watches & Gifts (الساعات والهدايا)", "5,624", "R$ 1,205,005", "8.9%", "R$ 201.99", "4.02 ⭐"],
        ["Bed Bath Table (المفروشات والمنزل)", "11,115", "R$ 1,036,988", "7.7%", "R$ 93.30", "3.88 ⭐"],
        ["Sports Leisure (الرياضة والترفيه)", "8,641", "R$ 988,048", "7.3%", "R$ 114.34", "4.11 ⭐"],
        ["Computers Accessories (إلكترونيات)", "7,827", "R$ 911,954", "6.7%", "R$ 116.51", "3.93 ⭐"],
    ]
    for r_idx, row_data in enumerate(cat_rows):
        for c_idx, val in enumerate(row_data):
            cell = t5.cell(r_idx + 1, c_idx)
            cell.text = val
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(11)
            p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT

    # Highlights box below
    cat_summary = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(4.9), Inches(12.1), Inches(2.1))
    cat_summary.fill.solid()
    cat_summary.fill.fore_color.rgb = COLOR_CARD_BG
    cat_summary.line.color.rgb = COLOR_CARD_BORDER
    tf_c = cat_summary.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = Inches(0.3)
    tf_c.margin_top = Inches(0.15)

    p = tf_c.paragraphs[0]
    p.text = "💡 أهم الملاحظات على المنتجات:"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    bullets_cat = [
        "الصحة والجمال والساعات هما القائدان في توليد الأرباح بأعلى إيرادات وأسعار متوسطة مرتفعة.",
        "قسم المفروشات (Bed Bath Table) هو الأكثر مبيعاً من حيث عدد القطع (11,115 قطعة) ولكنه يسجل تقييماً أقل (3.88) بسبب توقعات خامات الأقمشة.",
        "أعلى 5 فئات منتجات تمثل حوالي 40% من إجمالي مبيعات المنصة بالكامل."
    ]
    for b in bullets_cat:
        p = tf_c.add_paragraph()
        p.text = "✔ " + b
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # -------------------------------------------------------------
    # SLIDE 6: REPORT 3 - REGIONAL LOGISTICS
    # -------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "5. Regional Demand & Logistics | التوزيع الجغرافي والشحن", "Report 3: State-by-State Volumes, Transit Speeds & Delivery Bottlenecks")

    # Left Box: Southeast Dominance
    b_se = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4), Inches(5.8), Inches(5.6))
    b_se.fill.solid()
    b_se.fill.fore_color.rgb = COLOR_CARD_BG
    b_se.line.color.rgb = COLOR_CARD_BORDER
    tf_se = b_se.text_frame
    tf_se.word_wrap = True
    tf_se.margin_left = Inches(0.3)
    tf_se.margin_top = Inches(0.25)

    p = tf_se.paragraphs[0]
    p.text = "🏙️ تركز الطلب في الجنوب الشرقي (Southeast):"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    bullets_geo1 = [
        "ولاية ساو باولو (SP): تستحوذ بمفردها على 41.7% من إجمالي طلبات البرازيل (40,244 طلب).",
        "سرعة الشحن في SP: الأسرع على الإطلاق بمتوسط 8.3 أيام فقط وبنسبة التزام 95.2%.",
        "الولايات المجاورة (RJ, MG): تأتي في المركزين الثاني والثالث بـ 12.8 ألف و 11.3 ألف طلب.",
        "المنطقة الجنوبية والجنوبية الشرقية تمثل معاً أكثر من 78% من حجم التجارة بالكامل."
    ]
    for b in bullets_geo1:
        p = tf_se.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Right Box: Remote Regions Challenges
    b_no = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.7), Inches(1.4), Inches(6.0), Inches(5.6))
    b_no.fill.solid()
    b_no.fill.fore_color.rgb = COLOR_CARD_BG
    b_no.line.color.rgb = COLOR_CARD_BORDER
    tf_no = b_no.text_frame
    tf_no.word_wrap = True
    tf_no.margin_left = Inches(0.3)
    tf_no.margin_top = Inches(0.25)

    p = tf_no.paragraphs[0]
    p.text = "⚠️ التحديات اللوجستية في الشمال والشمال الشرقي:"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_DANGER

    bullets_geo2 = [
        "الولايات البعيدة (RR, AP, AM, AL, MA): يستغرق التوصيل فيها من 24 إلى 29 يوماً في المتوسط.",
        "ارتفاع مصاريف الشحن: تكلفة الشحن للمناطق البعيدة تصل لضعف تكلفة الشحن في ساو باولو (R$ 40+ مقابل R$ 15).",
        "انخفاض نسبة التوصيل في الموعد: تهبط نسبة الالتزام في بعض ولايات الشمال إلى 75%-80% مما يضر بتقييمات العملاء.",
        "فرصة استراتيجية: الحاجة لإنشاء مستودعات توزيع إقليمية (Regional Hubs) لتقليل زمن الشحن."
    ]
    for b in bullets_geo2:
        p = tf_no.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # -------------------------------------------------------------
    # SLIDE 7: REPORT 4 - PAYMENT METHODS
    # -------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "6. Payment Methods & Financing | طرق الدفع وسلوك التقسيط", "Report 4: Payment Shares, Installment Depths & Ticket Sizes")

    # 4 Cards for payment methods
    p_w = Inches(2.8)
    p_h = Inches(3.4)

    # Card 1: Credit Card
    c_cc = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4), p_w, p_h)
    c_cc.fill.solid()
    c_cc.fill.fore_color.rgb = COLOR_CARD_BG
    c_cc.line.color.rgb = COLOR_ACCENT
    c_cc.line.width = Pt(2)
    tf_cc = c_cc.text_frame
    tf_cc.word_wrap = True
    tf_cc.margin_left = Inches(0.2)
    tf_cc.margin_top = Inches(0.2)
    p = tf_cc.paragraphs[0]
    p.text = "💳 البطاقات الائتمانية\n(Credit Card)"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_PRIMARY
    p.alignment = PP_ALIGN.CENTER
    bullets_cc = ["الحصة: 75.4% من الإيرادات", "القيمة: R$ 12.54 Million", "متوسط التقسيط: 3.5 أقساط", "متوسط الطلب: R$ 163.32"]
    for b in bullets_cc:
        p = tf_cc.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Card 2: Boleto
    c_bo = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.7), Inches(1.4), p_w, p_h)
    c_bo.fill.solid()
    c_bo.fill.fore_color.rgb = COLOR_CARD_BG
    c_bo.line.color.rgb = COLOR_CARD_BORDER
    tf_bo = c_bo.text_frame
    tf_bo.word_wrap = True
    tf_bo.margin_left = Inches(0.2)
    tf_bo.margin_top = Inches(0.2)
    p = tf_bo.paragraphs[0]
    p.text = "📄 فواتير بوليتو\n(Boleto Bancário)"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_PRIMARY
    p.alignment = PP_ALIGN.CENTER
    bullets_bo = ["الحصة: 19.3% من الإيرادات", "القيمة: R$ 2.87 Million", "دفع كاش مباشر بدون تقسيط", "متوسط الطلب: R$ 145.03"]
    for b in bullets_bo:
        p = tf_bo.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Card 3: Voucher
    c_vo = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.4), p_w, p_h)
    c_vo.fill.solid()
    c_vo.fill.fore_color.rgb = COLOR_CARD_BG
    c_vo.line.color.rgb = COLOR_CARD_BORDER
    tf_vo = c_vo.text_frame
    tf_vo.word_wrap = True
    tf_vo.margin_left = Inches(0.2)
    tf_vo.margin_top = Inches(0.2)
    p = tf_vo.paragraphs[0]
    p.text = "🎟️ كوبونات وقسائم\n(Voucher)"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_PRIMARY
    p.alignment = PP_ALIGN.CENTER
    bullets_vo = ["الحصة: 2.3% من الإيرادات", "القيمة: R$ 380 Thousand", "تستخدم لتعويض أو خصم", "متوسط الطلب: R$ 65.70"]
    for b in bullets_vo:
        p = tf_vo.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Card 4: Debit Card
    c_dc = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.9), Inches(1.4), p_w, p_h)
    c_dc.fill.solid()
    c_dc.fill.fore_color.rgb = COLOR_CARD_BG
    c_dc.line.color.rgb = COLOR_CARD_BORDER
    tf_dc = c_dc.text_frame
    tf_dc.word_wrap = True
    tf_dc.margin_left = Inches(0.2)
    tf_dc.margin_top = Inches(0.2)
    p = tf_dc.paragraphs[0]
    p.text = "💳 بطاقات الخصم\n(Debit Card)"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_PRIMARY
    p.alignment = PP_ALIGN.CENTER
    bullets_dc = ["الحصة: 1.3% من الإيرادات", "القيمة: R$ 218 Thousand", "خصم مباشر من الحساب", "متوسط الطلب: R$ 142.57"]
    for b in bullets_dc:
        p = tf_dc.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Insights Box
    pay_insight = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(5.1), Inches(12.1), Inches(1.9))
    pay_insight.fill.solid()
    pay_insight.fill.fore_color.rgb = COLOR_CARD_BG
    pay_insight.line.color.rgb = COLOR_CARD_BORDER
    tf_pi = pay_insight.text_frame
    tf_pi.word_wrap = True
    tf_pi.margin_left = Inches(0.3)
    tf_pi.margin_top = Inches(0.15)

    p = tf_pi.paragraphs[0]
    p.text = "💡 تحليل سلوك التقسيط لدى المستهلك البرازيلي:"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    bullets_pay = [
        "ثقافة التقسيط (Installments) هي المحرك الأول للتجارة الإلكترونية في البرازيل، حيث يقسط المشترون حتى 10 و24 قسطاً للمشتريات ذات القيمة المرتفعة.",
        "توفير خيارات تمويل مرنة بدون فوائد يساهم بشكل مباشر في رفع متوسط سلة المشتريات (AOV)."
    ]
    for b in bullets_pay:
        p = tf_pi.add_paragraph()
        p.text = "✔ " + b
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # -------------------------------------------------------------
    # SLIDE 8: REPORT 5 - CSAT & REVIEWS
    # -------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "7. Customer Satisfaction & Reviews | رضا العملاء وتأثير التأخير", "Report 5: Rating Breakdown & The Direct Cost of Delivery Delays")

    # Left: Rating Distribution
    b_rev = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4), Inches(5.8), Inches(5.6))
    b_rev.fill.solid()
    b_rev.fill.fore_color.rgb = COLOR_CARD_BG
    b_rev.line.color.rgb = COLOR_CARD_BORDER
    tf_rv = b_rev.text_frame
    tf_rv.word_wrap = True
    tf_rv.margin_left = Inches(0.3)
    tf_rv.margin_top = Inches(0.25)

    p = tf_rv.paragraphs[0]
    p.text = "⭐ توزيع تقييمات العملاء (CSAT Breakdown):"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    bullets_rev = [
        "⭐⭐⭐⭐⭐ (5 نجوم): 57.8% (57,328 تقييم) - الأغلبية راضية جداً.",
        "⭐⭐⭐⭐ (4 نجوم): 19.3% (19,142 تقييم) - تجربة شراء جيدة جداً.",
        "⭐⭐⭐ (3 نجوم): 8.2% (8,179 تقييم) - تجربة حيادية.",
        "⭐⭐ (نجمتان): 3.2% (3,151 تقييم) - ملاحظات على المنتج أو الشحن.",
        "⭐ (نجمة واحدة): 11.5% (11,424 تقييم) - عملاء غاضبون ومحبطون.",
        "📌 ملاحظة: 70%+ من مقيمي الـ 1 نجمة يكتبون تعليقات شكوى مفصلة."
    ]
    for b in bullets_rev:
        p = tf_rv.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Right: Impact of Delay
    b_del = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.7), Inches(1.4), Inches(6.0), Inches(5.6))
    b_del.fill.solid()
    b_del.fill.fore_color.rgb = COLOR_CARD_BG
    b_del.line.color.rgb = COLOR_DANGER
    b_del.line.width = Pt(1.5)
    tf_dl = b_del.text_frame
    tf_dl.word_wrap = True
    tf_dl.margin_left = Inches(0.3)
    tf_dl.margin_top = Inches(0.25)

    p = tf_dl.paragraphs[0]
    p.text = "🚨 العلاقة الصادمة بين تأخير الشحن والتقييم:"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_DANGER

    bullets_del = [
        "الطلبات المسلمة في الموعد (On-Time):",
        "   ➔ متوسط التقييم = 4.27 من 5.0 ⭐ (رضا ممتاز).",
        "الطلبات المتأخرة عن الموعد (Delayed):",
        "   ➔ متوسط التقييم = 2.27 من 5.0 ⭐ (انهيار تام للتقييم بمقدار 2.0 نجمة كاملة!).",
        "الخلاصة الإدارية:",
        "   تأخير التوصيل هو العدو الأول لسمعة المنصة ومعدل احتفاظ العملاء (Retention Rate). ضمان الالتزام بالموعد يرفع فوراً التقييم العام للمنصة."
    ]
    for b in bullets_del:
        p = tf_dl.add_paragraph()
        p.text = b
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # -------------------------------------------------------------
    # SLIDE 9: STRATEGIC RECOMMENDATIONS
    # -------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "8. Strategic Action Plan | التوصيات وخطة العمل الاستراتيجية", "Data-Driven Recommendations for Executive Decision Makers")

    rec_w = Inches(5.8)
    rec_h = Inches(2.6)

    # Rec 1: Logistics
    r1 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4), rec_w, rec_h)
    r1.fill.solid()
    r1.fill.fore_color.rgb = COLOR_CARD_BG
    r1.line.color.rgb = COLOR_CARD_BORDER
    tf_r1 = r1.text_frame
    tf_r1.word_wrap = True
    tf_r1.margin_left = Inches(0.25)
    tf_r1.margin_top = Inches(0.2)
    p = tf_r1.paragraphs[0]
    p.text = "🚚 1. تحسين الكفاءة اللوجستية والشحن"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_PRIMARY
    bullets_rec1 = [
        "بناء مراكز فرز ومستودعات إقليمية (Hubs) في ولايات الشمال والشمال الشرقي.",
        "وضع خوارزمية أكثر دقة للمواعيد التقديرية (Estimated Dates) لتفادي مفاجأة العميل بالتأخير."
    ]
    for b in bullets_rec1:
        p = tf_r1.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Rec 2: Merchandising
    r2 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.7), Inches(1.4), rec_w, rec_h)
    r2.fill.solid()
    r2.fill.fore_color.rgb = COLOR_CARD_BG
    r2.line.color.rgb = COLOR_CARD_BORDER
    tf_r2 = r2.text_frame
    tf_r2.word_wrap = True
    tf_r2.margin_left = Inches(0.25)
    tf_r2.margin_top = Inches(0.2)
    p = tf_r2.paragraphs[0]
    p.text = "🛍️ 2. تحفيز الفئات القيادية وتطوير المفروشات"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_PRIMARY
    bullets_rec2 = [
        "زيادة الحملات الإعلانية لأقسام الصحة والجمال والساعات لتحقيق هوامش ربح أعلى.",
        "وضع معايير جودة ورقابة للبائعين في قسم المفروشات والمنزل لرفع تقييماته المتراجعة."
    ]
    for b in bullets_rec2:
        p = tf_r2.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Rec 3: Payments
    r3 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(4.3), rec_w, rec_h)
    r3.fill.solid()
    r3.fill.fore_color.rgb = COLOR_CARD_BG
    r3.line.color.rgb = COLOR_CARD_BORDER
    tf_r3 = r3.text_frame
    tf_r3.word_wrap = True
    tf_r3.margin_left = Inches(0.25)
    tf_r3.margin_top = Inches(0.2)
    p = tf_r3.paragraphs[0]
    p.text = "💳 3. شراكات التقسيط الترويجي (0% Interest)"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_PRIMARY
    bullets_rec3 = [
        "عقد شراكات مع البنوك لتقديم تقسيط بدون فوائد على الفئات ذات الأسعار المرتفعة.",
        "إضافة وسائل دفع رقمية سريعة مثل PIX لتقليل الاعتماد على فواتير بوليتو البطيئة."
    ]
    for b in bullets_rec3:
        p = tf_r3.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # Rec 4: CSAT & Retention
    r4 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.7), Inches(4.3), rec_w, rec_h)
    r4.fill.solid()
    r4.fill.fore_color.rgb = COLOR_CARD_BG
    r4.line.color.rgb = COLOR_CARD_BORDER
    tf_r4 = r4.text_frame
    tf_r4.word_wrap = True
    tf_r4.margin_left = Inches(0.25)
    tf_r4.margin_top = Inches(0.2)
    p = tf_r4.paragraphs[0]
    p.text = "⭐ 4. برامج استباقية لمعالجة الشكاوى"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_PRIMARY
    bullets_rec4 = [
        "إرسال تنبيهات واعتذارات وكوبونات خصم تلقائية لأي شحنة تتأخر قبل أن يكتب العميل تقييماً سلبياً.",
        "مكافأة أفضل البائعين التزاماً بمواعيد الشحن والتغليف الجيد."
    ]
    for b in bullets_rec4:
        p = tf_r4.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_TEXT_DARK

    # -------------------------------------------------------------
    # SLIDE 10: CONCLUSION SLIDE
    # -------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    bg10 = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg10.fill.solid()
    bg10.fill.fore_color.rgb = COLOR_PRIMARY
    bg10.line.color.rgb = COLOR_PRIMARY

    tb10 = s10.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.333), Inches(3.5))
    tf10 = tb10.text_frame
    tf10.word_wrap = True

    p = tf10.paragraphs[0]
    p.text = "Thank You! | شكراً لكم"
    p.font.size = Pt(38)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.alignment = PP_ALIGN.CENTER

    p2 = tf10.add_paragraph()
    p2.text = "تم تجهيز كافة التقارير وشيت الإكسيل التفاعلي وملفات البيانات المنظفة بالكامل."
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(203, 213, 225)
    p2.alignment = PP_ALIGN.CENTER

    p3 = tf10.add_paragraph()
    p3.text = "الملفات متوفرة في مجلد المشروع وجاهزة للاستخدام والعرض المباشر."
    p3.font.size = Pt(14)
    p3.font.color.rgb = COLOR_ACCENT
    p3.alignment = PP_ALIGN.CENTER

    prs.save(output_path)
    print(f">>> PowerPoint Presentation saved successfully at {output_path}!")

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    out_file = os.path.join(base_dir, 'Olist_Ecommerce_Presentation.pptx')
    create_presentation(out_file)
