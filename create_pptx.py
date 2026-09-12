import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Universal Font Constants (Guaranteed 100% standard on all Office versions)
    FONT_TITLE = 'Arial'
    FONT_BODY = 'Calibri'
    FONT_CODE = 'Consolas'

    # Color Palette (Executive Navy Dark Theme)
    BG_COLOR = RGBColor(15, 23, 42)        # #0f172a
    CARD_BG = RGBColor(30, 41, 59)         # #1e293b
    CARD_BORDER = RGBColor(51, 65, 85)     # #334155
    WHITE = RGBColor(255, 255, 255)
    MUTED = RGBColor(148, 163, 184)        # #94a3b8
    LIGHT_TEXT = RGBColor(203, 213, 225)   # #cbd5e1
    PRIMARY = RGBColor(129, 140, 248)      # #818cf8
    CYAN = RGBColor(56, 189, 248)          # #38bdf8
    GREEN = RGBColor(52, 211, 153)         # #34d399
    AMBER = RGBColor(251, 191, 36)         # #fbbf24
    RED = RGBColor(248, 113, 113)          # #f87171
    PURPLE = RGBColor(192, 132, 252)       # #c084fc

    blank_layout = prs.slide_layouts[6]

    def create_slide(category, title, subtitle, notes=""):
        slide = prs.slides.add_slide(blank_layout)
        
        # Background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_COLOR
        bg.line.fill.background()

        # Category Badge
        cat_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.35), Inches(3.8), Inches(0.32))
        cat_box.fill.solid()
        cat_box.fill.fore_color.rgb = CARD_BG
        cat_box.line.color.rgb = CYAN
        cat_box.line.width = Pt(1)
        tf_c = cat_box.text_frame
        tf_c.word_wrap = True
        tf_c.margin_left = tf_c.margin_right = tf_c.margin_top = tf_c.margin_bottom = 0
        p_c = tf_c.paragraphs[0]
        p_c.alignment = PP_ALIGN.CENTER
        p_c.text = category.upper()
        p_c.font.name = FONT_BODY
        p_c.font.size = Pt(9)
        p_c.font.bold = True
        p_c.font.color.rgb = CYAN

        # Title
        tb_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.55))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(22)
        p_t.font.bold = True
        p_t.font.color.rgb = WHITE

        # Subtitle
        tb_s = slide.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(11.7), Inches(0.35))
        tf_s = tb_s.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = tf_s.margin_right = tf_s.margin_top = tf_s.margin_bottom = 0
        p_s = tf_s.paragraphs[0]
        p_s.text = subtitle
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(11)
        p_s.font.color.rgb = MUTED

        if notes:
            slide.notes_slide.notes_text_frame.text = notes

        return slide

    def add_card(slide, x, y, w, h, border_color=None, bg_color=None):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color or CARD_BG
        card.line.color.rgb = border_color or CARD_BORDER
        card.line.width = Pt(1.2)
        return card

    def add_quiz_slide(slide_num, category, title, subtitle, badge, prompt, options, correct_idx, explanation, notes):
        slide = create_slide(category, title, subtitle, notes)
        
        # Prompt Box
        p_box = add_card(slide, 0.8, 1.8, 11.73, 1.25, border_color=PRIMARY)
        tf_p = p_box.text_frame
        tf_p.word_wrap = True
        tf_p.margin_left = Inches(0.4)
        tf_p.margin_top = Inches(0.2)
        p_badge = tf_p.paragraphs[0]
        p_badge.text = f"[{badge}]"
        p_badge.font.name = FONT_TITLE
        p_badge.font.bold = True
        p_badge.font.size = Pt(11)
        p_badge.font.color.rgb = AMBER

        p_txt = tf_p.add_paragraph()
        p_txt.text = prompt
        p_txt.font.name = FONT_BODY
        p_txt.font.bold = True
        p_txt.font.size = Pt(14)
        p_txt.font.color.rgb = WHITE
        p_txt.space_before = Pt(6)

        # 4 Options Grid (2x2)
        opt_coords = [
            (0.8, 3.2), (6.8, 3.2),
            (0.8, 4.3), (6.8, 4.3)
        ]
        for i, (ox, oy) in enumerate(opt_coords):
            is_correct = (i == correct_idx)
            btn_bg = RGBColor(6, 78, 59) if is_correct else CARD_BG
            btn_border = GREEN if is_correct else CARD_BORDER
            b_card = add_card(slide, ox, oy, 5.7, 0.9, border_color=btn_border, bg_color=btn_bg)
            tf_b = b_card.text_frame
            tf_b.word_wrap = True
            tf_b.margin_left = Inches(0.3)
            tf_b.margin_top = Inches(0.2)
            p_opt = tf_b.paragraphs[0]
            opt_text = options[i]
            if is_correct:
                opt_text += "  [CORRECT]"
            p_opt.text = opt_text
            p_opt.font.name = FONT_BODY
            p_opt.font.size = Pt(13)
            p_opt.font.bold = is_correct
            p_opt.font.color.rgb = GREEN if is_correct else LIGHT_TEXT

        # Explanation Box
        exp_card = add_card(slide, 0.8, 5.45, 11.73, 1.35, border_color=GREEN)
        tf_e = exp_card.text_frame
        tf_e.word_wrap = True
        tf_e.margin_left = Inches(0.4)
        tf_e.margin_top = Inches(0.18)
        p_ans = tf_e.paragraphs[0]
        p_ans.text = f"[CORRECT ANSWER: {options[correct_idx]}]"
        p_ans.font.name = FONT_TITLE
        p_ans.font.bold = True
        p_ans.font.size = Pt(11)
        p_ans.font.color.rgb = GREEN

        p_exp = tf_e.add_paragraph()
        p_exp.text = explanation
        p_exp.font.name = FONT_BODY
        p_exp.font.size = Pt(11.5)
        p_exp.font.color.rgb = LIGHT_TEXT
        p_exp.space_before = Pt(4)

        return slide

    # ========================================================
    # SLIDE 1: Cover
    # ========================================================
    s1 = create_slide(
        "English Grammar Masterclass",
        "PRONOUNS IN ENGLISH",
        "How Small Words Create Fluent, Natural English",
        "KỊCH BẢN NÓI: Xin kính chào quý vị và các bạn! Hôm nay chúng ta sẽ cùng chinh phục chủ đề: PRONOUNS IN ENGLISH - Đại từ trong tiếng Anh."
    )
    c1 = add_card(s1, 1.5, 2.0, 10.33, 4.5, border_color=PRIMARY)
    tf1 = c1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = Inches(0.6)
    tf1.margin_top = Inches(0.8)
    p = tf1.paragraphs[0]
    p.text = "Mastering the 5 Essential Pronoun Families"
    p.font.name = FONT_TITLE
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    
    p = tf1.add_paragraph()
    p.text = "Subject vs Object  |  Possessive Adjectives & Pronouns  |  Reflexive & Other Families"
    p.font.name = FONT_BODY
    p.font.size = Pt(14)
    p.font.color.rgb = WHITE
    p.space_before = Pt(16)

    p = tf1.add_paragraph()
    p.text = "5 Core Modules • Strategic Syntax Rules • 10 Interactive Classroom Polls"
    p.font.name = FONT_BODY
    p.font.size = Pt(13)
    p.font.color.rgb = CYAN
    p.space_before = Pt(24)

    # ========================================================
    # SLIDE 2: The Problem with Repetition
    # ========================================================
    s2 = create_slide(
        "Icebreaker & Motivation",
        "The Problem with Repetition",
        "Why natural English communication strictly depends on pronouns",
        "KỊCH BẢN NÓI: Hãy tưởng tượng một câu nói lặp lại từ 'Tom' 5 lần: 'Tom is a student. Tom likes football...' Nghe rất nặng nề và thiếu tự nhiên!"
    )
    c2_bad = add_card(s2, 0.8, 1.9, 5.7, 4.8, border_color=RED)
    tf2_b = c2_bad.text_frame
    tf2_b.word_wrap = True
    tf2_b.margin_left = Inches(0.4)
    tf2_b.margin_top = Inches(0.3)
    p = tf2_b.paragraphs[0]
    p.text = "[REPETITIVE & ROBOTIC]"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = RED
    p = tf2_b.add_paragraph()
    p.text = '“If I say:\nTom is a student.\nTom likes football.\nTom plays football every day.\nTom lives near my house.\n\nDoes that sound natural?”'
    p.font.name = FONT_BODY
    p.font.size = Pt(13)
    p.font.color.rgb = WHITE
    p.space_before = Pt(12)

    c2_good = add_card(s2, 6.8, 1.9, 5.7, 4.8, border_color=GREEN)
    tf2_g = c2_good.text_frame
    tf2_g.word_wrap = True
    tf2_g.margin_left = Inches(0.4)
    tf2_g.margin_top = Inches(0.3)
    p = tf2_g.paragraphs[0]
    p.text = "[NATURAL & FLUENT DELIVERY]"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = GREEN
    p = tf2_g.add_paragraph()
    p.text = '“Instead, we say:\nTom is a student.\nHe likes football.\nHe plays football every day.\n\n-> The pronoun HE replaces the noun Tom to create rhythm and fluency.”'
    p.font.name = FONT_BODY
    p.font.size = Pt(13)
    p.font.color.rgb = WHITE
    p.space_before = Pt(12)

    # ========================================================
    # SLIDE 3: Definition & Strategic Value
    # ========================================================
    s3 = create_slide(
        "Foundation • Core Concept",
        "What is a Pronoun & Why Do We Need Them?",
        "Replacing nouns to eliminate redundancy, ensure economy of language, and drive fluency",
        "KỊCH BẢN NÓI: Đại từ là từ dùng để thay thế cho danh từ hoặc cụm danh từ đã được nhắc đến, giúp câu văn gọn gàng, súc tích."
    )
    c3_l = add_card(s3, 0.8, 1.9, 5.7, 4.8, border_color=CYAN)
    tf3_l = c3_l.text_frame
    tf3_l.word_wrap = True
    tf3_l.margin_left = Inches(0.4)
    tf3_l.margin_top = Inches(0.3)
    p = tf3_l.paragraphs[0]
    p.text = "CORE DEFINITION"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = CYAN
    p = tf3_l.add_paragraph()
    p.text = "A pronoun is a word used in place of a noun or noun phrase to refer to people, objects, or concepts.\n\nKey Strategic Functions:\n• Prevents tiresome repetition\n• Connects ideas across sentences\n• Distinguishes subjects from objects\n• Clarifies ownership instantly"
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(12)

    c3_r = add_card(s3, 6.8, 1.9, 5.7, 4.8, border_color=PRIMARY)
    tf3_r = c3_r.text_frame
    tf3_r.word_wrap = True
    tf3_r.margin_left = Inches(0.4)
    tf3_r.margin_top = Inches(0.3)
    p = tf3_r.paragraphs[0]
    p.text = "RAPID MAPPING EXAMPLES"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = PRIMARY
    p = tf3_r.add_paragraph()
    p.text = "• Tom -> he (Singular male)\n• Anna -> she (Singular female)\n• The book -> it (Object/concept)\n• Tom and Anna -> they (Plural group)\n• My colleagues and I -> we (Inclusive plural)\n\nGolden Takeaway:\nMastering pronouns is mastering the backbone of English sentence structure!"
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(12)

    # ========================================================
    # SLIDE 4: Roadmap (5 Core Modules)
    # ========================================================
    s4 = create_slide(
        "Structure • Roadmap",
        "Presentation Roadmap: The 5 Core Modules",
        "A structured curriculum designed to build absolute syntactic confidence",
        "KỊCH BẢN NÓI: Lộ trình bài giảng hôm nay gồm 5 phần chính: Chủ ngữ, Tân ngữ, Tính từ sở hữu, Đại từ sở hữu, và Đại từ phản thân."
    )
    mod_info = [
        ("Part 01", "Subject Pronouns", "I, you, he, she, it, we, they", "Initiator before verb (S + V)", PRIMARY),
        ("Part 02", "Object Pronouns", "me, you, him, her, it, us, them", "Receiver after verb & prep", CYAN),
        ("Part 03", "Possessive Adjectives", "my, your, his, her, its, our, their", "Must be followed by a noun", AMBER),
        ("Part 04", "Possessive Pronouns", "mine, yours, his, hers, ours, theirs", "Stands alone without noun", GREEN),
        ("Part 05", "Reflexive & Other", "myself, yourself... & 4 Families", "Reflexive focus + Overview", PURPLE)
    ]
    for idx, (p_num, p_title, p_words, p_role, p_color) in enumerate(mod_info):
        mx = 0.8 + (idx * 2.4)
        m_card = add_card(s4, mx, 2.0, 2.2, 4.6, border_color=p_color)
        tf_m = m_card.text_frame
        tf_m.word_wrap = True
        tf_m.margin_left = Inches(0.2)
        tf_m.margin_top = Inches(0.3)
        p = tf_m.paragraphs[0]
        p.text = p_num
        p.font.name = FONT_TITLE
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = p_color
        p = tf_m.add_paragraph()
        p.text = p_title
        p.font.name = FONT_TITLE
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = WHITE
        p.space_before = Pt(8)
        p = tf_m.add_paragraph()
        p.text = p_words
        p.font.name = FONT_BODY
        p.font.size = Pt(10.5)
        p.font.color.rgb = p_color
        p.space_before = Pt(8)
        p = tf_m.add_paragraph()
        p.text = p_role
        p.font.name = FONT_BODY
        p.font.size = Pt(10)
        p.font.color.rgb = MUTED
        p.space_before = Pt(12)

    # ========================================================
    # SLIDE 5: Part 01 - Subject Pronouns
    # ========================================================
    s5 = create_slide(
        "Part 01 • Subject Pronouns",
        "Subject Pronouns",
        "The primary initiators of action: I, you, he, she, it, we, they",
        "KỊCH BẢN NÓI: Đại từ chủ ngữ luôn đứng TRƯỚC động từ chính để đóng vai trò chủ thể thực hiện hành động."
    )
    c5_l = add_card(s5, 0.8, 1.9, 5.7, 4.8, border_color=PRIMARY)
    tf5_l = c5_l.text_frame
    tf5_l.word_wrap = True
    tf5_l.margin_left = Inches(0.4)
    tf5_l.margin_top = Inches(0.3)
    p = tf5_l.paragraphs[0]
    p.text = "CORE SYNTAX RULE"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = PRIMARY
    p = tf5_l.add_paragraph()
    p.text = "[ SUBJECT PRONOUN ] + [ VERB ]\n\n• Always placed BEFORE the main verb\n• Initiates and controls the grammatical action\n• Conjugates the verb (He plays vs They play)"
    p.font.name = FONT_BODY
    p.font.size = Pt(13)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    c5_r = add_card(s5, 6.8, 1.9, 5.7, 4.8, border_color=CYAN)
    tf5_r = c5_r.text_frame
    tf5_r.word_wrap = True
    tf5_r.margin_left = Inches(0.4)
    tf5_r.margin_top = Inches(0.3)
    p = tf5_r.paragraphs[0]
    p.text = "EVERYDAY EXAMPLES"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = CYAN
    p = tf5_r.add_paragraph()
    p.text = "• I am a software engineer.\n• She is my colleague.\n• He plays football every weekend.\n• They submitted the report on time.\n\nClassic Anchor: In 'I love you', I is the Subject Pronoun initiating love!"
    p.font.name = FONT_BODY
    p.font.size = Pt(13)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    # ========================================================
    # QUIZ 1 & 2 (Part 01)
    # ========================================================
    add_quiz_slide(
        6, "Part 01 Practice • Question 01", "Subject Identification Assessment",
        "Selecting the correct initiator before the main predicate",
        "CHECK 01",
        "“Tom is a talented striker. _____ plays football every weekend with his local club.”",
        ["A. He", "B. Him", "C. His", "D. Himself"],
        0,
        "Positioned before verb 'plays' as subject ➔ strictly requires Subject Pronoun 'He'. Him (Object), His (Possessive), Himself (Reflexive) cannot function as subject.",
        "KỊCH BẢN NÓI: Câu 01 kiểm tra vị trí chủ ngữ trước động từ 'plays'. Đáp án đúng là A: He."
    )

    add_quiz_slide(
        7, "Part 01 Practice • Question 02", "Compound Subject Placement",
        "Navigating compound subject conjunctions without falling for object traps",
        "CHECK 02",
        "“After the conference ended, David and _____ submitted the project report to the director.”",
        ["A. me", "B. I", "C. myself", "D. mine"],
        1,
        "'David and I' act as compound subjects for verb 'submitted'. Verification trick: omit 'David and' ➔ 'I submitted' (correct), never 'me submitted'.",
        "KỊCH BẢN NÓI: Câu 02 kiểm tra chủ ngữ ghép. Mẹo loại trừ: bỏ 'David and', ta còn lại 'I submitted', do đó chọn B: I."
    )

    # ========================================================
    # SLIDE 8: Part 02 - Object Pronouns
    # ========================================================
    s8 = create_slide(
        "Part 02 • Object Pronouns",
        "Object Pronouns",
        "Recipients of action: me, you, him, her, it, us, them",
        "KỊCH BẢN NÓI: Đại từ tân ngữ luôn đứng SAU động từ hoặc SAU giới từ để tiếp nhận tác động."
    )
    c8_l = add_card(s8, 0.8, 1.9, 5.7, 4.8, border_color=CYAN)
    tf8_l = c8_l.text_frame
    tf8_l.word_wrap = True
    tf8_l.margin_left = Inches(0.4)
    tf8_l.margin_top = Inches(0.3)
    p = tf8_l.paragraphs[0]
    p.text = "THE 2 GOVERNING POSITIONS"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = CYAN
    p = tf8_l.add_paragraph()
    p.text = "1. After Transitive Verbs (V + O):\n   • Peter invited HER to lunch.\n   • We called THEM yesterday.\n\n2. After Prepositions (Prep + O):\n   • Listen to ME.\n   • Between you and ME (never 'I').\n   • Speak with HIM."
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    c8_r = add_card(s8, 6.8, 1.9, 5.7, 4.8, border_color=PRIMARY)
    tf8_r = c8_r.text_frame
    tf8_r.word_wrap = True
    tf8_r.margin_left = Inches(0.4)
    tf8_r.margin_top = Inches(0.3)
    p = tf8_r.paragraphs[0]
    p.text = "EXECUTIVE TAKEAWAYS"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = PRIMARY
    p = tf8_r.add_paragraph()
    p.text = "Classic Contrast:\n• 'I love you' (You = Object)\n• 'You love me' (Me = Object)\n\nTOEIC Trap Alert:\nNever put Subject Pronouns after prepositions:\n✗ between you and I\n✓ between you and me"
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    # ========================================================
    # QUIZ 3 & 4 (Part 02)
    # ========================================================
    add_quiz_slide(
        9, "Part 02 Practice • Question 03", "Object Replacement Analysis",
        "Identifying direct objects receiving verb action",
        "CHECK 03",
        "“Anna is very friendly. Peter likes _____ very much.”",
        ["A. she", "B. Anna likes him.", "C. her", "D. hers"],
        2,
        "'Anna' functions as direct object after transitive verb 'likes' ➔ requires Object Pronoun 'her'. Options 'she' (Subject) and 'hers' (Possessive Pronoun) are invalid.",
        "KỊCH BẢN NÓI: Câu 03 đứng sau ngoại động từ 'likes' làm tân ngữ, đáp án đúng là C: her."
    )

    add_quiz_slide(
        10, "Part 02 Practice • Question 04", "Prepositional Object Assessment",
        "Applying the universal rule: Prepositions mandate object pronouns",
        "CHECK 04",
        "“This confidential agreement must strictly remain between the client and _____.”",
        ["A. I", "B. me", "C. my", "D. mine"],
        1,
        "'Between' is a preposition. Prepositions strictly require Object Pronouns ➔ 'between the client and me'. Subject Pronoun 'I' cannot follow a preposition.",
        "KỊCH BẢN NÓI: Câu 04 sau giới từ 'between' bắt buộc dùng đại từ tân ngữ 'me'. Đáp án đúng là B: me."
    )

    # ========================================================
    # SLIDE 11: Part 03 - Possessive Adjectives (Theory 1)
    # ========================================================
    s11 = create_slide(
        "Part 03 • Possessive Adjectives",
        "Possessive Adjectives",
        "Essential noun modifiers establishing belonging: Possessive Adjective + Noun",
        "KỊCH BẢN NÓI: Tính từ sở hữu KHÔNG BAO GIỜ đứng một mình mà luôn đi kèm danh từ phía sau."
    )
    c11_l = add_card(s11, 0.8, 1.9, 5.7, 4.8, border_color=AMBER)
    tf11_l = c11_l.text_frame
    tf11_l.word_wrap = True
    tf11_l.margin_left = Inches(0.4)
    tf11_l.margin_top = Inches(0.3)
    p = tf11_l.paragraphs[0]
    p.text = "MANDATORY NOUN RULE"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = AMBER
    p = tf11_l.add_paragraph()
    p.text = "[ Possessive Adjective ] + [ NOUN / Noun Phrase ]\n\nExamples requested by user:\n• This is MY book.\n• HER laptop is on the table.\n\n* Never stands alone; always precedes and modifies an accompanying noun."
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    c11_r = add_card(s11, 6.8, 1.9, 5.7, 4.8, border_color=RED)
    tf11_r = c11_r.text_frame
    tf11_r.word_wrap = True
    tf11_r.margin_left = Inches(0.4)
    tf11_r.margin_top = Inches(0.3)
    p = tf11_r.paragraphs[0]
    p.text = "THE 'ITS' VS 'IT'S' TRAP"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = RED
    p = tf11_r.add_paragraph()
    p.text = "[RULE 1] ITS (no apostrophe) = Possessive Adjective\nDenotes ownership for objects, entities, or animals:\n> “The company updated its security protocol.”\n\n[RULE 2] IT'S (with apostrophe) = Contraction of 'it is' / 'it has'\nA complete clause (Subject + Verb), NOT an adjective:\n> “It's important to verify the numbers.”"
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    # ========================================================
    # SLIDE 12: Part 03 - The 7 Possessive Adjectives (Theory 2)
    # ========================================================
    s12 = create_slide(
        "Part 03 • Possessive Adjectives",
        "The 7 Possessive Adjectives",
        "Corresponding to each personal pronoun, there are 7 fundamental possessive adjectives",
        "KỊCH BẢN NÓI: Tương ứng với 7 đại từ nhân xưng, ta có 7 tính từ sở hữu cơ bản: my, your, his, her, its, our, their."
    )
    c12_l = add_card(s12, 0.8, 1.9, 5.7, 4.8, border_color=AMBER)
    tf12_l = c12_l.text_frame
    tf12_l.word_wrap = True
    tf12_l.margin_left = Inches(0.4)
    tf12_l.margin_top = Inches(0.3)
    p = tf12_l.paragraphs[0]
    p.text = "SINGULAR FORMS (INDIVIDUAL OWNERSHIP)"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(12.5)
    p.font.color.rgb = AMBER
    p = tf12_l.add_paragraph()
    p.text = "• I    -> MY    (e.g., my phone)\n• YOU  -> YOUR  (e.g., your idea)\n• HE   -> HIS   (e.g., his car)\n• SHE  -> HER   (e.g., her bag)\n• IT   -> ITS   (e.g., its tail)"
    p.font.name = FONT_CODE
    p.font.size = Pt(13)
    p.font.color.rgb = WHITE
    p.space_before = Pt(12)

    c12_r = add_card(s12, 6.8, 1.9, 5.7, 4.8, border_color=GREEN)
    tf12_r = c12_r.text_frame
    tf12_r.word_wrap = True
    tf12_r.margin_left = Inches(0.4)
    tf12_r.margin_top = Inches(0.3)
    p = tf12_r.paragraphs[0]
    p.text = "PLURAL FORMS & CORE SYNTAX RULE"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(12.5)
    p.font.color.rgb = GREEN
    p = tf12_r.add_paragraph()
    p.text = "• WE   -> OUR   (e.g., our team)\n• THEY -> THEIR (e.g., their house)\n\n[CORE RULE]:\n[ Possessive Adjective ] + [ NOUN / Noun Phrase ]\nAll 7 possessive adjectives strictly require an accompanying noun. They never stand alone!"
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(12)

    # ========================================================
    # QUIZ 5 & 6 (Part 03)
    # ========================================================
    add_quiz_slide(
        13, "Part 03 Practice • Question 05", "Noun Collocation Assessment",
        "Observing structural dependencies before the noun",
        "CHECK 05",
        "“John is looking for _____ keys. He cannot find them anywhere.”",
        ["A. he", "B. him", "C. his", "D. himself"],
        2,
        "Immediately following the blank is the plural noun keys. Apply rule: [Possessive Adjective + Noun], this position strictly requires a possessive adjective to modify keys. Referring to singular male John ➔ select 'his'.",
        "KỊCH BẢN NÓI: Câu 05 ngay sau khoảng trống là danh từ 'keys'. Cần tính từ sở hữu chỉ John, đáp án đúng là C: his."
    )

    add_quiz_slide(
        14, "Part 03 Practice • Question 06", "Subject-Referent Possessive Assessment",
        "Selecting the correct possessive determiner before plural nouns",
        "CHECK 06",
        "“We invited all of _____ friends to the end-of-year party.”",
        ["A. our", "B. us", "C. ours", "D. we"],
        0,
        "Following the blank is the plural noun friends. According to formula: [Possessive Adjective + Noun], a possessive adjective is required. For subject We ➔ select 'our' (our friends). Note: 'ours' stands alone without a noun.",
        "KỊCH BẢN NÓI: Câu 06 trước danh từ 'friends', đại từ chủ ngữ là We, nên tính từ sở hữu tương ứng là A: our."
    )

    # ========================================================
    # SLIDE 15: Part 04 - Possessive Pronouns
    # ========================================================
    s15 = create_slide(
        "Part 04 • Possessive Pronouns",
        "Possessive Pronouns",
        "Addressing the fundamental question: “Whose is this?”",
        "KỊCH BẢN NÓI: Đại từ sở hữu đứng ĐỘC LẬP một mình, không có danh từ đi kèm phía sau."
    )
    c15_l = add_card(s15, 0.8, 1.9, 5.7, 4.8, border_color=PURPLE)
    tf15_l = c15_l.text_frame
    tf15_l.word_wrap = True
    tf15_l.margin_left = Inches(0.4)
    tf15_l.margin_top = Inches(0.3)
    p = tf15_l.paragraphs[0]
    p.text = "THE 7 POSSESSIVE PRONOUNS"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = PURPLE
    p = tf15_l.add_paragraph()
    p.text = "• I    -> MINE\n• YOU  -> YOURS\n• HE   -> HIS\n• SHE  -> HERS\n• IT   -> ITS\n• WE   -> OURS\n• THEY -> THEIRS"
    p.font.name = FONT_CODE
    p.font.size = Pt(13)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    c15_r = add_card(s15, 6.8, 1.9, 5.7, 4.8, border_color=GREEN)
    tf15_r = c15_r.text_frame
    tf15_r.word_wrap = True
    tf15_r.margin_left = Inches(0.4)
    tf15_r.margin_top = Inches(0.3)
    p = tf15_r.paragraphs[0]
    p.text = "STRATEGIC ROLE & USAGE"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = GREEN
    p = tf15_r.add_paragraph()
    p.text = "Used to declare ownership definitively without repeating the noun:\n“Whose book is this?”\n\nDemonstration:\n“This book is MINE.”\n= This book belongs to me.\n\n* Key point: MINE stands completely alone at clause end."
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    # ========================================================
    # SLIDE 16: Synthesis: Possessive Adjective vs Possessive Pronoun
    # ========================================================
    s16 = create_slide(
        "Part 03 & 04 • Comparative Synthesis",
        "Possessive Adjective vs. Possessive Pronoun",
        "The algebraic formula of English grammar: Possessive Pronoun = Possessive Adjective + Noun",
        "KỊCH BẢN NÓI: Công thức đại số cốt lõi: Đại từ sở hữu = Tính từ sở hữu + Danh từ. Ví dụ: mine = my books."
    )
    c16_l = add_card(s16, 0.8, 1.9, 5.7, 4.8, border_color=AMBER)
    tf16_l = c16_l.text_frame
    tf16_l.word_wrap = True
    tf16_l.margin_left = Inches(0.4)
    tf16_l.margin_top = Inches(0.3)
    p = tf16_l.paragraphs[0]
    p.text = "CATEGORY A: POSSESSIVE ADJECTIVES"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = AMBER
    p = tf16_l.add_paragraph()
    p.text = "my, your, his, her, its, our, their\n\nRule: Possessive Adjective + Noun\nExample:\n• My house is beautiful.\n(Never stands alone)"
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    c16_r = add_card(s16, 6.8, 1.9, 5.7, 4.8, border_color=CYAN)
    tf16_r = c16_r.text_frame
    tf16_r.word_wrap = True
    tf16_r.margin_left = Inches(0.4)
    tf16_r.margin_top = Inches(0.3)
    p = tf16_r.paragraphs[0]
    p.text = "CATEGORY B: POSSESSIVE PRONOUNS"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = CYAN
    p = tf16_r.add_paragraph()
    p.text = "mine, yours, his, hers, ours, theirs\n\nThe Algebraic Equation:\nPossessive Pronoun = Possessive Adj + Noun\nExample:\n• “Your books are heavy, MINE are heavy too.”\n  (mine = my books, avoids repeating 'books')"
    p.font.name = FONT_BODY
    p.font.size = Pt(12.5)
    p.font.color.rgb = WHITE
    p.space_before = Pt(10)

    # ========================================================
    # QUIZ 7 & 8 (Part 04)
    # ========================================================
    add_quiz_slide(
        17, "Part 04 Practice • Question 07", "Standalone Position Assessment",
        "Evaluating standalone ownership markers without succeeding nouns",
        "CHECK 07",
        "“My car is in the repair shop, but _____ is parked outside.”",
        ["A. your", "B. yours", "C. you", "D. yourself"],
        1,
        "Follows clause transition without an accompanying noun ➔ strictly requires Possessive Pronoun 'yours' (= your car). 'Your' is an adjective and requires a noun.",
        "KỊCH BẢN NÓI: Câu 07 đứng một mình làm chủ ngữ vế sau (yours = your car), đáp án đúng là B: yours."
    )

    add_quiz_slide(
        18, "Part 04 Practice • Question 08", "Predicate Possessive Construction",
        "Distinguishing belonging without accompanying nouns",
        "CHECK 08",
        "“These project documents belong to the marketing team; in fact, they are _____.”",
        ["A. their", "B. them", "C. theirs", "D. themselves"],
        2,
        "Positioned after verb 'are' at sentence end to denote independent ownership (they are theirs = they are their documents) ➔ requires Possessive Pronoun 'theirs'.",
        "KỊCH BẢN NÓI: Câu 08 đứng sau động từ 'are' ở cuối câu chỉ quyền sở hữu, đáp án đúng là C: theirs."
    )

    # ========================================================
    # SLIDE 19: Part 05 - Reflexive Pronouns (Core Focus)
    # ========================================================
    s19 = create_slide(
        "Part 05 • Reflexive Pronouns",
        "Reflexive Pronouns: 3 Strategic Functions",
        "When the action returns to the subject: myself, yourself, himself, herself, itself, ourselves, themselves",
        "KỊCH BẢN NÓI: Đại từ phản thân có 3 chức năng chính: Phản thân (Subject = Object), Nhấn mạnh (Emphatic), và Thành ngữ 'by + reflexive' = một mình."
    )
    f_info = [
        ("1. REFLECTION (Subject = Object)", "Action reflects back onto the actor:\n• “I love myself.”\n• “He cut himself while cooking.”\n• Subject & Object are the same person.", PRIMARY),
        ("2. EMPHATIC INTENSITY", "Emphasizes personal agency:\n• “She carried these books herself.”\n• “The CEO himself signed the contract.”\n• Can be removed without breaking grammar.", CYAN),
        ("3. BY + REFLEXIVE = ALONE", "Classic idiomatic syntax:\n• by myself = on my own / alone\n• “I completed the project by myself.”\n• Signifies independence and self-reliance.", GREEN)
    ]
    for idx, (f_title, f_body, f_color) in enumerate(f_info):
        fx = 0.8 + (idx * 4.0)
        fc = add_card(s19, fx, 1.9, 3.73, 4.8, border_color=f_color)
        tff = fc.text_frame
        tff.word_wrap = True
        tff.margin_left = Inches(0.3)
        tff.margin_top = Inches(0.3)
        p = tff.paragraphs[0]
        p.text = f_title
        p.font.name = FONT_TITLE
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = f_color
        p = tff.add_paragraph()
        p.text = f_body
        p.font.name = FONT_BODY
        p.font.size = Pt(12)
        p.font.color.rgb = WHITE
        p.space_before = Pt(12)

    # ========================================================
    # SLIDE 20: Part 05 - Overview of Other Pronoun Families
    # ========================================================
    s20 = create_slide(
        "Part 05 • Advanced Overview",
        "Overview of Other Pronoun Families",
        "A rapid strategic tour of Demonstrative, Indefinite, Interrogative, and Relative pronouns",
        "KỊCH BẢN NÓI: Tổng quan 4 nhóm đại từ còn lại: Chỉ định (this/that), Bất định (everyone chia số ít), Nghi vấn (who/whose), Quan hệ (who/which/that)."
    )
    other_fams = [
        ("1. DEMONSTRATIVE PRONOUNS", "Near: This / These • Far: That / Those\nNote: This/That also pair with uncountable nouns (This advice, That water).", CYAN),
        ("2. INDEFINITE PRONOUNS", "Someone, anyone, everyone, no one, nothing...\n[!] GOLDEN RULE: Always takes a singular verb: “Everyone is ready.”", GREEN),
        ("3. INTERROGATIVE PRONOUNS", "Who, whom, whose, which, what...\nIntroduces direct or indirect inquiries: “Whose jacket is this?”", PURPLE),
        ("4. RELATIVE PRONOUNS", "Who, whom, which, that, whose...\nConnects relative clauses to antecedent nouns: “The candidate who won the election...”", AMBER)
    ]
    coords_2x2 = [(0.8, 1.9), (6.8, 1.9), (0.8, 4.3), (6.8, 4.3)]
    for idx, (ox, oy) in enumerate(coords_2x2):
        otitle, obody, ocolor = other_fams[idx]
        ocard = add_card(s20, ox, oy, 5.7, 2.2, border_color=ocolor)
        tfo = ocard.text_frame
        tfo.word_wrap = True
        tfo.margin_left = Inches(0.3)
        tfo.margin_top = Inches(0.2)
        p = tfo.paragraphs[0]
        p.text = otitle
        p.font.name = FONT_TITLE
        p.font.bold = True
        p.font.size = Pt(11.5)
        p.font.color.rgb = ocolor
        p = tfo.add_paragraph()
        p.text = obody
        p.font.name = FONT_BODY
        p.font.size = Pt(11.5)
        p.font.color.rgb = WHITE
        p.space_before = Pt(6)

    # ========================================================
    # QUIZ 9 & 10 (Part 05)
    # ========================================================
    add_quiz_slide(
        21, "Part 05 Practice • Question 09", "Self-Action Analysis",
        "Handling actions where the subject and object are the same person",
        "CHECK 09",
        "“John was preparing dinner when he accidentally cut _____ with the knife.”",
        ["A. him", "B. his", "C. himself", "D. he"],
        2,
        "John performs the action of cutting and John himself receives the injury (Subject = Object). When the subject and object are identical ➔ requires Reflexive Pronoun 'himself'.",
        "KỊCH BẢN NÓI: Câu 09 chủ ngữ và tân ngữ là cùng một người (John tự làm đứt tay mình), đáp án đúng là C: himself."
    )

    add_quiz_slide(
        22, "Part 05 Practice • Question 10", "Solo Execution Assessment (By + Reflexive)",
        "Recognizing the classic idiomatic construction 'by + reflexive'",
        "CHECK 10",
        "“Did someone assist him, or did he complete the assignment by _____?”",
        ["A. him", "B. himself", "C. his", "D. he"],
        1,
        "Classic idiomatic syntax: 'by + Reflexive Pronoun' = without assistance (alone / on one's own). Since subject is 'He' ➔ select 'by himself'.",
        "KỊCH BẢN NÓI: Câu 10 cấu trúc 'by + reflexive' mang nghĩa tự làm một mình, đáp án đúng là B: himself."
    )

    # ========================================================
    # SLIDE 23: Strategic Decision Framework
    # ========================================================
    s23 = create_slide(
        "Synthesis • Decision Framework",
        "The 3 Golden Questions for Any Sentence",
        "A mental flowchart to eliminate guesswork and achieve 100% syntactic accuracy",
        "KỊCH BẢN NÓI: Mỗi khi đứng trước một khoảng trống đại từ, chỉ cần tự hỏi 3 câu hỏi vàng: 1. Trước hay sau động từ? 2. Có danh từ đi sau không? 3. Chủ ngữ có trùng tân ngữ không?"
    )
    q_framework = [
        ("QUESTION 1: POSITION?", "Is it BEFORE or AFTER the verb / preposition?\n• Initiator before Verb -> Subject Pronoun (He plays)\n• Target after Verb/Prep -> Object Pronoun (I like him / between you and me)", PRIMARY),
        ("QUESTION 2: NOUN FOLLOWING?", "Is there a NOUN immediately after the blank?\n• YES -> Possessive Adjective (my book, her car)\n• NO (Stands alone) -> Possessive Pronoun (The car is mine / hers)", CYAN),
        ("QUESTION 3: SAME PERSON / SOLO?", "Does action reflect back on the initiator or done solo?\n• YES -> Reflexive Pronoun (He cut himself / did it by himself)\n• NO -> Normal Object Pronoun", GREEN)
    ]
    for idx, (q_title, q_body, q_color) in enumerate(q_framework):
        qx = 0.8 + (idx * 4.0)
        qc = add_card(s23, qx, 1.9, 3.73, 4.8, border_color=q_color)
        tfq = qc.text_frame
        tfq.word_wrap = True
        tfq.margin_left = Inches(0.3)
        tfq.margin_top = Inches(0.3)
        p = tfq.paragraphs[0]
        p.text = q_title
        p.font.name = FONT_TITLE
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = q_color
        p = tfq.add_paragraph()
        p.text = q_body
        p.font.name = FONT_BODY
        p.font.size = Pt(12)
        p.font.color.rgb = WHITE
        p.space_before = Pt(12)

    # ========================================================
    # SLIDE 24: Master Taxonomy Matrix
    # ========================================================
    s24 = create_slide(
        "Master Reference • Taxonomy",
        "All-in-One Pronouns Master Matrix",
        "The comprehensive reference taxonomy for the 7 personal pronoun entities",
        "KỊCH BẢN NÓI: Đây là bảng ma trận tổng hợp toàn diện cả 7 ngôi qua 5 cột chức năng ngữ pháp."
    )
    # Add Table
    table_shape = s24.shapes.add_table(8, 6, Inches(0.8), Inches(1.9), Inches(11.73), Inches(4.8))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(2.4)
    tbl.columns[1].width = Inches(1.8)
    tbl.columns[2].width = Inches(1.8)
    tbl.columns[3].width = Inches(1.9)
    tbl.columns[4].width = Inches(1.9)
    tbl.columns[5].width = Inches(1.93)

    headers = ["ENTITY", "SUBJECT", "OBJECT", "POSS. ADJ", "POSS. PRON", "REFLEXIVE"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = CARD_BG
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = FONT_TITLE
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = CYAN
        p.alignment = PP_ALIGN.CENTER

    matrix_data = [
        ("1st Sing. (I)", "I", "me", "my + N", "mine", "myself"),
        ("2nd Sing./Plur. (You)", "you", "you", "your + N", "yours", "yourself / -selves"),
        ("3rd Male (He)", "he", "him", "his + N", "his", "himself"),
        ("3rd Female (She)", "she", "her", "her + N", "hers", "herself"),
        ("3rd Neutral (It)", "it", "it", "its + N", "its", "itself"),
        ("1st Plural (We)", "we", "us", "our + N", "ours", "ourselves"),
        ("3rd Plural (They)", "they", "them", "their + N", "theirs", "themselves")
    ]
    for i, row in enumerate(matrix_data, 1):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = BG_COLOR if i % 2 == 1 else CARD_BG
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = FONT_BODY
            p.font.size = Pt(10.5)
            p.font.bold = (j > 0)
            p.font.color.rgb = WHITE if j == 0 else LIGHT_TEXT
            p.alignment = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT

    # ========================================================
    # SLIDE 25: Grand Challenge
    # ========================================================
    s25 = create_slide(
        "Grand Challenge • Comprehensive Assessment",
        "The Ultimate Test: Sarah's New Laptop",
        "Combining possessive adjectives and possessive pronouns in a single sentence",
        "KỊCH BẢN NÓI: Thử thách tổng kết: câu hỏi kết hợp cả tính từ sở hữu và đại từ sở hữu. Chỗ trống 1 đi trước danh từ 'laptop' -> Her. Chỗ trống 2 đứng cuối câu -> hers. Đáp án đúng là B: Her / hers."
    )
    c25 = add_card(s25, 1.5, 1.9, 10.33, 4.8, border_color=RED)
    tf25 = c25.text_frame
    tf25.word_wrap = True
    tf25.margin_left = Inches(0.5)
    tf25.margin_top = Inches(0.3)
    p = tf25.paragraphs[0]
    p.text = "[GRAND CHALLENGE QUESTION]"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = RED
    p = tf25.add_paragraph()
    p.text = "“Sarah has a new laptop. _____ laptop is very expensive, but the laptop is not _____.”"
    p.font.name = FONT_BODY
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = WHITE
    p.space_before = Pt(8)

    p = tf25.add_paragraph()
    p.text = "Options:\nA. Hers / her\nB. Her / hers  [CORRECT ANSWER]\nC. She / her\nD. Her / she"
    p.font.name = FONT_BODY
    p.font.size = Pt(13)
    p.font.color.rgb = LIGHT_TEXT
    p.space_before = Pt(12)

    p = tf25.add_paragraph()
    p.text = "Detailed Breakdown:\n1. 'Her laptop': Precedes the noun 'laptop' ➔ Possessive Adjective (Her).\n2. 'not hers': Stands independently at sentence end without an accompanying noun ➔ Possessive Pronoun (Hers)."
    p.font.name = FONT_BODY
    p.font.size = Pt(12)
    p.font.color.rgb = GREEN
    p.space_before = Pt(14)

    # ========================================================
    # SLIDE 26: Conclusion & Discussion
    # ========================================================
    s26 = create_slide(
        "Conclusion • Key Takeaway",
        "Small Words, Massive Impact",
        "Why mastering pronouns transforms both fluency and precision",
        "KỊCH BẢN NÓI: Xin cảm ơn quý thầy cô và các bạn đã chú ý lắng nghe. Giờ là phần thảo luận và giải đáp thắc mắc!"
    )
    c26 = add_card(s26, 1.5, 2.0, 10.33, 4.5, border_color=PRIMARY)
    tf26 = c26.text_frame
    tf26.word_wrap = True
    tf26.margin_left = Inches(0.6)
    tf26.margin_top = Inches(0.6)
    p = tf26.paragraphs[0]
    p.text = '“Pronouns may be small words, but they are very important in English.”'
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(22)
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    
    p = tf26.add_paragraph()
    p.text = "Pronouns appear in almost every sentence. By mastering who/what they replace,\ntheir grammatical position, and ownership rules, your English becomes natural, fluent, and precise."
    p.font.name = FONT_BODY
    p.font.size = Pt(14)
    p.font.color.rgb = MUTED
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(20)

    p = tf26.add_paragraph()
    p.text = "Thank You! Opening Floor for Q&A"
    p.font.name = FONT_TITLE
    p.font.bold = True
    p.font.size = Pt(15)
    p.font.color.rgb = PRIMARY
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(30)

    # Save output
    out_path = r"C:\Users\Win\.gemini\antigravity-ide\scratch\pronouns-presentation\Pronouns_Keynote_Masterclass_26_Slides.pptx"
    prs.save(out_path)
    print(f"Presentation saved successfully to {out_path} with {len(prs.slides)} slides.")

if __name__ == "__main__":
    build_presentation()
