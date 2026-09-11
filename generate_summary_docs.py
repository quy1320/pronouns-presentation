import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_callout_box(doc, text_prefix, title, body_text, bg_hex="F0F4F8", border_hex="3B82F6"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
    # Left border only
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_hex}"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    
    r_prefix = p.add_run(f"{text_prefix} ")
    r_prefix.bold = True
    r_prefix.font.size = Pt(10.5)
    r_prefix.font.name = "Arial"
    
    r_title = p.add_run(f"{title}\n")
    r_title.bold = True
    r_title.font.size = Pt(10.5)
    r_title.font.name = "Arial"
    
    r_body = p.add_run(body_text)
    r_body.font.size = Pt(10)
    r_body.font.name = "Arial"
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_word_document():
    doc = docx.Document()
    
    # Page setup - Margins 0.8 inch
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)
        
        # Header / Footer
        footer = s.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("Tóm tắt 32 Slide: English Pronouns Masterclass | Tài liệu học tập & tra cứu nhanh")
        f_run.font.size = Pt(8.5)
        f_run.font.color.rgb = RGBColor(148, 163, 184)
        f_run.font.name = "Arial"
    
    # Document Title Block
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    r_main_title = title_p.add_run("TÓM TẮT TRỌNG TÂM 32 SLIDE BÁO CÁO\nCHỦ ĐỀ: ĐẠI TỪ TRONG TIẾNG ANH (PRONOUNS)")
    r_main_title.font.name = "Arial"
    r_main_title.font.size = Pt(20)
    r_main_title.bold = True
    r_main_title.font.color.rgb = RGBColor(15, 23, 42) # Deep Navy
    
    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(2)
    sub_p.paragraph_format.space_after = Pt(14)
    r_sub = sub_p.add_run("Bản Đồ Ghi Nhớ Nhanh • Công Thức Đại Số Ngôn Ngữ • Quy Tắc Vàng & Giải Mã 12 Câu Hỏi Thực Chiến")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(79, 70, 229) # Indigo
    
    # Metadata Badge Box
    meta_tbl = doc.add_table(rows=1, cols=3)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False
    widths = [Inches(2.2), Inches(2.3), Inches(2.2)]
    cell_data = [
        ("📚 Cấu trúc báo cáo", "32 Slide Chuẩn Quốc Tế"),
        ("🎯 Trọng tâm nội dung", "8 Gia Đình Đại Từ & Bẫy TOEIC"),
        ("⚡ Mục đích tài liệu", "Tóm Tắt Dễ Nhớ - Gửi Đi Ngay")
    ]
    for idx, (label, val) in enumerate(cell_data):
        c = meta_tbl.cell(0, idx)
        c.width = widths[idx]
        set_cell_background(c, "F8FAFC")
        set_cell_margins(c, top=80, bottom=80, left=100, right=100)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(f"{label}\n")
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = RGBColor(100, 116, 139)
        r2 = p.add_run(val)
        r2.bold = True
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = RGBColor(30, 41, 59)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    # Table of Contents Overview Table
    toc_heading = doc.add_heading("I. MA TRẬN PHÂN BỔ 32 SLIDE THEO 12 PHẦN NỘI DUNG", level=1)
    toc_heading.paragraph_format.space_before = Pt(14)
    toc_heading.paragraph_format.space_after = Pt(6)
    
    toc_table = doc.add_table(rows=13, cols=3)
    toc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    toc_table.autofit = False
    col_widths = [Inches(1.2), Inches(2.8), Inches(2.7)]
    headers = ["Phân Nhóm", "Slide Tương Ứng", "Trọng Tâm Kiến Thức Cốt Lõi"]
    
    # Format Header Row
    for i, h_text in enumerate(headers):
        c = toc_table.cell(0, i)
        c.width = col_widths[i]
        set_cell_background(c, "1E293B")
        set_cell_margins(c, top=100, bottom=100, left=120, right=120)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(h_text)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    toc_rows = [
        ("Phần 1", "Slide 01 - 04", "Đặt vấn đề, định nghĩa Pronoun & Lộ trình 8 gia đình đại từ"),
        ("Phần 2", "Slide 05 - 08", "Đại từ Nhân xưng (Subject vs Object), mỏ neo câu & Quiz 01 - 02"),
        ("Phần 3", "Slide 09 - 12", "Đại từ Sở hữu, công thức đại số (mine = my + N) & Quiz 03 - 04"),
        ("Phần 4", "Slide 13 - 14", "Đại từ Phản thân (3 chức năng: S=O, nhấn mạnh, by myself) & Quiz 05"),
        ("Phần 5", "Slide 15 - 17", "Đại từ Chỉ định (This/That/These/Those), N không đếm được & Quiz 06 - 07"),
        ("Phần 6", "Slide 18 - 19", "Đại từ Bất định & Luật Động từ số ít bắt buộc (is/has) & Quiz 08"),
        ("Phần 7", "Slide 20 - 22", "Đại từ Nghi vấn (Who, Whom, Whose, What, Which) & Quiz 09 - 10"),
        ("Phần 8", "Slide 23 - 25", "Đại từ Quan hệ (Ma trận TOEIC, 2 cấm kỵ của THAT, rút gọn) & Quiz 11"),
        ("Phần 9", "Slide 26 - 27", "Đại từ Ghép nâng cao (Whoever = Anyone who, Whichever) & Quiz 12"),
        ("Phần 10", "Slide 28 - 29", "Bộ khung chiến thuật 3 câu hỏi & Bảng cheatsheet tối thượng"),
        ("Phần 11", "Slide 30 - 31", "Grand Challenge (Sarah's Laptop) & Vòng phản xạ tương tác tốc độ"),
        ("Phần 12", "Slide 32", "Thông điệp đúc kết vàng & Mở phiên thảo luận Q&A")
    ]
    
    for r_idx, (col1, col2, col3) in enumerate(toc_rows, start=1):
        bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate([col1, col2, col3]):
            c = toc_table.cell(r_idx, c_idx)
            c.width = col_widths[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=60, bottom=60, left=100, right=100)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(val)
            r.font.size = Pt(9)
            if c_idx == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(79, 70, 229)
                
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # Detailed Slide-by-Slide Section
    sec_heading = doc.add_heading("II. TÓM TẮT CHI TIẾT TỪNG SLIDE (DỄ NHỚ - CÔNG THỨC - MẸO THI)", level=1)
    sec_heading.paragraph_format.space_before = Pt(16)
    sec_heading.paragraph_format.space_after = Pt(10)
    
    slides_summary_data = [
        # Slide 1
        {
            "num": "01",
            "title": "Title & Cover: Pronouns in English",
            "cat": "English Communication Insights",
            "key_takeaway": "Giới thiệu chủ đề: Đại từ là từ ngắn nhưng có sức ảnh hưởng cực lớn đến độ tự nhiên, trôi chảy và chuẩn xác trong tiếng Anh.",
            "formula": "Khung bài báo cáo: 8 gia đình đại từ + Công thức đại số + 12 câu hỏi tương tác.",
            "memory_trick": "💡 Mẹo nhớ: Muốn tiếng Anh tự nhiên như người bản xứ, việc đầu tiên là không được lặp lại danh từ liên tục.",
            "quiz": None
        },
        # Slide 2
        {
            "num": "02",
            "title": "The Problem with Repetition (Vấn đề lặp từ & Case Study Tom)",
            "cat": "01. Introduction",
            "key_takeaway": "Khi nghe đoạn văn: 'Tom is a student. Tom likes football. Tom plays football every day...', người nghe thấy rất mệt mỏi và robot vì từ 'Tom' bị lặp lại 4 lần.",
            "formula": "Giải pháp: 'Tom is a student. He likes football. He plays football every day.'",
            "memory_trick": "💡 Mẹo nhớ: 'Đại từ sinh ra là để cứu vãn sự lặp từ!'. Từ 'He' mượt mà thay thế danh từ Tom.",
            "quiz": None
        },
        # Slide 3
        {
            "num": "03",
            "title": "What is a Pronoun & Why Do We Need Them? (Định nghĩa & 4 Lợi ích)",
            "cat": "01. Introduction",
            "key_takeaway": "Từ nguyên Latin: 'Pro' (thay thế) + 'Noun' (danh từ) ➔ Đại từ là từ đứng ra đại diện và thay thế cho Danh từ hoặc Cụm danh từ (Tom ➔ he, Anna ➔ she, Book ➔ it, Tom & Anna ➔ they).",
            "formula": "4 Lợi ích vàng: (1) Tránh lặp từ máy móc; (2) Câu ngắn gọn, súc tích; (3) Tạo nhịp điệu tự nhiên bản xứ; (4) Ngữ cảnh rõ ràng, không mơ hồ.",
            "memory_trick": "💡 Mẹo nhớ: Pro = Thay thế. Đại từ = Người đại diện pháp lý cho danh từ!",
            "quiz": None
        },
        # Slide 4
        {
            "num": "04",
            "title": "The 8 Essential Pronoun Families (Lộ trình 8 Gia đình Đại từ)",
            "cat": "02. Executive Overview",
            "key_takeaway": "Bộ khung 8 trụ cột toàn diện: Nền tảng (Nhân xưng, Sở hữu, Phản thân) ➔ Định vị (Chỉ định, Bất định) ➔ Liên kết & Mệnh đề (Nghi vấn, Quan hệ, Ghép nâng cao).",
            "formula": "8 Trụ cột: Personal ➔ Possessive ➔ Reflexive ➔ Demonstrative ➔ Indefinite ➔ Interrogative ➔ Relative ➔ Compound.",
            "memory_trick": "💡 Mẹo nhớ: Học theo cây phả hệ 8 nhánh, không học vẹt từng từ riêng lẻ.",
            "quiz": None
        },
        # Slide 5
        {
            "num": "05",
            "title": "Subject Pronouns (Đại từ làm Chủ ngữ)",
            "cat": "03. Personal Pronouns",
            "key_takeaway": "Chủ ngữ là chủ thể thực hiện hành động. Gồm 7 đại từ: I, You, He, She, It, We, They.",
            "formula": "Vị trí chuẩn: [ Subject Pronoun ] + Verb (Luôn đứng TRƯỚC động từ chính).",
            "memory_trick": "💡 Ví dụ & Mẹo: 'He plays football' (He đứng trước động từ plays). Mỏ neo: 'I' trong 'I love you'.",
            "quiz": None
        },
        # Slide 6
        {
            "num": "06",
            "title": "Object Pronouns (Đại từ làm Tân ngữ)",
            "cat": "03. Personal Pronouns",
            "key_takeaway": "Tân ngữ là đối tượng tiếp nhận tác động của hành động. Biến đổi: I➔me, You➔you, He➔him, She➔her, It➔it, We➔us, They➔them.",
            "formula": "Vị trí chuẩn: Verb / Preposition + [ Object Pronoun ] (Luôn đứng SAU động từ hoặc giới từ).",
            "memory_trick": "💡 Mô hình Anna & Peter: 'She likes him' (She là người thích ➔ Subject; him là người được thích ➔ Object. Không bao giờ nói 'She likes he').",
            "quiz": None
        },
        # Slide 7
        {
            "num": "07",
            "title": "Interactive Check 01 (Trắc nghiệm Chủ ngữ - Tom)",
            "cat": "Interactive Check - Question 01",
            "key_takeaway": "Kiểm tra nhận diện vị trí Chủ ngữ trước động từ.",
            "formula": "Đứng trước động từ 'is' ➔ Cần Subject Pronoun.",
            "memory_trick": "Tom là nam giới ➔ Chọn He.",
            "quiz": {
                "question": "Tom is my friend. ___ is very nice.",
                "options": "A. He | B. Him | C. His | D. Himself",
                "answer": "✅ Đáp án đúng: A. He",
                "explanation": "Tom đứng ở đầu mệnh đề thứ 2, trước động từ 'is' nên giữ vai trò Chủ ngữ ➔ Tom = He."
            }
        },
        # Slide 8
        {
            "num": "08",
            "title": "Interactive Check 02 (Trắc nghiệm Tân ngữ - Anna & Peter)",
            "cat": "Interactive Check - Question 02",
            "key_takeaway": "Kiểm tra thay thế danh từ đứng sau động từ chỉ hành động.",
            "formula": "Đứng sau động từ 'likes' ➔ Bắt buộc mang dạng Tân ngữ (Object).",
            "memory_trick": "Peter là nam giới ở vị trí tân ngữ ➔ him.",
            "quiz": {
                "question": "“Anna likes Peter.” — If we replace “Peter”, which sentence is correct?",
                "options": "A. Anna likes he. | B. Anna likes him. | C. Anna likes his. | D. Anna likes himself.",
                "answer": "✅ Đáp án đúng: B. Anna likes him.",
                "explanation": "Peter là người nhận tình cảm từ Anna sau động từ 'likes', do đó phải chuyển thành tân ngữ 'him'."
            }
        },
        # Slide 9
        {
            "num": "09",
            "title": "Possessive Pronouns (Đại từ Sở hữu độc lập)",
            "cat": "04. Possessive Pronouns",
            "key_takeaway": "Dùng để trả lời câu hỏi quyền sở hữu: 'Whose is this?' (Cái này của ai?) mà không cần lặp lại danh từ đồ vật.",
            "formula": "Bảng quy chiếu: I➔mine, You➔yours, He➔his, She➔hers, It➔its, We➔ours, They➔theirs.",
            "memory_trick": "💡 Mẹo nhớ: Từ 'mine' trong 'This book is mine' đứng độc lập cuối câu, tự thân nó đã mang nghĩa 'của tôi', không cần thêm từ nào nữa!",
            "quiz": None
        },
        # Slide 10
        {
            "num": "10",
            "title": "Possessive Adjective vs. Possessive Pronoun (Công thức Đại số)",
            "cat": "04. Possessive Pronouns",
            "key_takeaway": "Điểm ngữ pháp dễ nhầm lẫn số 1 trong các bài thi: Phân biệt Tính từ sở hữu và Đại từ sở hữu.",
            "formula": "PHƯƠNG TRÌNH ĐẠI SỐ:\nĐại từ Sở hữu (mine, yours, hers...) = Tính từ Sở hữu (my, your, her...) + Danh từ (Noun)\n• mine = my books | hers = her laptop",
            "memory_trick": "💡 Mẹo vàng nằm lòng:\n- Có Danh từ theo sau ➔ Dùng Tính từ sở hữu (My house is beautiful).\n- Đứng MỘT MÌNH lẻ loi cuối câu / thay thế Noun đã biết ➔ Dùng Đại từ sở hữu (This house is mine).",
            "quiz": None
        },
        # Slide 11
        {
            "num": "11",
            "title": "Interactive Check 03 (Trắc nghiệm có Noun theo sau)",
            "cat": "Interactive Check - Question 03",
            "key_takeaway": "Áp dụng công thức có danh từ phía sau chỗ trống.",
            "formula": "[ my/your/his... ] + Noun.",
            "memory_trick": "Thấy 'book' là danh từ ➔ Phải chọn Tính từ sở hữu 'my'.",
            "quiz": {
                "question": "This is ___ book.",
                "options": "A. mine | B. my | C. me | D. I",
                "answer": "✅ Đáp án đúng: B. my",
                "explanation": "Ngay sau khoảng trống là danh từ 'book', theo công thức [my + noun], ta bắt buộc chọn tính từ sở hữu 'my'."
            }
        },
        # Slide 12
        {
            "num": "12",
            "title": "Interactive Check 04 (Trắc nghiệm đứng độc lập cuối câu)",
            "cat": "Interactive Check - Question 04",
            "key_takeaway": "Áp dụng công thức đứng độc lập không có danh từ theo sau.",
            "formula": "[ mine/yours/hers... ] đứng một mình.",
            "memory_trick": "Sau 'is' và trước dấu chấm không có danh từ ➔ Chọn 'mine'.",
            "quiz": {
                "question": "This book is ___.",
                "options": "A. my | B. me | C. mine | D. I",
                "answer": "✅ Đáp án đúng: C. mine",
                "explanation": "Sau chỗ trống là dấu chấm kết câu, hoàn toàn không có danh từ nào ➔ Chọn đại từ sở hữu độc lập 'mine'."
            }
        },
        # Slide 13
        {
            "num": "13",
            "title": "Reflexive Pronouns: 3 Strategic Functions (3 Công năng Phản thân)",
            "cat": "05. Reflexive Pronouns",
            "key_takeaway": "Đại từ phản thân có 3 chức năng khác nhau chứ không chỉ đơn thuần là 'tự làm mình':\n1. Phản chiếu (Subject = Object): 'I love myself'.\n2. Nhấn mạnh (Emphatic): Đích thân làm ('She carries these books herself').\n3. Thành ngữ: 'by + reflexive' = alone / on one's own ('I did it by myself' - tự làm một mình không ai giúp).",
            "formula": "Số ít đuôi -self (myself, yourself, himself, herself, itself) | Số nhiều đuôi -selves (ourselves, yourselves, themselves).",
            "memory_trick": "💡 Mẹo nhớ: Giống như soi gương (Mirror) - Mình làm hành động và hình ảnh phản chiếu quay lại chính mình!",
            "quiz": None
        },
        # Slide 14
        {
            "num": "14",
            "title": "Interactive Check 05 (Trắc nghiệm Tai nạn nấu ăn)",
            "cat": "Interactive Check - Question 05",
            "key_takeaway": "Nhận diện người gây ra hành động và người bị tác động là một.",
            "formula": "Subject = Object ➔ Dùng Reflexive Pronoun.",
            "memory_trick": "John tự cầm dao cắt vào tay mình ➔ himself.",
            "quiz": {
                "question": "John cut ___ while cooking.",
                "options": "A. him | B. his | C. himself | D. he",
                "answer": "✅ Đáp án đúng: C. himself",
                "explanation": "John là người cắt và chính John là người bị thương (Doer = Receiver) ➔ Chọn đại từ phản thân 'himself'."
            }
        },
        # Slide 15
        {
            "num": "15",
            "title": "Demonstrative Pronouns (Ma trận Không gian & Noun không đếm được)",
            "cat": "04. Demonstrative Pronouns",
            "key_takeaway": "Đại từ chỉ định xác định vật theo khoảng cách (Gần/Xa) và số lượng (Ít/Nhiều).\n- GẦN (Near): THIS (1 cái) / THESE (nhiều cái)\n- XA (Far): THAT (1 cái) / THOSE (nhiều cái)",
            "formula": "QUY TẮC ĐẮT GIÁ: Danh từ không đếm được (Uncountable Nouns) luôn dùng THIS / THAT (This information, That advice, This water).",
            "memory_trick": "💡 Mẹo nhớ: Cầm trên tay ➔ This/These. Chỉ ngón tay ra đằng xa ➔ That/Those.",
            "quiz": None
        },
        # Slide 16
        {
            "num": "16",
            "title": "Interactive Check 06 (Cầm sách trên tay)",
            "cat": "Interactive Check - Question 06",
            "key_takeaway": "Tọa độ: Cự ly GẦN + SỐ ÍT (1 quyển).",
            "formula": "Near + Singular ➔ THIS.",
            "memory_trick": "Sách trên tay mình ➔ This is my book.",
            "quiz": {
                "question": "You are holding a book in your hand. You say: “___ is my book.”",
                "options": "A. Those | B. These | C. This | D. That",
                "answer": "✅ Đáp án đúng: C. This",
                "explanation": "Quyển sách đang cầm trên tay là cự ly GẦN và chỉ có 1 quyển (Số ít) ➔ Dùng 'This'."
            }
        },
        # Slide 17
        {
            "num": "17",
            "title": "Interactive Check 07 (Ô tô đỗ đằng xa)",
            "cat": "Interactive Check - Question 07",
            "key_takeaway": "Tọa độ: Cự ly XA + SỐ NHIỀU (nhiều ô tô).",
            "formula": "Far + Plural ➔ THOSE.",
            "memory_trick": "Xe ở cuối phố (xa) + are cars (nhiều) ➔ Those are cars.",
            "quiz": {
                "question": "You see multiple cars parked far down the road. You say: “___ are cars.”",
                "options": "A. This | B. That | C. These | D. Those",
                "answer": "✅ Đáp án đúng: D. Those",
                "explanation": "Nhiều chiếc xe (Số nhiều) và nằm ở khoảng cách xa tầm mắt ➔ Dùng 'Those'."
            }
        },
        # Slide 18
        {
            "num": "18",
            "title": "Indefinite Pronouns & Singular Verb Law (Luật Động từ Số ít)",
            "cat": "05. Indefinite Pronouns",
            "key_takeaway": "Đại từ bất định chỉ đối tượng không xác định cụ thể:\n- Chỉ người: Someone, Anyone, Everyone, No one (và đuôi -body).\n- Chỉ vật: Something, Anything, Everything, Nothing.",
            "formula": "QUY TẮC VÀNG TOEIC & HỌC THUẬT BẤT BIẾN:\nIndefinite Pronoun + SINGULAR VERB (Động từ số ít: is, was, has, V-s/es).\n• ❌ Everyone are ready ➔ ✅ Everyone IS ready.\n• ❌ Somebody have left ➔ ✅ Somebody HAS left.",
            "memory_trick": "💡 Bẫy tâm lý: Dù dịch là 'mọi người' (nghe như số nhiều), ngữ pháp tiếng Anh LUÔN chia động từ số ít!",
            "quiz": None
        },
        # Slide 19
        {
            "num": "19",
            "title": "Interactive Check 08 (Hội đồng quản trị - Everyone)",
            "cat": "Interactive Check - Question 08",
            "key_takeaway": "Hóa giải bẫy số nhiều trong đề thi chuẩn hóa.",
            "formula": "Everyone + Verb số ít (has).",
            "memory_trick": "Thấy 'Everyone' ➔ Bỏ qua cụm bổ ngữ 'on the executive board', chia ngay động từ số ít 'has'.",
            "quiz": {
                "question": "“Everyone on the executive board ___ agreed with the proposed budget.”",
                "options": "A. have | B. has | C. are | D. were",
                "answer": "✅ Đáp án đúng: B. has",
                "explanation": "'Everyone' là đại từ bất định bắt buộc chia động từ số ít ➔ 'has agreed' (Không bao giờ dùng 'have' hay 'are/were')."
            }
        },
        # Slide 20
        {
            "num": "20",
            "title": "Interrogative Pronouns (5 Đại từ Nghi vấn đặt câu hỏi)",
            "cat": "06. Interrogative Pronouns",
            "key_takeaway": "5 công cụ ngôn ngữ để khởi tạo câu hỏi chính xác:\n- WHO: Hỏi người (Chủ ngữ) ➔ 'Who is she?'\n- WHOM: Hỏi người (Tân ngữ, trang trọng) ➔ 'Whom did you meet?'\n- WHOSE: Hỏi sở hữu (Của ai) ➔ 'Whose phone is this?'\n- WHAT: Hỏi sự vật/hành động tổng quát ➔ 'What is this?'\n- WHICH: Hỏi lựa chọn trong danh sách giới hạn ➔ 'Which one do you prefer?'",
            "formula": "Who (chủ ngữ người) vs Whom (tân ngữ người) vs Whose (sở hữu) vs Which (lựa chọn giới hạn).",
            "memory_trick": "💡 Mẹo nhớ: Có sự lựa chọn giữa các phương án cụ thể ➔ Luôn dùng WHICH (không dùng What).",
            "quiz": None
        },
        # Slide 21
        {
            "num": "21",
            "title": "Interactive Check 09 (Hỏi bạn thân nhất)",
            "cat": "Interactive Check - Question 09",
            "key_takeaway": "Hỏi về con người ở vị trí chủ ngữ.",
            "formula": "Who + is + Person?",
            "memory_trick": "Người làm bạn thân nhất là ai ➔ Who.",
            "quiz": {
                "question": "“___ is your best friend?”",
                "options": "A. What | B. Who | C. Whose | D. Which",
                "answer": "✅ Đáp án đúng: B. Who",
                "explanation": "Câu hỏi nhắm vào một người cụ thể ở vị trí chủ ngữ ➔ Dùng từ để hỏi 'Who'."
            }
        },
        # Slide 22
        {
            "num": "22",
            "title": "Interactive Check 10 (Hỏi chủ sở hữu điện thoại)",
            "cat": "Interactive Check - Question 10",
            "key_takeaway": "Truy vấn quyền sở hữu của đồ vật để quên.",
            "formula": "Whose + Noun + is this?",
            "memory_trick": "Hỏi 'của ai' ➔ Whose.",
            "quiz": {
                "question": "“___ phone is this?”",
                "options": "A. Who | B. What | C. Whose | D. Which",
                "answer": "✅ Đáp án đúng: C. Whose",
                "explanation": "Câu hỏi điều tra về người sở hữu chiếc điện thoại (Whose = Của ai) ➔ Dùng 'Whose'."
            }
        },
        # Slide 23
        {
            "num": "23",
            "title": "Relative Pronouns: Clause Connectors (Ma trận Công thức TOEIC)",
            "cat": "07. Relative Pronouns",
            "key_takeaway": "Nối 2 câu đơn thành 1 câu phức mạch lạc không lặp từ. 4 Công thức thi TOEIC chuẩn xác:\n1. WHO (thay cho người làm chủ ngữ): N(người) + WHO + Verb\n2. WHOM (thay cho người làm tân ngữ): N(người) + WHOM + S + V\n3. WHICH (thay cho vật): N(vật) + WHICH + V / S+V\n4. WHOSE (chỉ quan hệ sở hữu): N + WHOSE + Noun + V",
            "formula": "Ma trận TOEIC:\n• Người + WHO + V\n• Người + WHOM + S + V\n• Vật + WHICH + V / S+V\n• Noun + WHOSE + Noun + V",
            "memory_trick": "💡 Mẹo nhớ 5 giây: Nhìn từ đứng NGAY TRƯỚC và thành phần NGAY SAU chỗ trống để áp đúng công thức.",
            "quiz": None
        },
        # Slide 24
        {
            "num": "24",
            "title": "The 'THAT' Restrictions & Participle Reduction (2 Vùng cấm & Bẫy rút gọn)",
            "cat": "07. Relative Pronouns",
            "key_takeaway": "2 Lưu ý sống còn trong các bài thi ngữ pháp cấp cao:\n1. 'THAT' rất đa năng (thay cho cả Who và Which), NHƯNG CÓ 2 VÙNG CẤM TUYỆT ĐỐI:\n   - ❌ Tuyệt đối KHÔNG đứng sau dấu phẩy (,)\n   - ❌ Tuyệt đối KHÔNG đứng sau giới từ (in that, with that...)\n2. BẪY RÚT GỌN MỆNH ĐỀ PHÂN TỪ: 1 câu chỉ có 1 động từ chính. Nếu bỏ đại từ 'who/which':\n   - Chủ động ➔ chuyển động từ thành V-ing ('The woman living near my house')\n   - Bị động ➔ chuyển động từ thành V-ed/3 ('The proposal submitted yesterday')",
            "formula": "• Dấu phẩy (,) hoặc Giới từ ➔ CẤM CHỌN THAT!\n• Bỏ Who/Which ➔ V biến thành V-ing (chủ động) hoặc V-ed/3 (bị động).",
            "memory_trick": "💡 Dòng ghi chú viết tay: 'Thấy dấu phẩy là đuổi THAT đi ngay!'.",
            "quiz": None
        },
        # Slide 25
        {
            "num": "25",
            "title": "Interactive Check 11 (Bẫy dấu phẩy - Senior Manager)",
            "cat": "Interactive Check - Question 11",
            "key_takeaway": "Ứng dụng luật trừ dần: Người ➔ Dấu phẩy ➔ Vị trí tân ngữ.",
            "formula": "N(người) + , + WHOM + S + V.",
            "memory_trick": "Có dấu phẩy loại THAT. Người loại WHICH. Sau có 'we met' (S+V) ➔ Chọn WHOM.",
            "quiz": {
                "question": "“The senior manager, ___ we met at the corporate summit, approved our project.”",
                "options": "A. that | B. which | C. whom | D. whose",
                "answer": "✅ Đáp án đúng: C. whom",
                "explanation": "1. 'manager' là người ➔ loại Which. 2. Có dấu phẩy (,) ➔ Cấm chọn That. 3. Phía sau là 'we met' (S+V) ➔ Tân ngữ chỉ người bắt buộc chọn 'whom'."
            }
        },
        # Slide 26
        {
            "num": "26",
            "title": "Advanced Compounds: Whoever & Whichever (Chuyên đề Nâng cao)",
            "cat": "08. Advanced Compound Pronouns",
            "key_takeaway": "1. WHOEVER = Anyone who (Bất kỳ ai mà...): Đi với động từ số ít trong đề thi TOEIC. Ví dụ: 'Whoever wants to join can register here'.\n2. WHICHEVER = Any from a limited set (Bất kỳ cái nào trong một nhóm đã giới hạn trước). Rất đa năng: làm Chủ ngữ, Tân ngữ hoặc Tính từ ('Whichever option you choose').",
            "formula": "PHƯƠNG TRÌNH TOEIC:\nWhoever + V(số ít) = Anyone who + V(số ít)",
            "memory_trick": "💡 Mẹo TOEIC Part 5: Nếu đầu câu là chỗ trống, ngay sau là động từ chia số ít (arrives, is, wants) và không có danh từ đi trước ➔ Chọn ngay WHOEVER!",
            "quiz": None
        },
        # Slide 27
        {
            "num": "27",
            "title": "Interactive Check 12 (Access Badges - Access Protocol)",
            "cat": "Interactive Check - Question 12",
            "key_takeaway": "Nhận diện đại từ ghép làm chủ ngữ toàn câu.",
            "formula": "Whoever + arrives = Anyone who arrives.",
            "memory_trick": "'Anyone' đứng một mình không thể nối mệnh đề có động từ 'arrives'. Phải là 'Whoever'.",
            "quiz": {
                "question": "“___ arrives at the conference venue first should collect the company access badges.”",
                "options": "A. Anyone | B. Whoever | C. Whom | D. Whichever",
                "answer": "✅ Đáp án đúng: B. Whoever",
                "explanation": "'Anyone' cần có 'who' mới nối được mệnh đề. 'Whoever' = 'Anyone who' đóng vai trò chủ ngữ hoàn chỉnh của câu."
            }
        },
        # Slide 28
        {
            "num": "28",
            "title": "The 3 Golden Questions for Any Sentence (Thuật toán 3 Câu hỏi vàng)",
            "cat": "09. Strategic Decision Framework",
            "key_takeaway": "Bộ khung tư duy phản xạ giải mã đại từ trong 3 giây:\n- CÂU HỎI 1: Ai làm hành động? ➔ Chọn Subject Pronoun (He, She, They).\n- CÂU HỎI 2: Ai / Cái gì nhận tác động? ➔ Chọn Object Pronoun (him, her, them).\n- CÂU HỎI 3: Tác động có quay ngược lại chính mình không? ➔ Chọn Reflexive Pronoun (himself, themselves).",
            "formula": "Q1 (Doer ➔ Subject) | Q2 (Receiver ➔ Object) | Q3 (Self-bounce ➔ Reflexive).",
            "memory_trick": "💡 Mẹo nhớ: Hỏi đúng 3 câu này sẽ triệt tiêu 100% sự lúng túng khi làm bài thi hay giao tiếp!",
            "quiz": None
        },
        # Slide 29
        {
            "num": "29",
            "title": "All-in-One Pronouns Master Matrix (Bảng Cheatsheet Tối thượng)",
            "cat": "10. Complete Synthesis",
            "key_takeaway": "Bảng tổng hợp toàn diện cả 8 gia đình đại từ: Subject, Object, Possessive, Reflexive, Demonstrative, Indefinite, Interrogative, Relative & Compound.",
            "formula": "Bảng đối chiếu tổng tài liệu để ôn tập nhanh trước kỳ thi.",
            "memory_trick": "💡 Mẹo tra cứu: In bảng này dán trước bàn học hoặc lưu vào điện thoại làm tài liệu tham khảo nhanh.",
            "quiz": None
        },
        # Slide 30
        {
            "num": "30",
            "title": "Grand Challenge: Sarah's Laptop (Thử thách Đỉnh cao)",
            "cat": "Interactive Case - Grand Challenge",
            "key_takeaway": "Câu đố tổng hợp lồng ghép cả Tính từ sở hữu và Đại từ sở hữu trong cùng 1 câu.",
            "formula": "Chỗ trống 1: Trước 'laptop' (Noun) ➔ Tính từ sở hữu 'Her'.\nChỗ trống 2: Đứng độc lập cuối câu (No Noun) ➔ Đại từ sở hữu 'hers'.",
            "memory_trick": "Her laptop (có Noun) vs not hers (không Noun).",
            "quiz": {
                "question": "Sarah has a new laptop. ___ laptop is very expensive, but the laptop is not ___.",
                "options": "A. Hers / her | B. Her / hers | C. She / her | D. Her / she",
                "answer": "✅ Đáp án đúng: B. Her / hers",
                "explanation": "1. 'Her laptop': Đứng trước danh từ 'laptop' cần Tính từ sở hữu 'Her'.\n2. 'not hers': Đứng độc lập cuối câu sau 'is not' không có danh từ cần Đại từ sở hữu 'hers'."
            }
        },
        # Slide 31
        {
            "num": "31",
            "title": "Rapid Fire: Spot the Target Pronoun! (Vòng Phản xạ Tốc độ)",
            "cat": "11. Audience Engagement",
            "key_takeaway": "5 câu phản xạ chớp nhoáng kiểm tra phản xạ tức thì:",
            "formula": "1. Tom is a student. [ HE ] studies English. (Chủ ngữ thay thế Tom)\n2. I like Anna. I often talk to [ HER ]. (Tân ngữ nhận hành động)\n3. This is my pen. The pen is [ MINE ]. (Đại từ sở hữu độc lập)\n4. He made the cake by [ HIMSELF ]. (Thành ngữ by himself = tự làm một mình)\n5. [ THESE ] are my shoes. (Chỉ vào giày dưới chân = Gần + Số nhiều)",
            "memory_trick": "💡 Luyện phản xạ: Đọc đề ➔ Bắt từ khóa ➔ Bật ngay đại từ trong 1 giây.",
            "quiz": None
        },
        # Slide 32
        {
            "num": "32",
            "title": "Small Words, Massive Impact (Đúc kết & Mở phiên Q&A)",
            "cat": "12. Conclusion & Discussion",
            "key_takeaway": "Thông điệp đúc kết: 'Pronouns may be small words, but they are very important in English.' Đại từ tuy ngắn nhưng là sợi dây kết nối toàn bộ sự mạch lạc, tự nhiên và chuẩn mực.",
            "formula": "3 Trụ cột thành thạo: Biết thay thế ai/cái gì + Đặt đúng vị trí ngữ pháp + Tận dụng sức mạnh liên kết mệnh đề.",
            "memory_trick": "💡 Lời kết: Làm chủ Đại từ là bước chuyển ngoạn mục từ tiếng Anh 'dịch từng từ máy móc' sang tiếng Anh 'bản xứ lưu loát'.",
            "quiz": None
        }
    ]
    
    for s_data in slides_summary_data:
        s_h = doc.add_heading(f"Slide {s_data['num']}: {s_data['title']}", level=2)
        s_h.paragraph_format.space_before = Pt(12)
        s_h.paragraph_format.space_after = Pt(4)
        
        # Category tag
        p_cat = doc.add_paragraph()
        p_cat.paragraph_format.space_before = Pt(0)
        p_cat.paragraph_format.space_after = Pt(4)
        r_tag = p_cat.add_run(f"📂 Phân nhóm: {s_data['cat']}")
        r_tag.font.size = Pt(9.5)
        r_tag.bold = True
        r_tag.font.color.rgb = RGBColor(14, 165, 233) # Sky blue
        
        # Key takeaway
        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.space_before = Pt(2)
        p_desc.paragraph_format.space_after = Pt(4)
        r_label = p_desc.add_run("🎯 Ý chính & Bản chất: ")
        r_label.bold = True
        r_label.font.size = Pt(10)
        r_label.font.color.rgb = RGBColor(30, 41, 59)
        r_val = p_desc.add_run(s_data['key_takeaway'])
        r_val.font.size = Pt(10)
        
        # Formula / Core Rule Box
        add_callout_box(
            doc,
            text_prefix="⚡ CÔNG THỨC & QUY TẮC CỐT LÕI:",
            title="",
            body_text=s_data['formula'],
            bg_hex="F1F5F9",
            border_hex="4F46E5"
        )
        
        # Memory trick
        if s_data['memory_trick']:
            p_mem = doc.add_paragraph()
            p_mem.paragraph_format.space_before = Pt(2)
            p_mem.paragraph_format.space_after = Pt(4)
            r_mem = p_mem.add_run(s_data['memory_trick'])
            r_mem.font.size = Pt(9.5)
            r_mem.font.italic = True
            r_mem.font.color.rgb = RGBColor(5, 150, 105) # Green
            
        # Quiz Box if any
        if s_data['quiz']:
            q = s_data['quiz']
            quiz_text = f"Câu hỏi: {q['question']}\nLựa chọn: {q['options']}\n{q['answer']}\n👉 Giải thích: {q['explanation']}"
            add_callout_box(
                doc,
                text_prefix="❓ CÂU HỎI TRẮC NGHIỆM & ĐÁP ÁN:",
                title="",
                body_text=quiz_text,
                bg_hex="FEF3C7", # Amber tint
                border_hex="D97706"
            )
            
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        
    # Final Reference Master Matrix Table (from Slide 29)
    doc.add_page_break()
    mat_h = doc.add_heading("III. BẢNG TỔNG HỢP TOÀN DIỆN 8 GIA ĐÌNH ĐẠI TỪ (MASTER MATRIX)", level=1)
    mat_h.paragraph_format.space_before = Pt(14)
    mat_h.paragraph_format.space_after = Pt(8)
    
    matrix_table = doc.add_table(rows=9, cols=3)
    matrix_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    matrix_table.autofit = False
    m_widths = [Inches(1.8), Inches(2.2), Inches(2.7)]
    m_headers = ["Gia Đình Đại Từ", "Các Đại Từ Đại Diện", "Vị Trí Cú Pháp & Quy Tắc Vàng"]
    
    for i, h_text in enumerate(m_headers):
        c = matrix_table.cell(0, i)
        c.width = m_widths[i]
        set_cell_background(c, "0F172A")
        set_cell_margins(c, top=100, bottom=100, left=120, right=120)
        p = c.paragraphs[0]
        r = p.add_run(h_text)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    m_rows = [
        ("1. Personal (Subject)", "I, you, he, she, it, we, they", "Đứng TRƯỚC động từ chính: [ S + Verb ]"),
        ("2. Personal (Object)", "me, you, him, her, it, us, them", "Đứng SAU động từ hoặc giới từ: [ Verb/Prep + O ]"),
        ("3. Possessive Pronoun", "mine, yours, his, hers, ours, theirs", "Đứng ĐỘC LẬP cuối câu: ĐTSH = TTSH + Noun"),
        ("4. Reflexive Pronoun", "myself, yourself, himself, herself, itself, ourselves, themselves", "3 vai: (1) S=O phản chiếu; (2) Nhấn mạnh; (3) by myself = alone"),
        ("5. Demonstrative", "this, that, these, those", "Gần/Xa & Số lượng. Đặc biệt: This/That đi với N không đếm được"),
        ("6. Indefinite Pronoun", "everyone, someone, anyone, nobody, everything, nothing...", "QUY TẮC BẮT BUỘC: Luôn chia ĐỘNG TỪ SỐ ÍT (is, has, Vs/es)"),
        ("7. Interrogative", "who, whom, whose, what, which", "Đặt câu hỏi chính xác: Who (chủ ngữ), Whom (tân ngữ), Whose (sở hữu)"),
        ("8. Relative & Compound", "who, whom, which, that, whose, whoever, whichever", "Nối mệnh đề TOEIC: Cấm THAT sau dấu phẩy & giới từ. Whoever = Anyone who")
    ]
    
    for r_idx, (col1, col2, col3) in enumerate(m_rows, start=1):
        bg = "F1F5F9" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate([col1, col2, col3]):
            c = matrix_table.cell(r_idx, c_idx)
            c.width = m_widths[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(val)
            r.font.size = Pt(9)
            if c_idx == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
                
    output_docx = os.path.abspath("Tom_Tat_32_Slide_Dai_Tu_Tieng_Anh.docx")
    doc.save(output_docx)
    print(f"✅ Successfully created Word document: {output_docx}")
    return output_docx

def convert_docx_to_pdf(docx_path):
    import win32com.client
    pdf_path = docx_path.replace(".docx", ".pdf")
    
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17) # 17 = wdFormatPDF
        doc.Close(SaveChanges=False)
        print(f"✅ Successfully exported PDF document: {pdf_path}")
    except Exception as e:
        print(f"Error converting to PDF via win32com: {e}")
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass
                
    if os.path.exists(pdf_path):
        print(f"PDF exists, size: {os.path.getsize(pdf_path)} bytes")
        return pdf_path
    else:
        print("PDF conversion failed or file not created.")
        return None

if __name__ == "__main__":
    docx_file = build_word_document()
    pdf_file = convert_docx_to_pdf(docx_file)
    print(f"FINISHED: DOCX={docx_file}, PDF={pdf_file}")
