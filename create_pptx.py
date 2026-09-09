"""
Script to generate English_Pronouns_Masterclass.pptx using python-pptx
- All slide text is in ENGLISH
- All presenter notes are in VIETNAMESE
- Professional 16:9 widescreen layout with color palettes and tables
"""

import sys
import os

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
except ImportError:
    print("python-pptx not found. Installing python-pptx...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx"])
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Color Palette (Deep Navy Modern Dark Theme)
    BG_COLOR = RGBColor(15, 23, 42)        # #0f172a
    CARD_BG = RGBColor(30, 41, 59)         # #1e293b
    CARD_BORDER = RGBColor(51, 65, 85)     # #334155
    WHITE = RGBColor(255, 255, 255)
    MUTED_TEXT = RGBColor(148, 163, 184)   # #94a3b8
    PRIMARY = RGBColor(129, 140, 248)      # #818cf8
    CYAN = RGBColor(56, 189, 248)          # #38bdf8
    GREEN = RGBColor(52, 211, 153)         # #34d399
    AMBER = RGBColor(251, 191, 36)         # #fbbf24
    RED = RGBColor(248, 113, 113)          # #f87171

    slides_data = [
        # Slide 1: Cover
        {
            "category": "ENGLISH GRAMMAR MASTERCLASS",
            "title": "PRONOUNS IN ENGLISH",
            "subtitle": "How Small Words Create Fluent, Natural English",
            "body": [
                ("Mastering the 5 Essential Pronoun Families", PRIMARY, True, 22),
                ("Subject vs Object • Possessive • Reflexive • Demonstrative • Interrogative", WHITE, False, 16),
                ("Includes Interactive Classroom Quizzes & Real-time Exercises", CYAN, False, 15)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 1 - 2 phút):\n"
                "- Chào lớp: 'Xin chào các bạn! Hôm nay chúng ta sẽ cùng tìm hiểu một chủ đề cực kỳ quen thuộc nhưng lại rất dễ nhầm lẫn trong tiếng Anh: ĐẠI TỪ (PRONOUNS).'\n"
                "- Đặt vấn đề: 'Đại từ tuy là những từ rất ngắn như He, She, It, Mine, This... nhưng nếu dùng sai thì câu văn sẽ rất ngô nghê.'\n"
                "- Hôm nay chúng ta sẽ vừa học mẹo nhớ, vừa làm quiz trắc nghiệm tương tác ngay tại lớp!"
            )
        },

        # Slide 2: Icebreaker
        {
            "category": "01. INTRODUCTION",
            "title": "THE PROBLEM OF REPETITION",
            "subtitle": "Listen to this story and feel the difference",
            "body": [
                ("Unnatural Story (Repetitive):", RED, True, 18),
                ("“Tom is a student. Tom likes football. Tom plays football every day. Tom lives near my house.”", MUTED_TEXT, False, 16),
                ("Classroom Question: Does repeating 'Tom' sound awkward?", AMBER, True, 18),
                ("The Natural Solution:", GREEN, True, 18),
                ("“Tom is a student. He likes football. He plays football every day.”", WHITE, True, 17),
                ("➔ The word 'HE' smoothly replaces the noun 'Tom'.", CYAN, False, 15)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 2 - 2 phút):\n"
                "- Đọc to đoạn văn với giọng nhấn mạnh từng từ Tom: 'Tom is a student. Tom likes football. Tom plays football every day...'\n"
                "- Hỏi cả lớp: 'Các bạn nghe có thấy từ Tom bị lặp quá nhiều không? Nghe có mệt tai không?'\n"
                "- Chờ lớp phản ứng, sau đó bấm tiếp:\n"
                "- 'Đó là lý do ta thay Tom bằng HE. Đại từ sinh ra là để cứu vãn sự lặp từ!'"
            )
        },

        # Slide 3: Definition & Benefits
        {
            "category": "01. INTRODUCTION",
            "title": "WHAT IS A PRONOUN & WHY NEED THEM?",
            "subtitle": "Core definition and four golden advantages",
            "body": [
                ("DEFINITION:", PRIMARY, True, 18),
                ("A Pronoun is a word used in place of a noun or noun phrase.", WHITE, False, 16),
                ("• Tom ➔ He  |  Anna ➔ She  |  The Book ➔ It  |  Tom & Anna ➔ They", CYAN, False, 15),
                ("WHY DO WE NEED PRONOUNS? (4 Golden Reasons)", AMBER, True, 18),
                ("1. Avoid word repetition and robotic phrasing", WHITE, False, 15),
                ("2. Make sentences shorter & punchier", WHITE, False, 15),
                ("3. Create natural, fluent conversation flow", WHITE, False, 15),
                ("4. Clearly identify the subject or object referent", WHITE, False, 15)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 3 - 2 phút):\n"
                "- Giải thích: 'Pro' nghĩa là thay cho, 'Noun' là danh từ. Đại từ là từ thay thế cho danh từ.\n"
                "- Điểm qua 4 lợi ích: Tránh lặp từ, ngắn gọn, tự nhiên, và rõ ràng.\n"
                "- Lấy ví dụ nhanh với học sinh trong lớp để tạo không khí."
            )
        },

        # Slide 4: Roadmap
        {
            "category": "02. THE BIG PICTURE",
            "title": "5 ESSENTIAL PRONOUN FAMILIES",
            "subtitle": "The roadmap of our masterclass today",
            "body": [
                ("1. Personal Pronouns (Đại từ nhân xưng):", PRIMARY, True, 16),
                ("   • Subject (I, you, he, she, we, they) & Object (me, you, him, her, us, them)", WHITE, False, 14),
                ("2. Possessive Pronouns (Đại từ sở hữu):", CYAN, True, 16),
                ("   • Express ownership: mine, yours, his, hers, ours, theirs", WHITE, False, 14),
                ("3. Reflexive Pronouns (Đại từ phản thân):", GREEN, True, 16),
                ("   • Action reflects to self: myself, yourself, himself, themselves", WHITE, False, 14),
                ("4. Demonstrative Pronouns (Đại từ chỉ định):", AMBER, True, 16),
                ("   • Distance & Quantity: this, that, these, those", WHITE, False, 14),
                ("5. Interrogative Pronouns (Đại từ nghi vấn):", RED, True, 16),
                ("   • Asking questions: who, whom, whose, what, which", WHITE, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 4 - 1.5 phút):\n"
                "- Giới thiệu 5 nhóm đại từ sẽ học.\n"
                "- 'Tiếng Anh có rất nhiều đại từ, nhưng chỉ cần làm chủ 5 nhóm này là các bạn đã tự tin giao tiếp và xử lý 90% bài thi ngữ pháp rồi!'"
            )
        },

        # Slide 5: Subject Pronouns
        {
            "category": "03. PERSONAL PRONOUNS",
            "title": "SUBJECT PRONOUNS (ĐẠI TỪ CHỦ NGỮ)",
            "subtitle": "The doer of the action in the sentence",
            "body": [
                ("RULE TO REMEMBER: Subject Pronoun + Verb", AMBER, True, 18),
                ("List of Subject Pronouns:", PRIMARY, True, 16),
                ("• I (Tôi) | You (Bạn/Các bạn) | He (Anh ấy) | She (Cô ấy)", WHITE, False, 15),
                ("• It (Nó/Đồ vật) | We (Chúng tôi/ta) | They (Họ/Chúng nó)", WHITE, False, 15),
                ("Examples:", CYAN, True, 16),
                ("• I am a student.  |  She is my friend.", WHITE, False, 15),
                ("• He plays football. (He = Subject doing the action 'plays')", GREEN, False, 15),
                ("• They play football. (They = Subject doing the action)", GREEN, False, 15)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 5 - 2 phút):\n"
                "- Giải thích: Chủ ngữ là người/vật LÀM hành động, đứng trước động từ.\n"
                "- Công thức: Subject Pronoun + Verb.\n"
                "- Lưu ý: 'You' vừa là số ít (bạn), vừa là số nhiều (các bạn)."
            )
        },

        # Slide 6: Object Pronouns
        {
            "category": "03. PERSONAL PRONOUNS",
            "title": "OBJECT PRONOUNS (ĐẠI TỪ TÂN NGỮ)",
            "subtitle": "The receiver of the action",
            "body": [
                ("Subject ➔ Object Transformation:", PRIMARY, True, 16),
                ("• I ➔ me  |  You ➔ you  |  He ➔ him  |  She ➔ her", WHITE, False, 15),
                ("• It ➔ it  |  We ➔ us   |  They ➔ them", WHITE, False, 15),
                ("GOLDEN RULE:", AMBER, True, 18),
                ("Subject = Person/Thing doing action  |  Object = Person/Thing receiving action", WHITE, False, 15),
                ("Clear Comparison: “She likes him.”", GREEN, True, 17),
                ("• 'She' is the doer of liking ➔ Subject Pronoun", CYAN, False, 14),
                ("• 'him' is the person being liked ➔ Object Pronoun", CYAN, False, 14),
                ("More: She likes me • I know him • They help us", WHITE, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 6 - 2 phút):\n"
                "- Nhấn mạnh sự khác nhau: Chủ ngữ làm, Tân ngữ nhận.\n"
                "- Ví dụ câu: 'She likes him'. Tại sao không nói 'She likes he'? Vì anh ấy là người được thích (nhận tác động), nên phải dùng him.\n"
                "- Chuẩn bị cho lớp làm câu hỏi trắc nghiệm!"
            )
        },

        # Slide 7: Quiz 1
        {
            "category": "PRACTICE QUIZ",
            "title": "QUESTION 1: SUBJECT OR OBJECT?",
            "subtitle": "Which pronoun should replace 'Tom'?",
            "body": [
                ("Tom is my friend. ___ is very nice.", WHITE, True, 20),
                ("A. He", GREEN, True, 17),
                ("B. Him", WHITE, False, 16),
                ("C. His", WHITE, False, 16),
                ("D. Himself", WHITE, False, 16),
                ("CORRECT ANSWER: A. He", GREEN, True, 18),
                ("Explanation: Tom is the subject of the sentence performing/being described before the verb 'is'. Therefore, Tom = He.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 7 - 1.5 phút):\n"
                "- Đọc câu hỏi: 'Tom is my friend. ___ is very nice.'\n"
                "- Cho học sinh 5 giây giơ tay hoặc hô đáp án A, B, C, D.\n"
                "- Bấm hiện đáp án A: He.\n"
                "- Giải thích: Đứng ở vị trí chủ ngữ trước 'is' nên dùng He."
            )
        },

        # Slide 8: Quiz 2
        {
            "category": "PRACTICE QUIZ",
            "title": "QUESTION 2: REPLACING THE OBJECT",
            "subtitle": "Anna likes Peter. If we replace 'Peter', which is correct?",
            "body": [
                ("Sentence: “Anna likes Peter.”", WHITE, True, 20),
                ("A. Anna likes he.", WHITE, False, 16),
                ("B. Anna likes him.", GREEN, True, 17),
                ("C. Anna likes his.", WHITE, False, 16),
                ("D. Anna likes himself.", WHITE, False, 16),
                ("CORRECT ANSWER: B. Anna likes him.", GREEN, True, 18),
                ("Explanation: Peter is the person who receives Anna's affection (Object of the verb 'likes'). Therefore, Peter = him.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 8 - 1.5 phút):\n"
                "- Đọc câu: 'Anna likes Peter. Thay thế Peter bằng từ nào?'\n"
                "- Nhiều bạn sẽ nhầm sang he. Nhấn mạnh: Peter đứng sau động từ likes, là đối tượng nhận tác động ➔ Chọn B. him."
            )
        },

        # Slide 9: Possessive Pronouns
        {
            "category": "04. POSSESSIVE PRONOUNS",
            "title": "POSSESSIVE PRONOUNS (ĐẠI TỪ SỞ HỮU)",
            "subtitle": "Declaring: “Whose is this?” (Cái này thuộc về ai?)",
            "body": [
                ("The Possessive Pronouns List:", PRIMARY, True, 16),
                ("• I ➔ mine (của tôi)  |  You ➔ yours (của bạn)", WHITE, False, 15),
                ("• He ➔ his (của anh ấy) | She ➔ hers (của cô ấy)", WHITE, False, 15),
                ("• It ➔ its (của nó)     | We ➔ ours (của chúng tôi)", WHITE, False, 15),
                ("• They ➔ theirs (của họ)", WHITE, False, 15),
                ("Classic Example:", AMBER, True, 17),
                ("“This book is mine.” = Quyển sách này là của tôi.", GREEN, True, 16),
                ("Notice: 'mine' stands at the end without needing any noun after it.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 9 - 2 phút):\n"
                "- Chuyển sang nhóm 2: Đại từ sở hữu.\n"
                "- Dùng khi muốn nói: Cái này là CỦA AI. Ví dụ: This book is mine (Quyển sách này là của tôi).\n"
                "- Đọc qua bảng: mine, yours, his, hers, ours, theirs."
            )
        },

        # Slide 10: Possessive Adjective vs Pronoun
        {
            "category": "04. POSSESSIVE PRONOUNS",
            "title": "POSSESSIVE ADJECTIVE VS. PRONOUN",
            "subtitle": "The #1 most common confusion in English tests",
            "body": [
                ("POSSESSIVE ADJECTIVE (Tính từ sở hữu):", AMBER, True, 17),
                ("my, your, his, her, its, our, their", WHITE, False, 15),
                ("➔ MUST be followed by a NOUN: [my + noun]", RED, True, 15),
                ("Example: “This is my book.”", WHITE, False, 14),
                ("POSSESSIVE PRONOUN (Đại từ sở hữu):", PRIMARY, True, 17),
                ("mine, yours, his, hers, ours, theirs", WHITE, False, 15),
                ("➔ STANDS ALONE without any noun: [mine]", GREEN, True, 15),
                ("Example: “This book is mine.”", WHITE, False, 14),
                ("Memory Trick: Có danh từ theo sau ➔ my book. Đứng một mình lẻ loi ➔ mine.", CYAN, True, 15)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 10 - 2 phút):\n"
                "- Nhấn mạnh: Đây là phần dễ nhầm nhất!\n"
                "- Mẹo nhớ cực đơn giản:\n"
                "  + Nếu phía sau CÓ danh từ: dùng tính từ sở hữu (my car, your house).\n"
                "  + Nếu đứng MỘT MÌNH không có danh từ: dùng đại từ sở hữu (The car is mine, The house is yours)."
            )
        },

        # Slide 11: Quiz 3
        {
            "category": "PRACTICE QUIZ",
            "title": "QUESTION 3: NOUN ATTACHED?",
            "subtitle": "Choose the correct answer",
            "body": [
                ("Sentence: “This is ___ book.”", WHITE, True, 20),
                ("A. mine", WHITE, False, 16),
                ("B. my", GREEN, True, 17),
                ("C. me", WHITE, False, 16),
                ("D. I", WHITE, False, 16),
                ("CORRECT ANSWER: B. my", GREEN, True, 18),
                ("Explanation: The noun 'book' follows immediately after the blank. According to the formula [my + noun], we must choose 'my'.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 11 - 1.5 phút):\n"
                "- Hỏi lớp: 'Nhìn phía sau có từ book là danh từ. Vậy chọn mine hay my?'\n"
                "- Chốt đáp án B: my book."
            )
        },

        # Slide 12: Quiz 4
        {
            "category": "PRACTICE QUIZ",
            "title": "QUESTION 4: STANDING ALONE?",
            "subtitle": "Choose the correct answer",
            "body": [
                ("Sentence: “This book is ___.”", WHITE, True, 20),
                ("A. my", WHITE, False, 16),
                ("B. me", WHITE, False, 16),
                ("C. mine", GREEN, True, 17),
                ("D. I", WHITE, False, 16),
                ("CORRECT ANSWER: C. mine", GREEN, True, 18),
                ("Explanation: After 'is' and before the period, there is NO noun. Therefore, we use the standalone possessive pronoun: 'mine'.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 12 - 1.5 phút):\n"
                "- Hỏi lớp: 'Câu này sau chỗ trống không còn từ nào cả, đứng một mình ở cuối câu. Ta chọn gì?'\n"
                "- Chốt đáp án C: mine."
            )
        },

        # Slide 13: Reflexive Pronouns
        {
            "category": "05. REFLEXIVE PRONOUNS",
            "title": "REFLEXIVE PRONOUNS (ĐẠI TỪ PHẢN THÂN)",
            "subtitle": "When doer and receiver are the EXACT same person",
            "body": [
                ("CORE CONCEPT: Action reflects back to oneself", PRIMARY, True, 17),
                ("“Tôi làm gì đó với chính bản thân tôi.”", AMBER, True, 16),
                ("Singular (-self): myself • yourself • himself • herself • itself", WHITE, False, 15),
                ("Plural (-selves): ourselves • yourselves • themselves", WHITE, False, 15),
                ("Real-Life Examples:", GREEN, True, 16),
                ("• I hurt myself. (Tôi tự làm đau chính mình)", CYAN, False, 15),
                ("• He taught himself English. (Anh ấy tự học tiếng Anh)", CYAN, False, 15),
                ("• She looked at herself in the mirror. (Cô ấy nhìn chính mình trong gương)", CYAN, False, 15)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 13 - 2 phút):\n"
                "- Nhóm 3: Đại từ phản thân (Reflexive Pronouns).\n"
                "- Giống như soi gương: Mình làm hành động và chính mình nhận tác động.\n"
                "- Nhớ đuôi: 1 người thì đuôi -self, nhiều người thì đuôi -selves."
            )
        },

        # Slide 14: Quiz 5
        {
            "category": "PRACTICE QUIZ",
            "title": "QUESTION 5: COOKING ACCIDENT",
            "subtitle": "Choose the correct reflexive pronoun",
            "body": [
                ("Sentence: “John cut ___ while cooking.”", WHITE, True, 20),
                ("A. him", WHITE, False, 16),
                ("B. his", WHITE, False, 16),
                ("C. himself", GREEN, True, 17),
                ("D. he", WHITE, False, 16),
                ("CORRECT ANSWER: C. himself", GREEN, True, 18),
                ("Explanation: John did the cutting, and John received the cut. Doer and receiver are identical ➔ John = himself.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 14 - 1.5 phút):\n"
                "- John nấu ăn bị đứt tay. John tự cắt vào tay mình.\n"
                "- Người làm và người bị là 1 ➔ Chọn C: himself."
            )
        },

        # Slide 15: Demonstrative Pronouns
        {
            "category": "06. DEMONSTRATIVE PRONOUNS",
            "title": "DEMONSTRATIVE PRONOUNS (ĐẠI TỪ CHỈ ĐỊNH)",
            "subtitle": "Two golden axes: Distance (Near/Far) & Quantity (1/Many)",
            "body": [
                ("THE 2x2 MATRIX:", PRIMARY, True, 18),
                ("• NEAR (Gần) + Singular (1 cái) ➔ THIS  (“This is my phone”)", CYAN, True, 15),
                ("• FAR (Xa)   + Singular (1 cái) ➔ THAT  (“That is my car”)", AMBER, True, 15),
                ("• NEAR (Gần) + Plural (Nhiều)  ➔ THESE (“These are my books”)", CYAN, True, 15),
                ("• FAR (Xa)   + Plural (Nhiều)  ➔ THOSE (“Those are my shoes”)", AMBER, True, 15),
                ("QUICK MEMORY FORMULA:", GREEN, True, 17),
                ("This / These = GẦN  |  That / Those = XA", WHITE, False, 15),
                ("This / That = 1 CÁI |  These / Those = NHIỀU CÁI", WHITE, False, 15)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 15 - 2 phút):\n"
                "- Nhóm 4: Đại từ chỉ định (This, That, These, Those).\n"
                "- Vẽ ra 2 trục trong đầu:\n"
                "  + Gần hay xa? Gần thì This/These, xa thì That/Those.\n"
                "  + 1 hay nhiều? 1 thì This/That, nhiều thì These/Those."
            )
        },

        # Slide 16: Quiz 6
        {
            "category": "PRACTICE QUIZ",
            "title": "QUESTION 6: BOOK IN HAND",
            "subtitle": "You are holding a single book in your hand",
            "body": [
                ("Situation: You are holding a book in your hand.", WHITE, False, 16),
                ("You say: “___ is my book.”", WHITE, True, 20),
                ("A. Those", WHITE, False, 16),
                ("B. These", WHITE, False, 16),
                ("C. This", GREEN, True, 17),
                ("D. That", WHITE, False, 16),
                ("CORRECT ANSWER: C. This", GREEN, True, 18),
                ("Explanation: The book is in your hand (NEAR) and it is 1 book (SINGULAR) ➔ We use 'This'.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 16 - 1.5 phút):\n"
                "- Bạn đang cầm 1 quyển sách ngay trên tay: Vừa gần, vừa 1 quyển.\n"
                "- Chốt đáp án C: This is my book."
            )
        },

        # Slide 17: Quiz 7
        {
            "category": "PRACTICE QUIZ",
            "title": "QUESTION 7: CARS IN THE DISTANCE",
            "subtitle": "You see multiple cars parked far away",
            "body": [
                ("Situation: Looking at several cars parked down the street.", WHITE, False, 16),
                ("You say: “___ are cars.”", WHITE, True, 20),
                ("A. This", WHITE, False, 16),
                ("B. That", WHITE, False, 16),
                ("C. These", WHITE, False, 16),
                ("D. Those", GREEN, True, 17),
                ("CORRECT ANSWER: D. Those", GREEN, True, 18),
                ("Explanation: Multiple cars (PLURAL) + located in the distance (FAR) ➔ We use 'Those'.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 17 - 1.5 phút):\n"
                "- Nhìn xe đằng xa: Vừa xa, vừa số nhiều (are cars).\n"
                "- Chốt đáp án D: Those are cars."
            )
        },

        # Slide 18: Interrogative Pronouns
        {
            "category": "07. INTERROGATIVE PRONOUNS",
            "title": "INTERROGATIVE PRONOUNS (ĐẠI TỪ NGHI VẤN)",
            "subtitle": "Key words used to ask questions",
            "body": [
                ("WHO (Ai) — Asking about a person as subject", CYAN, True, 16),
                ("• Example: “Who is she?” (Cô ấy là ai?)", WHITE, False, 14),
                ("WHOM (Ai) — Asking about a person as object (formal)", CYAN, True, 16),
                ("• Example: “Whom did you see?” (Bạn đã gặp ai?)", WHITE, False, 14),
                ("WHOSE (Của ai) — Asking about ownership", AMBER, True, 16),
                ("• Example: “Whose book is this?” (Quyển sách này của ai?)", WHITE, False, 14),
                ("WHAT (Cái gì/Điều gì) — Asking about things/events", GREEN, True, 16),
                ("• Example: “What is this?” (Đây là cái gì?)", WHITE, False, 14),
                ("WHICH (Cái nào/Người nào) — Choosing among specific options", PRIMARY, True, 16),
                ("• Example: “Which one do you prefer?” (Bạn thích cái nào hơn?)", WHITE, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 18 - 2 phút):\n"
                "- Nhóm 5: Đại từ nghi vấn dùng để hỏi.\n"
                "- Who: hỏi người (chủ ngữ).\n"
                "- Whom: hỏi người (tân ngữ, trang trọng).\n"
                "- Whose: hỏi của ai (sở hữu).\n"
                "- What: hỏi cái gì.\n"
                "- Which: hỏi cái nào (khi có sự lựa chọn giữa các vật cụ thể)."
            )
        },

        # Slide 19: Quiz 8
        {
            "category": "PRACTICE QUIZ",
            "title": "QUESTION 8: ASKING ABOUT A PERSON",
            "subtitle": "Choose the correct interrogative pronoun",
            "body": [
                ("Sentence: “___ is your best friend?”", WHITE, True, 20),
                ("A. What", WHITE, False, 16),
                ("B. Who", GREEN, True, 17),
                ("C. Whose", WHITE, False, 16),
                ("D. Which", WHITE, False, 16),
                ("CORRECT ANSWER: B. Who", GREEN, True, 18),
                ("Explanation: The question asks about 'best friend' (a person as the subject) ➔ We use 'Who'.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 19 - 1.5 phút):\n"
                "- Hỏi người bạn thân nhất của bạn là ai? Người thì chọn B: Who."
            )
        },

        # Slide 20: Quiz 9
        {
            "category": "PRACTICE QUIZ",
            "title": "QUESTION 9: ASKING ABOUT OWNERSHIP",
            "subtitle": "Choose the correct interrogative pronoun",
            "body": [
                ("Sentence: “___ phone is this?”", WHITE, True, 20),
                ("A. Who", WHITE, False, 16),
                ("B. What", WHITE, False, 16),
                ("C. Whose", GREEN, True, 17),
                ("D. Which", WHITE, False, 16),
                ("CORRECT ANSWER: C. Whose", GREEN, True, 18),
                ("Explanation: The question wants to find the owner of the phone (Whose = Của ai) ➔ We use 'Whose'.", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 20 - 1.5 phút):\n"
                "- Điện thoại này là của ai? Hỏi sở hữu ➔ Chọn C: Whose."
            )
        },

        # Slide 21: 3 Golden Questions
        {
            "category": "08. STRATEGY & MASTERY",
            "title": "THE 3 GOLDEN QUESTIONS FOR ANY SENTENCE",
            "subtitle": "How to instantly choose the right pronoun in exams & conversations",
            "body": [
                ("QUESTION 1: Who performs the action? (Ai làm hành động?)", PRIMARY, True, 16),
                ("➔ Use SUBJECT PRONOUN (He plays football)", WHITE, False, 15),
                ("QUESTION 2: Who/What receives the action? (Ai/vật nhận tác động?)", CYAN, True, 16),
                ("➔ Use OBJECT PRONOUN (I like him)", WHITE, False, 15),
                ("QUESTION 3: Does action bounce back to the doer? (Có tác động ngược lại chính mình?)", GREEN, True, 16),
                ("➔ Use REFLEXIVE PRONOUN (He hurt himself)", WHITE, False, 15),
                ("Mastery Tip: Asking these 3 questions eliminates 100% of pronoun errors!", AMBER, True, 16)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 21 - 2 phút):\n"
                "- Tổng hợp tư duy: Khi gặp một câu, hãy tự hỏi 3 câu:\n"
                "  1. Ai làm? ➔ Subject (He).\n"
                "  2. Ai bị tác động? ➔ Object (him).\n"
                "  3. Tự làm với chính mình? ➔ Reflexive (himself)."
            )
        },

        # Slide 22: Summary Table
        {
            "category": "09. COMPLETE REFERENCE",
            "title": "ALL-IN-ONE PRONOUNS CHEATSHEET",
            "subtitle": "The master table for quick revision",
            "body": [
                ("• Subject:       I, you, he, she, it, we, they (Doer of action)", PRIMARY, True, 15),
                ("• Object:        me, you, him, her, it, us, them (Receiver of action)", CYAN, True, 15),
                ("• Possessive:    mine, yours, his, hers, ours, theirs (Ownership - standalone)", AMBER, True, 15),
                ("• Reflexive:     myself, yourself, himself, themselves (Self-reflection)", GREEN, True, 15),
                ("• Demonstrative: this, that, these, those (Pointing near/far)", WHITE, True, 15),
                ("• Interrogative: who, whom, whose, what, which (Question words)", RED, True, 15)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 22 - 1.5 phút):\n"
                "- Bảng tóm tắt toàn bộ 6 nhóm đại từ.\n"
                "- Khuyên người học chụp lại màn hình làm tài liệu ôn tập."
            )
        },

        # Slide 23: Grand Challenge Quiz 10
        {
            "category": "GRAND CHALLENGE",
            "title": "QUESTION 10: SARAH'S NEW LAPTOP",
            "subtitle": "Combining two possessive rules in a single sentence!",
            "body": [
                ("Sarah has a new laptop. ___ laptop is very expensive, but the laptop is not ___.", WHITE, True, 18),
                ("A. Hers / her", WHITE, False, 16),
                ("B. Her / hers", GREEN, True, 17),
                ("C. She / her", WHITE, False, 16),
                ("D. Her / she", WHITE, False, 16),
                ("CORRECT ANSWER: B. Her / hers", GREEN, True, 18),
                ("Explanation: 1. 'Her laptop' ➔ Before noun 'laptop' = Possessive Adjective (Her).", CYAN, False, 14),
                ("2. 'not hers' ➔ Stands alone without a noun = Possessive Pronoun (hers).", CYAN, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 23 - 2 phút):\n"
                "- Câu hỏi tổng hợp đỉnh cao: Kết hợp cả tính từ sở hữu và đại từ sở hữu.\n"
                "- Vị trí 1 có laptop ➔ dùng Her.\n"
                "- Vị trí 2 đứng một mình ➔ dùng hers.\n"
                "- Chúc mừng các bạn chọn B: Her / hers!"
            )
        },

        # Slide 24: Mini-Game Speed Round
        {
            "category": "10. RAPID FIRE MINI-GAME",
            "title": "RAPID FIRE: GUESS THE PRONOUN!",
            "subtitle": "Speed test for the entire classroom",
            "body": [
                ("1. Tom is a student. [ ___ ] studies English.  ➔  HE", CYAN, True, 16),
                ("2. I like Anna. I often talk to [ ___ ].         ➔  HER", PRIMARY, True, 16),
                ("3. This is my pen. The pen is [ ___ ].          ➔  MINE", AMBER, True, 16),
                ("4. He made the cake by [ ___ ].                 ➔  HIMSELF", GREEN, True, 16),
                ("5. [ ___ ] are my shoes (pointing to feet).     ➔  THESE", RED, True, 16),
                ("Class activity: Split into 2 teams or raise hands rapidly!", WHITE, False, 14)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 24 - 3 phút):\n"
                "- Tổ chức mini-game cho cả lớp:\n"
                "  + Đọc từng câu và đếm 1, 2, 3 để lớp hô to đáp án.\n"
                "  + Câu 1: He, Câu 2: her, Câu 3: mine, Câu 4: himself, Câu 5: These.\n"
                "  + Có thể chia 2 đội xem đội nào trả lời nhanh hơn!"
            )
        },

        # Slide 25: Conclusion
        {
            "category": "11. FINAL TAKEAWAY",
            "title": "SMALL WORDS, MASSIVE IMPACT",
            "subtitle": "Master English by mastering the fundamentals",
            "body": [
                ("“Pronouns may be small words, but they are very important in English.”", AMBER, True, 20),
                ("Summary of 5 Golden Rules:", PRIMARY, True, 16),
                ("1. Pronoun = Replaces a noun to avoid repetition", WHITE, False, 15),
                ("2. Subject Pronoun (I, he, they) = Doer before verb", WHITE, False, 15),
                ("3. Object Pronoun (me, him, them) = Receiver after verb", WHITE, False, 15),
                ("4. Possessive Pronoun (mine, hers) = Standalone ownership", WHITE, False, 15),
                ("5. Reflexive Pronoun (-self/-selves) = Action reflects back to doer", WHITE, False, 15),
                ("THANK YOU FOR YOUR PARTICIPATION! Q&A TIME.", GREEN, True, 16)
            ],
            "notes": (
                "KỊCH BẢN THUYẾT TRÌNH (Slide 25 - 1.5 phút):\n"
                "- Kết thúc bài học bằng câu chốt: 'Pronouns may be small words, but they are very important in English.'\n"
                "- Cảm ơn cả lớp đã chú ý lắng nghe và tham gia rất nhiệt tình.\n"
                "- Mở rộng phần hỏi đáp (Q&A) nếu lớp có thắc mắc."
            )
        }
    ]

    # Create slides
    blank_layout = prs.slide_layouts[6] # Blank slide

    for idx, data in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)

        # 1. Background rectangle
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_COLOR
        bg.line.fill.background()

        # 2. Category badge & Slide Counter
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.4))
        tf_h = header_box.text_frame
        tf_h.word_wrap = True
        p_h = tf_h.paragraphs[0]
        p_h.text = f"{data['category']}   |   SLIDE {idx+1} OF {len(slides_data)}"
        p_h.font.size = Pt(11)
        p_h.font.bold = True
        p_h.font.color.rgb = CYAN

        # 3. Slide Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(0.8))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = data["title"]
        p_t.font.size = Pt(28)
        p_t.font.bold = True
        p_t.font.color.rgb = WHITE

        # 4. Slide Subtitle
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.55), Inches(11.7), Inches(0.4))
        tf_s = sub_box.text_frame
        tf_s.word_wrap = True
        p_s = tf_s.paragraphs[0]
        p_s.text = data["subtitle"]
        p_s.font.size = Pt(13)
        p_s.font.color.rgb = MUTED_TEXT

        # 5. Content Card Shape
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.1), Inches(11.73), Inches(4.8))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.5)

        # 6. Card Body Content
        body_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.3), Inches(11.0), Inches(4.4))
        tf_b = body_box.text_frame
        tf_b.word_wrap = True

        for i, (line_text, color, is_bold, size) in enumerate(data["body"]):
            p = tf_b.paragraphs[0] if i == 0 else tf_b.add_paragraph()
            p.text = line_text
            p.font.size = Pt(size)
            p.font.bold = is_bold
            p.font.color.rgb = color
            p.space_after = Pt(8)

        # 7. Add VIETNAMESE Speaker Notes to PowerPoint Presenter Pane
        notes_slide = slide.notes_slide
        tf_notes = notes_slide.notes_text_frame
        tf_notes.text = data["notes"]

    output_filename = "English_Pronouns_Masterclass.pptx"
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), output_filename)
    prs.save(output_path)
    print(f"✅ Successfully created PowerPoint presentation at: {output_path}")
    return output_path

if __name__ == "__main__":
    create_presentation()
