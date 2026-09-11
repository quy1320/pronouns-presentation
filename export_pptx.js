// Dedicated Executive-Grade PowerPoint (.pptx) Exporter
// Produces 16:9 widescreen slides with modern card layouts, embedded 3D illustrations, styled quiz cards, and verbatim Vietnamese speaker notes.

window.exportPPTX = function() {
  const btn = document.getElementById('exportPptxBtn');
  const originalText = btn ? btn.innerHTML : '';
  if (btn) {
    btn.innerHTML = `<span>Äang táº¡o PowerPoint 26 Slide cao cáº¥p...</span>`;
    btn.disabled = true;
  }

  try {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'WIDE169', width: 13.333, height: 7.5 });
    pptx.layout = 'WIDE169';
    pptx.author = 'English Masterclass Keynote';
    pptx.company = 'English Mastery';
    pptx.title = 'English Pronouns Keynote Presentation - 26 Slides';

    const assets = window.SLIDE_ASSETS || {};

    // Helper: Create slide with 100% full-bleed dark navy background
    function createSlide() {
      const slide = pptx.addSlide();
      slide.background = { fill: '0F172A' };
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 13.333, h: 7.5,
        fill: { color: '0F172A' },
        line: { color: '0F172A', width: 0 }
      });
      return slide;
    }

    // Helper: Add consistent executive header
    function addHeader(slide, category, title, subtitle) {
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 0.35, w: 3.8, h: 0.3,
        fill: { color: '1E293B' },
        line: { color: '38BDF8', width: 1 },
        rectRadius: 0.1
      });
      slide.addText(category.toUpperCase(), {
        x: 0.8, y: 0.35, w: 3.8, h: 0.3,
        fontSize: 9, fontFace: 'Calibri', bold: true, color: '38BDF8',
        align: 'center', valign: 'middle'
      });

      slide.addText(title, {
        x: 0.8, y: 0.75, w: 11.7, h: 0.55,
        fontSize: 22, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        valign: 'middle'
      });

      slide.addText(subtitle, {
        x: 0.8, y: 1.3, w: 11.7, h: 0.3,
        fontSize: 11, fontFace: 'Calibri', color: '94A3B8',
        valign: 'top'
      });
    }

    // Helper: Add clean Vietnamese speaker notes
    function attachNotes(slide, rawNotes) {
      const cleanNotes = (rawNotes || '')
        .replace(/<[^>]+>/g, ' ')
        .replace(/&quot;/g, '"')
        .replace(/&amp;/g, '&')
        .replace(/\s+/g, ' ')
        .trim();
      slide.addNotes(cleanNotes);
    }

    // Helper: Generic Quiz Slide Builder (2x2 Grid + Explanation Card)
    function buildQuizSlide(slideItem, qBadge, qText, optionsArr, correctOpt, explanation) {
      const slide = createSlide();
      addHeader(slide, slideItem.category, slideItem.title, slideItem.subtitle);

      // Question Box
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.7, w: 11.73, h: 1.1,
        fill: { color: '1E293B' },
        line: { color: '38BDF8', width: 1.5 },
        rectRadius: 0.12
      });
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: 1.9, w: 1.3, h: 0.32,
        fill: { color: '38BDF8' },
        line: { color: '38BDF8' },
        rectRadius: 0.08
      });
      slide.addText(qBadge, {
        x: 1.1, y: 1.9, w: 1.3, h: 0.32,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '0F172A',
        align: 'center', valign: 'middle'
      });
      slide.addText(qText, {
        x: 2.6, y: 1.85, w: 9.6, h: 0.8,
        fontSize: 14, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        valign: 'middle'
      });

      // 4 Options Grid (2x2)
      const optPositions = [
        { x: 0.8, y: 3.0 }, { x: 6.83, y: 3.0 },
        { x: 0.8, y: 4.1 }, { x: 6.83, y: 4.1 }
      ];

      optionsArr.forEach((optStr, idx) => {
        const isCorrect = optStr.startsWith(correctOpt + ".");
        const pos = optPositions[idx];

        slide.addShape(pptx.ShapeType.roundRect, {
          x: pos.x, y: pos.y, w: 5.7, h: 0.9,
          fill: { color: isCorrect ? '064E3B' : '1E293B' },
          line: { color: isCorrect ? '10B981' : '334155', width: isCorrect ? 2 : 1 },
          rectRadius: 0.1
        });

        // Letter pill
        const letter = optStr.substring(0, 1);
        const textOnly = optStr.substring(3);

        slide.addShape(pptx.ShapeType.roundRect, {
          x: pos.x + 0.25, y: pos.y + 0.22, w: 0.45, h: 0.45,
          fill: { color: isCorrect ? '10B981' : '334155' },
          line: { color: isCorrect ? '10B981' : '334155' },
          rectRadius: 0.08
        });
        slide.addText(letter, {
          x: pos.x + 0.25, y: pos.y + 0.22, w: 0.45, h: 0.45,
          fontSize: 12, fontFace: 'Arial', bold: true, color: 'FFFFFF',
          align: 'center', valign: 'middle'
        });

        slide.addText(textOnly + (isCorrect ? "  âœ“ [CORRECT]" : ""), {
          x: pos.x + 0.9, y: pos.y + 0.15, w: 4.6, h: 0.6,
          fontSize: 13, fontFace: 'Calibri', bold: isCorrect,
          color: isCorrect ? '34D399' : 'E2E8F0',
          valign: 'middle'
        });
      });

      // Explanation Box
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 5.25, w: 11.73, h: 1.55,
        fill: { color: '0F2338' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.12
      });
      slide.addText("EXPLANATION & LINGUISTIC RULE:", {
        x: 1.1, y: 5.35, w: 11.1, h: 0.3,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '34D399'
      });
      slide.addText(explanation, {
        x: 1.1, y: 5.7, w: 11.1, h: 0.95,
        fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 18
      });

      attachNotes(slide, slideItem.speakerNotes);
    }

    // --- SLIDE 1: COVER ---
    {
      const s1 = createSlide();
      s1.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 0.9, w: 6.0, h: 5.7,
        fill: { color: '1E293B' },
        line: { color: '6366F1', width: 1.5 },
        rectRadius: 0.15
      });
      s1.addShape(pptx.ShapeType.roundRect, {
        x: 1.2, y: 1.3, w: 0.7, h: 0.7,
        fill: { color: '6366F1' },
        line: { color: '6366F1' },
        rectRadius: 0.15
      });
      s1.addText("PRONOUNS IN ENGLISH", {
        x: 1.2, y: 2.2, w: 5.2, h: 1.2,
        fontSize: 32, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        lineSpacing: 38
      });
      s1.addText("How Small Words Drive Natural, Fluent Communication", {
        x: 1.2, y: 3.5, w: 5.2, h: 0.7,
        fontSize: 14, fontFace: 'Calibri', color: '94A3B8', lineSpacing: 20
      });
      s1.addShape(pptx.ShapeType.roundRect, {
        x: 1.2, y: 4.4, w: 5.2, h: 0.05,
        fill: { color: '334155' },
        line: { color: '334155' }
      });
      s1.addText("5 Core Modules â€¢ Strategic Syntax Rules â€¢ 10 Interactive Polls", {
        x: 1.2, y: 4.7, w: 5.2, h: 0.8,
        fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 18
      });
      s1.addShape(pptx.ShapeType.roundRect, {
        x: 1.2, y: 5.7, w: 2.8, h: 0.4,
        fill: { color: '064E3B' },
        line: { color: '10B981', width: 1 },
        rectRadius: 0.08
      });
      s1.addText("Masterclass Series â€¢ 26 Slides", {
        x: 1.2, y: 5.7, w: 2.8, h: 0.4,
        fontSize: 11, fontFace: 'Arial', bold: true, color: '34D399',
        align: 'center', valign: 'middle'
      });

      if (assets.hero_cover) {
        s1.addImage({ data: assets.hero_cover, x: 7.3, y: 0.9, w: 5.23, h: 5.7, round: true });
      }
      attachNotes(s1, slidesData[0].speakerNotes);
    }

    // --- SLIDE 2: THE PROBLEM WITH REPETITION ---
    {
      const s2 = createSlide();
      addHeader(s2, slidesData[1].category, slidesData[1].title, slidesData[1].subtitle);

      // Left Box: Problem Story
      s2.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 6.2, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: 'EF4444', width: 1.5 },
        rectRadius: 0.12
      });
      s2.addText("â€œIf I say: Tom is a student. Tom likes football. Tom plays football every day. Tom lives near my house...â€", {
        x: 1.1, y: 2.1, w: 5.6, h: 1.1,
        fontSize: 13.5, fontFace: 'Calibri', color: 'FEE2E2', lineSpacing: 22
      });

      s2.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: 3.35, w: 5.6, h: 0.45,
        fill: { color: '3B181C' },
        line: { color: 'F87171', width: 1 },
        rectRadius: 0.08
      });
      s2.addText("Observation: How does repeating 'Tom' affect the natural flow?", {
        x: 1.1, y: 3.35, w: 5.6, h: 0.45,
        fontSize: 10.5, fontFace: 'Calibri', bold: true, color: 'FCA5A5',
        align: 'center', valign: 'middle'
      });

      s2.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: 4.0, w: 5.6, h: 2.5,
        fill: { color: '064E3B' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.1
      });
      s2.addText("THE FLUENT ENGLISH DELIVERY:\nâ€œTom is a student. He likes football. He plays football every day.â€\n\nðŸ‘‰ The pronoun He replaces the noun Tom to make speech smooth.", {
        x: 1.3, y: 4.2, w: 5.2, h: 2.1,
        fontSize: 12, fontFace: 'Calibri', color: 'E0F2FE', lineSpacing: 18
      });

      // Right: Tom Football Image
      if (assets.tom_football) {
        s2.addImage({ data: assets.tom_football, x: 7.4, y: 1.8, w: 5.13, h: 5.0, round: true });
      }
      attachNotes(s2, slidesData[1].speakerNotes);
    }

    // --- SLIDE 3: CONCEPT DEFINITION & 4 ADVANTAGES ---
    {
      const s3 = createSlide();
      addHeader(s3, slidesData[2].category, slidesData[2].title, slidesData[2].subtitle);

      // Left Box: Definition
      s3.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '6366F1', width: 1.5 },
        rectRadius: 0.12
      });
      s3.addText("DEFINITION: PRONOUN = REPLACES A NOUN", {
        x: 1.1, y: 2.0, w: 5.1, h: 0.35,
        fontSize: 11, fontFace: 'Arial', bold: true, color: '818CF8'
      });
      s3.addText("A pronoun is a word used in place of a noun or noun phrase to refer to people, objects, or concepts.", {
        x: 1.1, y: 2.45, w: 5.1, h: 0.9,
        fontSize: 13, fontFace: 'Calibri', color: 'FFFFFF', lineSpacing: 20
      });

      s3.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: 3.5, w: 5.1, h: 3.0,
        fill: { color: '0F172A' },
        line: { color: '334155', width: 1 },
        rectRadius: 0.08
      });
      s3.addText("â€¢ Tom âž” he\nâ€¢ Anna âž” she\nâ€¢ The book âž” it\nâ€¢ Tom and Anna âž” they", {
        x: 1.4, y: 3.8, w: 4.5, h: 2.4,
        fontSize: 13, fontFace: 'Courier New', color: '67E8F9', lineSpacing: 28
      });

      // Right Box: 4 Strategic Benefits
      s3.addShape(pptx.ShapeType.roundRect, {
        x: 6.83, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: 'F59E0B', width: 1.5 },
        rectRadius: 0.12
      });
      s3.addText("WHY DO WE NEED PRONOUNS? (STRATEGIC BENEFITS)", {
        x: 7.1, y: 2.0, w: 5.1, h: 0.35,
        fontSize: 11, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s3.addText("01. Eliminate Word Repetition:\nKeeps speech clean, sharp, and executive.\n\n02. Concise Delivery:\nConveys high-impact ideas with fewer syllables.\n\n03. Natural Sentence Rhythm:\nMimics native cadence and speech tempo.\n\n04. Context Clarity:\nIdentifies referents with zero ambiguity.", {
        x: 7.1, y: 2.5, w: 5.1, h: 4.1,
        fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 18
      });

      attachNotes(s3, slidesData[2].speakerNotes);
    }

    // --- SLIDE 4: ROADMAP (5 CORE MODULES) ---
    {
      const s4 = createSlide();
      addHeader(s4, slidesData[3].category, slidesData[3].title, slidesData[3].subtitle);

      const modules = [
        { num: "PART 01", title: "Subject\nPronouns", words: "I, you, he, she, it, we, they", role: "Khá»Ÿi xÆ°á»›ng hÃ nh Ä‘á»™ng trÆ°á»›c Äá»™ng tá»« (S + V)", color: '6366F1' },
        { num: "PART 02", title: "Object\nPronouns", words: "me, you, him, her, it, us, them", role: "Tiáº¿p nháº­n tÃ¡c Ä‘á»™ng sau Verb & Giá»›i tá»«", color: '06B6D4' },
        { num: "PART 03", title: "Possessive\nAdjectives", words: "my, your, his, her, its, our, their", role: "TÃ­nh tá»« bá»• nghÄ©a; Báº®T BUá»˜C cÃ³ Noun Ä‘i sau", color: 'F59E0B' },
        { num: "PART 04", title: "Possessive\nPronouns", words: "mine, yours, his, hers, ours, theirs", role: "Äáº¡i tá»« Ä‘á»™c láº­p; Possessive Pronoun = Possessive Adjective + Noun", color: '10B981' },
        { num: "PART 05", title: "Reflexive &\nOverview", words: "myself, yourself... & 4 Families", role: "Trá»ng tÃ¢m Pháº£n thÃ¢n + Tá»•ng quan Ä‘áº¡i tá»« khÃ¡c", color: 'EC4899' }
      ];

      modules.forEach((m, idx) => {
        const xPos = 0.8 + (idx * 2.4);
        s4.addShape(pptx.ShapeType.roundRect, {
          x: xPos, y: 1.8, w: 2.2, h: 4.8,
          fill: { color: '1E293B' },
          line: { color: m.color, width: 2 },
          rectRadius: 0.12
        });
        s4.addShape(pptx.ShapeType.roundRect, {
          x: xPos + 0.3, y: 2.1, w: 1.6, h: 0.32,
          fill: { color: m.color },
          line: { color: m.color },
          rectRadius: 0.08
        });
        s4.addText(m.num, {
          x: xPos + 0.3, y: 2.1, w: 1.6, h: 0.32,
          fontSize: 10, fontFace: 'Arial', bold: true, color: 'FFFFFF',
          align: 'center', valign: 'middle'
        });
        s4.addText(m.title, {
          x: xPos + 0.15, y: 2.6, w: 1.9, h: 0.8,
          fontSize: 13, fontFace: 'Arial', bold: true, color: 'FFFFFF',
          align: 'center', lineSpacing: 18
        });
        s4.addText(m.words, {
          x: xPos + 0.15, y: 3.5, w: 1.9, h: 0.8,
          fontSize: 10, fontFace: 'Courier New', color: '38BDF8',
          align: 'center', lineSpacing: 14
        });
        s4.addShape(pptx.ShapeType.line, {
          x: xPos + 0.3, y: 4.4, w: 1.6, h: 0,
          line: { color: '334155', width: 1 }
        });
        s4.addText(m.role, {
          x: xPos + 0.15, y: 4.6, w: 1.9, h: 1.8,
          fontSize: 11, fontFace: 'Calibri', color: 'CBD5E1',
          align: 'center', lineSpacing: 16
        });
      });

      attachNotes(s4, slidesData[3].speakerNotes);
    }

    // --- SLIDE 5: PART 01 - SUBJECT PRONOUNS ---
    {
      const s5 = createSlide();
      addHeader(s5, slidesData[4].category, slidesData[4].title, slidesData[4].subtitle);

      // Left Table
      s5.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '6366F1', width: 1.5 },
        rectRadius: 0.12
      });
      s5.addText("PERSON / ROLE", { x: 1.1, y: 2.05, w: 3.0, h: 0.35, fontSize: 11, fontFace: 'Arial', bold: true, color: '94A3B8' });
      s5.addText("SUBJECT PRONOUN", { x: 4.3, y: 2.05, w: 2.0, h: 0.35, fontSize: 11, fontFace: 'Arial', bold: true, color: '818CF8' });

      const s5Rows = [
        { role: "1st Person Singular", pron: "I" },
        { role: "2nd Person (Sing./Plur.)", pron: "You" },
        { role: "3rd Person Male", pron: "He" },
        { role: "3rd Person Female", pron: "She" },
        { role: "3rd Person Inanimate", pron: "It" },
        { role: "1st Person Plural", pron: "We" },
        { role: "3rd Person Plural", pron: "They" }
      ];
      s5Rows.forEach((r, idx) => {
        const yPos = 2.5 + (idx * 0.58);
        s5.addText(r.role, { x: 1.1, y: yPos, w: 3.0, h: 0.35, fontSize: 11.5, fontFace: 'Calibri', color: 'FFFFFF' });
        s5.addText(r.pron, { x: 4.3, y: yPos, w: 2.0, h: 0.35, fontSize: 12, fontFace: 'Courier New', bold: true, color: '67E8F9' });
      });

      // Right Box: Rule & Examples
      s5.addShape(pptx.ShapeType.roundRect, {
        x: 6.83, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: 'F59E0B', width: 1.5 },
        rectRadius: 0.12
      });
      s5.addText("CORE POSITION RULE: [ Subject Pronoun ] + Verb", {
        x: 7.1, y: 2.1, w: 5.1, h: 0.4,
        fontSize: 12, fontFace: 'Courier New', bold: true, color: 'FBBF24'
      });
      s5.addText("Sentence Position: Initiates the clause, placed directly BEFORE the verb.\nðŸ‘‰ Example: I love you. ('I' = Subject initiating the action)", {
        x: 7.1, y: 2.7, w: 5.1, h: 1.2,
        fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 18
      });
      s5.addShape(pptx.ShapeType.roundRect, {
        x: 7.1, y: 4.2, w: 5.1, h: 2.3,
        fill: { color: '0F172A' },
        line: { color: '38BDF8', width: 1 },
        rectRadius: 0.08
      });
      s5.addText("EVERYDAY EXAMPLES:\nâ€¢ I am a student.\nâ€¢ She is my friend.\nâ€¢ He plays football.\nâ€¢ They play football.", {
        x: 7.3, y: 4.4, w: 4.7, h: 1.9,
        fontSize: 12.5, fontFace: 'Calibri', color: 'E0F2FE', lineSpacing: 22
      });

      attachNotes(s5, slidesData[4].speakerNotes);
    }

    // --- SLIDE 6: PART 01 PRACTICE - QUESTION 01 ---
    buildQuizSlide(slidesData[5], "CHECK 01",
      "â€œTom is a talented striker. ___ plays football every weekend with his local club.â€",
      ["A. Him", "B. He", "C. His", "D. Himself"], "B",
      "Äá»©ng trÆ°á»›c Ä‘á»™ng tá»« 'plays' lÃ m chá»§ ngá»¯ cá»§a cÃ¢u âž” báº¯t buá»™c chá»n Ä‘áº¡i tá»« chá»§ ngá»¯ 'He'. Him (tÃ¢n ngá»¯), His (sá»Ÿ há»¯u), Himself (pháº£n thÃ¢n) Ä‘á»u sai."
    );

    // --- SLIDE 7: PART 01 PRACTICE - QUESTION 02 ---
    buildQuizSlide(slidesData[6], "CHECK 02",
      "â€œAfter the conference ended, David and ___ submitted the project report to the executive director.â€",
      ["A. me", "B. I", "C. myself", "D. mine"], "B",
      "'David and I' cÃ¹ng lÃ m chá»§ ngá»¯ cho Ä‘á»™ng tá»« 'submitted'. Táº¡m bá» 'David and', ta cÃ³ 'I submitted' (Ä‘Ãºng), chá»© khÃ´ng thá»ƒ dÃ¹ng 'me submitted'."
    );

    // --- SLIDE 8: PART 02 - OBJECT PRONOUNS ---
    {
      const s8 = createSlide();
      addHeader(s8, slidesData[7].category, slidesData[7].title, slidesData[7].subtitle);

      s8.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 6.2, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '06B6D4', width: 1.5 },
        rectRadius: 0.12
      });
      s8.addText("GRAMMATICAL FUNCTION: RECEIVER OF ACTION", {
        x: 1.1, y: 2.0, w: 5.6, h: 0.3,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '22D3EE'
      });
      s8.addText("Object pronouns receive the direct action of a transitive verb OR immediately follow a preposition.", {
        x: 1.1, y: 2.4, w: 5.6, h: 0.8,
        fontSize: 13, fontFace: 'Calibri', color: 'E2E8F0', lineSpacing: 20
      });
      s8.addText("TWO GOLDEN POSITIONS:\n1. Direct/Indirect Object: After action verbs (invited them)\n2. Prepositional Object: After prepositions (between you and me)", {
        x: 1.1, y: 3.3, w: 5.6, h: 1.2,
        fontSize: 12, fontFace: 'Calibri', color: '6EE7B7', lineSpacing: 20
      });
      s8.addText("Complete Set: me, you, him, her, it, us, them", {
        x: 1.1, y: 4.6, w: 5.6, h: 0.5,
        fontSize: 12, fontFace: 'Courier New', color: 'CBD5E1'
      });
      s8.addText("Example: â€œAnna is friendly. Peter invited HER to lunch.â€", {
        x: 1.1, y: 5.3, w: 5.6, h: 0.7,
        fontSize: 11, fontFace: 'Calibri', italic: true, color: '67E8F9'
      });

      if (assets.anna_and_peter) {
        s8.addImage({ data: assets.anna_and_peter, x: 7.4, y: 1.8, w: 5.13, h: 5.0, round: true });
      }
      attachNotes(s8, slidesData[7].speakerNotes);
    }

    // --- SLIDE 9: PART 02 PRACTICE - QUESTION 03 ---
    buildQuizSlide(slidesData[8], "CHECK 03",
      "â€œWe met our new international partners yesterday and invited ___ to visit our head office.â€",
      ["A. they", "B. their", "C. them", "D. theirs"], "C",
      "Äá»©ng sau ngoáº¡i Ä‘á»™ng tá»« 'invited' lÃ m tÃ¢n ngá»¯ trá»±c tiáº¿p âž” báº¯t buá»™c chá»n Ä‘áº¡i tá»« tÃ¢n ngá»¯ 'them'. They (chá»§ ngá»¯), Their (Possessive Adjective), Theirs (Possessive Pronoun) Ä‘á»u khÃ´ng há»£p lá»‡."
    );

    // --- SLIDE 10: PART 02 PRACTICE - QUESTION 04 ---
    buildQuizSlide(slidesData[9], "CHECK 04",
      "â€œThis confidential agreement must strictly remain between the client and ___.â€",
      ["A. I", "B. me", "C. my", "D. mine"], "B",
      "'Between' lÃ  giá»›i tá»«. Sau giá»›i tá»« báº¯t buá»™c dÃ¹ng Ä‘áº¡i tá»« tÃ¢n ngá»¯ (Object Pronoun) âž” 'between the client and me'. KhÃ´ng thá»ƒ dÃ¹ng 'I'."
    );

    // --- SLIDE 11: PART 03 - POSSESSIVE ADJECTIVES (THEORY 1) ---
    {
      const s11 = createSlide();
      addHeader(s11, slidesData[10].category, slidesData[10].title, slidesData[10].subtitle);

      // Left Box: Rule & Examples
      s11.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: 'F59E0B', width: 1.5 },
        rectRadius: 0.12
      });
      s11.addText("CORE STRUCTURAL LAW: POSSESSIVE ADJECTIVE + NOUN", {
        x: 1.1, y: 2.0, w: 5.1, h: 0.3,
        fontSize: 10, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s11.addText("Possessive adjectives cannot stand alone. They function strictly as determiners that modify and establish ownership over a following noun.", {
        x: 1.1, y: 2.35, w: 5.1, h: 0.7,
        fontSize: 12, fontFace: 'Calibri', color: 'E2E8F0', lineSpacing: 18
      });
      s11.addText("FORMULA:  [ Possessive Adjective ] + [ MANDATORY NOUN ]", {
        x: 1.1, y: 3.15, w: 5.1, h: 0.35,
        fontSize: 11, fontFace: 'Courier New', bold: true, color: '6EE7B7'
      });
      s11.addText("Complete Set: my, your, his, her, its, our, their", {
        x: 1.1, y: 3.65, w: 5.1, h: 0.4,
        fontSize: 11.5, fontFace: 'Courier New', color: 'CBD5E1'
      });

      // Examples Box
      s11.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: 4.25, w: 5.1, h: 2.1,
        fill: { color: '0F172A' },
        line: { color: 'F59E0B', width: 1 },
        rectRadius: 0.08
      });
      s11.addText("VERIFIED EXAMPLES:", {
        x: 1.3, y: 4.4, w: 4.7, h: 0.3,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s11.addText("1. \"This is my book.\"\n    my modifies 'book' (Possessive Adjective + Noun)\n\n2. \"Her laptop is on the table.\"\n    Her modifies 'laptop' (Never stands alone)", {
        x: 1.3, y: 4.75, w: 4.7, h: 1.4,
        fontSize: 11, fontFace: 'Calibri', color: 'FDE68A', lineSpacing: 18
      });

      // Right Box: Its vs It's Trap
      s11.addShape(pptx.ShapeType.roundRect, {
        x: 6.83, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '06B6D4', width: 1.5 },
        rectRadius: 0.12
      });
      s11.addText("THE NOTORIOUS 'ITS' VS 'IT'S' TRAP", {
        x: 7.1, y: 2.0, w: 5.1, h: 0.3,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '22D3EE'
      });
      s11.addShape(pptx.ShapeType.roundRect, {
        x: 7.1, y: 2.5, w: 5.1, h: 1.6,
        fill: { color: '0C2D48' },
        line: { color: '38BDF8', width: 1 },
        rectRadius: 0.08
      });
      s11.addText("🔵 ITS (NO APOSTROPHE) = Tính từ sở hữu\nDùng để chỉ quyền sở hữu của đồ vật, sự vật, con vật:\n👉 “The company updated its security protocol.”", {
        x: 7.3, y: 2.6, w: 4.7, h: 1.4,
        fontSize: 11.5, fontFace: 'Calibri', color: 'E0F2FE', lineSpacing: 18
      });

      s11.addShape(pptx.ShapeType.roundRect, {
        x: 7.1, y: 4.4, w: 5.1, h: 1.6,
        fill: { color: '2A1215' },
        line: { color: 'EF4444', width: 1 },
        rectRadius: 0.08
      });
      s11.addText("🔴 IT'S (WITH APOSTROPHE) = Viết tắt của 'it is' / 'it has'\nLà cụm Chủ ngữ + Động từ, không phải tính từ sở hữu:\n👉 “It's important to double-check the figures.”", {
        x: 7.3, y: 4.5, w: 4.7, h: 1.4,
        fontSize: 11.5, fontFace: 'Calibri', color: 'FEE2E2', lineSpacing: 18
      });

      attachNotes(s11, slidesData[10].speakerNotes);
    }

    // --- SLIDE 12: PART 03 - THE 7 POSSESSIVE ADJECTIVES (THEORY 2 - BỔ SUNG) ---
    {
      const s12 = createSlide();
      addHeader(s12, slidesData[11].category, slidesData[11].title, slidesData[11].subtitle);

      // Left Box: Singular (5 rows)
      s12.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: 'F59E0B', width: 1.5 },
        rectRadius: 0.12
      });
      s12.addText("SINGULAR FORMS (DẠNG SỐ ÍT)", {
        x: 1.1, y: 2.0, w: 5.1, h: 0.3,
        fontSize: 10, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });

      const singularPairs = [
        { pro: "I", adj: "MY", eg: "my phone", vn: "điện thoại của tôi" },
        { pro: "YOU", adj: "YOUR", eg: "your idea", vn: "ý tưởng của bạn" },
        { pro: "HE", adj: "HIS", eg: "his car", vn: "xe của anh ấy" },
        { pro: "SHE", adj: "HER", eg: "her bag", vn: "túi của cô ấy" },
        { pro: "IT", adj: "ITS", eg: "its tail", vn: "cái đuôi của nó" }
      ];

      singularPairs.forEach((item, idx) => {
        const yPos = 2.45 + (idx * 0.82);
        s12.addShape(pptx.ShapeType.roundRect, {
          x: 1.1, y: yPos, w: 5.1, h: 0.68,
          fill: { color: '0F172A' },
          line: { color: '334155', width: 1 },
          rectRadius: 0.06
        });
        s12.addText(`${item.pro}  →  ${item.adj}`, {
          x: 1.25, y: yPos + 0.12, w: 2.2, h: 0.4,
          fontSize: 12, fontFace: 'Arial', bold: true, color: 'FBBF24'
        });
        s12.addText(`${item.eg} (${item.vn})`, {
          x: 3.3, y: yPos + 0.14, w: 2.8, h: 0.4,
          fontSize: 11, fontFace: 'Calibri', color: 'E2E8F0'
        });
      });

      // Right Box: Plural (2 rows) + Core Law
      s12.addShape(pptx.ShapeType.roundRect, {
        x: 6.83, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.12
      });
      s12.addText("PLURAL FORMS & CORE GRAMMAR LAW", {
        x: 7.1, y: 2.0, w: 5.1, h: 0.3,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '34D399'
      });

      const pluralPairs = [
        { pro: "WE", adj: "OUR", eg: "our team", vn: "đội của chúng tôi" },
        { pro: "THEY", adj: "THEIR", eg: "their house", vn: "nhà của họ" }
      ];

      pluralPairs.forEach((item, idx) => {
        const yPos = 2.45 + (idx * 0.82);
        s12.addShape(pptx.ShapeType.roundRect, {
          x: 7.1, y: yPos, w: 5.1, h: 0.68,
          fill: { color: '0F172A' },
          line: { color: '334155', width: 1 },
          rectRadius: 0.06
        });
        s12.addText(`${item.pro}  →  ${item.adj}`, {
          x: 7.25, y: yPos + 0.12, w: 2.2, h: 0.4,
          fontSize: 12, fontFace: 'Arial', bold: true, color: '34D399'
        });
        s12.addText(`${item.eg} (${item.vn})`, {
          x: 9.3, y: yPos + 0.14, w: 2.8, h: 0.4,
          fontSize: 11, fontFace: 'Calibri', color: 'E2E8F0'
        });
      });

      // Core Grammar Law Card
      s12.addShape(pptx.ShapeType.roundRect, {
        x: 7.1, y: 4.3, w: 5.1, h: 2.1,
        fill: { color: '0F172A' },
        line: { color: 'F59E0B', width: 1.5 },
        rectRadius: 0.08
      });
      s12.addText("📌 QUY TẮC NGỮ PHÁP BẤT BIẾN:", {
        x: 7.3, y: 4.45, w: 4.7, h: 0.3,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s12.addText("[ Tính từ sở hữu ] + [ NOUN / Cụm Danh Từ ]", {
        x: 7.3, y: 4.85, w: 4.7, h: 0.35,
        fontSize: 12, fontFace: 'Courier New', bold: true, color: '6EE7B7'
      });
      s12.addText("Cả 7 tính từ sở hữu bắt buộc phải có danh từ đi liền phía sau để xác định quyền sở hữu. Chúng TUYỆT ĐỐI không bao giờ đứng độc lập một mình!", {
        x: 7.3, y: 5.25, w: 4.7, h: 0.95,
        fontSize: 11, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 18
      });

      attachNotes(s12, slidesData[11].speakerNotes);
    }

    // --- SLIDE 13: PART 03 PRACTICE - QUESTION 05 (CHECK 1) ---
    buildQuizSlide(slidesData[12], "CHECK 05",
      "\"John is looking for _______ keys. He cannot find them anywhere.\"",
      ["A. he", "B. him", "C. his", "D. himself"], "C",
      "Ngay sau khoảng trống là danh từ số nhiều keys (chìa khóa). Áp dụng quy tắc: [Tính từ sở hữu + Noun], vị trí này cần một tính từ sở hữu để bổ nghĩa cho keys. Chủ thể là danh từ chỉ người nam số ít (John), do đó ta chọn his."
    );

    // --- SLIDE 14: PART 03 PRACTICE - QUESTION 06 (CHECK 2) ---
    buildQuizSlide(slidesData[13], "CHECK 06",
      "\"We invited all of _______ friends to the end-of-year party.\"",
      ["A. our", "B. us", "C. ours", "D. we"], "A",
      "Phía sau khoảng trống có danh từ friends (những người bạn). Theo công thức [Tính từ sở hữu + Noun], ta cần một tính từ sở hữu. Với đại từ chủ ngữ We (chúng tôi), tính từ sở hữu tương ứng là our (our friends = những người bạn của chúng tôi). Lưu ý: ours là đại từ sở hữu, đứng một mình và không có danh từ friends theo sau."
    );


    // --- SLIDE 15: PART 04 - POSSESSIVE PRONOUNS (NO BROKEN IMAGES - RICH TABLE & PILLS) ---
    {
      const s15 = createSlide();
      addHeader(s15, slidesData[14].category, slidesData[14].title, slidesData[14].subtitle);

      // Left Box: Owner Table
      s15.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '8B5CF6', width: 1.5 },
        rectRadius: 0.12
      });
      s15.addText("OWNER", { x: 1.2, y: 2.1, w: 2.2, h: 0.3, fontSize: 11, fontFace: 'Arial', bold: true, color: '94A3B8' });
      s15.addText("POSSESSIVE PRONOUN", { x: 3.6, y: 2.1, w: 2.6, h: 0.3, fontSize: 11, fontFace: 'Arial', bold: true, color: 'C084FC' });

      const pronRows = [
        { owner: "I", pron: "mine" }, { owner: "You", pron: "yours" },
        { owner: "He", pron: "his" }, { owner: "She", pron: "hers" },
        { owner: "It", pron: "its" }, { owner: "We", pron: "ours" },
        { owner: "They", pron: "theirs" }
      ];

      pronRows.forEach((r, idx) => {
        const yPos = 2.5 + (idx * 0.55);
        s15.addText(r.owner, { x: 1.2, y: yPos, w: 2.2, h: 0.35, fontSize: 12, fontFace: 'Calibri', color: 'FFFFFF' });
        s15.addText(r.pron, { x: 3.6, y: yPos, w: 2.6, h: 0.35, fontSize: 12, fontFace: 'Courier New', bold: true, color: 'C084FC' });
      });

      // Right Box: Strategic Demonstration
      s15.addShape(pptx.ShapeType.roundRect, {
        x: 6.83, y: 1.8, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.12
      });
      s15.addText("STRATEGIC ROLE: STANDALONE OWNERSHIP", {
        x: 7.1, y: 2.0, w: 5.1, h: 0.3,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '34D399'
      });
      s15.addText("Used to declare ownership definitively without repeating the noun:\nâ€œWhose is this?â€", {
        x: 7.1, y: 2.4, w: 5.1, h: 0.8,
        fontSize: 13, fontFace: 'Calibri', color: 'E2E8F0', lineSpacing: 20
      });

      s15.addShape(pptx.ShapeType.roundRect, {
        x: 7.1, y: 3.4, w: 5.1, h: 1.6,
        fill: { color: '064E3B' },
        line: { color: '10B981', width: 1 },
        rectRadius: 0.08
      });
      s15.addText("EXECUTIVE DEMONSTRATION:\nâ€œThis book is MINE.â€\n= This book belongs to me.\n(Stands completely alone without any accompanying noun)", {
        x: 7.3, y: 3.5, w: 4.7, h: 1.4,
        fontSize: 12, fontFace: 'Calibri', bold: true, color: 'A7F3D0', lineSpacing: 18
      });

      attachNotes(s15, slidesData[14].speakerNotes);
    }

    // --- SLIDE 16: PART 03 & 04 SYNTHESIS ---
    {
      const s16 = createSlide();
      addHeader(s16, slidesData[15].category, slidesData[15].title, slidesData[15].subtitle);

      s16.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 11.73, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '6366F1', width: 1 },
        rectRadius: 0.12
      });

      const synthRows = [
        { person: "1st Person Singular", adj: "my laptop", pron: "mine" },
        { person: "2nd Person (Sing./Plur.)", adj: "your report", pron: "yours" },
        { person: "3rd Person Male", adj: "his office", pron: "his" },
        { person: "3rd Person Female", adj: "her proposal", pron: "hers" },
        { person: "1st Person Plural", adj: "our project", pron: "ours" },
        { person: "3rd Person Plural", adj: "their strategy", pron: "theirs" }
      ];

      // Table Header
      s16.addText("PERSON / ENTITY", { x: 1.2, y: 2.1, w: 3.5, h: 0.35, fontSize: 11, fontFace: 'Arial', bold: true, color: '94A3B8' });
      s16.addText("POSSESSIVE ADJ (+ NOUN)", { x: 4.8, y: 2.1, w: 3.8, h: 0.35, fontSize: 11, fontFace: 'Arial', bold: true, color: 'FBBF24' });
      s16.addText("POSSESSIVE PRONOUN (STANDALONE)", { x: 8.7, y: 2.1, w: 3.5, h: 0.35, fontSize: 11, fontFace: 'Arial', bold: true, color: '34D399' });

      synthRows.forEach((r, idx) => {
        const yPos = 2.6 + (idx * 0.65);
        s16.addText(r.person, { x: 1.2, y: yPos, w: 3.5, h: 0.4, fontSize: 12, fontFace: 'Calibri', color: 'FFFFFF' });
        s16.addText(r.adj, { x: 4.8, y: yPos, w: 3.8, h: 0.4, fontSize: 12, fontFace: 'Calibri', bold: true, color: 'FDE68A' });
        s16.addText(r.pron, { x: 8.7, y: yPos, w: 3.5, h: 0.4, fontSize: 12, fontFace: 'Calibri', bold: true, color: '6EE7B7' });
      });

      attachNotes(s16, slidesData[15].speakerNotes);
    }

    // --- SLIDE 17: PART 04 PRACTICE - QUESTION 07 ---
    buildQuizSlide(slidesData[16], "CHECK 07",
      "â€œMy car is in the repair shop, but ___ is parked outside.â€",
      ["A. your", "B. yours", "C. you", "D. yourself"], "B",
      "Chá»— trá»‘ng lÃ m chá»§ ngá»¯ cá»§a váº¿ sau vÃ  khÃ´ng cÃ³ danh tá»« Ä‘i kÃ¨m âž” báº¯t buá»™c dÃ¹ng Possessive Pronoun 'yours' (= your car). PhÆ°Æ¡ng Ã¡n 'your' báº¯t buá»™c pháº£i cÃ³ danh tá»« phÃ­a sau."
    );

    // --- SLIDE 18: PART 04 PRACTICE - QUESTION 08 ---
    buildQuizSlide(slidesData[17], "CHECK 08",
      "â€œThese project documents belong to the marketing team; in fact, they are ___.â€",
      ["A. their", "B. them", "C. theirs", "D. themselves"], "C",
      "Äá»©ng sau Ä‘á»™ng tá»« 'are' á»Ÿ cuá»‘i cÃ¢u Ä‘á»ƒ chá»‰ sá»± sá»Ÿ há»¯u Ä‘á»™c láº­p (they are theirs = they are their documents). Báº¯t buá»™c dÃ¹ng Possessive Pronoun 'theirs'."
    );

    // --- SLIDE 19: PART 05 - REFLEXIVE PRONOUNS (CORE FOCUS) ---
    {
      const s19 = createSlide();
      addHeader(s19, slidesData[18].category, slidesData[18].title, slidesData[18].subtitle);

      s19.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 6.2, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: 'EC4899', width: 1.5 },
        rectRadius: 0.12
      });
      s19.addText("3 STRATEGIC WORKPLACE FUNCTIONS", {
        x: 1.1, y: 2.0, w: 5.6, h: 0.3,
        fontSize: 10, fontFace: 'Arial', bold: true, color: 'F472B6'
      });
      s19.addText("1. Reflection (Subject = Object):\nAction reflects back onto the actor: â€œI love myself.â€\n\n2. Emphatic Intensifier:\nEmphasizes personal execution: â€œShe carries these books herself.â€\n\n3. Solo Execution (By + Oneself):\nCompleting a task alone: â€œI did it by myself.â€ (= alone)", {
        x: 1.1, y: 2.4, w: 5.6, h: 2.6,
        fontSize: 11.5, fontFace: 'Calibri', color: 'E2E8F0', lineSpacing: 18
      });
      s19.addText("Forms: myself, yourself, himself, herself, itself, ourselves, yourselves, themselves", {
        x: 1.1, y: 5.3, w: 5.6, h: 0.8,
        fontSize: 10.5, fontFace: 'Courier New', color: 'FBCFE8'
      });

      if (assets.reflexive_mirror) {
        s19.addImage({ data: assets.reflexive_mirror, x: 7.4, y: 1.8, w: 5.13, h: 5.0, round: true });
      }
      attachNotes(s19, slidesData[18].speakerNotes);
    }

    // --- SLIDE 20: PART 05 - OVERVIEW OF OTHER PRONOUNS ---
    {
      const s20 = createSlide();
      addHeader(s20, slidesData[19].category, slidesData[19].title, slidesData[19].subtitle);

      // Left: Spatial pointer image
      if (assets.spatial_pointers) {
        s20.addImage({ data: assets.spatial_pointers, x: 0.8, y: 1.8, w: 4.8, h: 4.9, round: true });
      }

      // Right: 4 Overview Cards
      const extCards = [
        {
          title: "1. DEMONSTRATIVE (Chá»‰ Ä‘á»‹nh)",
          words: "This, That, These, Those",
          rules: "Gáº§n: This / These â€¢ Xa: That / Those. CÃ³ thá»ƒ Ä‘i kÃ¨m danh tá»« khÃ´ng Ä‘áº¿m Ä‘Æ°á»£c (This advice).",
          color: '38BDF8'
        },
        {
          title: "2. INDEFINITE (Báº¥t Ä‘á»‹nh)",
          words: "everyone, someone, anything...",
          rules: "âš ï¸ QUY Táº®C VÃ€NG: LuÃ´n chia Ä‘á»™ng tá»« sá»‘ Ã­t (Everyone is ready).",
          color: '34D399'
        },
        {
          title: "3. INTERROGATIVE (Nghi váº¥n)",
          words: "who, whom, whose, what, which",
          rules: "DÃ¹ng Ä‘á»ƒ Ä‘áº·t cÃ¢u há»i trá»±c tiáº¿p hoáº·c giÃ¡n tiáº¿p: â€œWhose jacket is this?â€",
          color: '60A5FA'
        },
        {
          title: "4. RELATIVE (Quan há»‡)",
          words: "who, which, that, whose, whoever",
          rules: "Ná»‘i má»‡nh Ä‘á» phá»¥: â€œThe candidate who won the election...â€",
          color: 'FBBF24'
        }
      ];

      extCards.forEach((c, idx) => {
        const yPos = 1.8 + (idx * 1.25);
        s20.addShape(pptx.ShapeType.roundRect, {
          x: 5.8, y: yPos, w: 6.73, h: 1.15,
          fill: { color: '1E293B' },
          line: { color: c.color, width: 1.2 },
          rectRadius: 0.08
        });
        s20.addText(c.title + "  |  " + c.words, {
          x: 6.0, y: yPos + 0.1, w: 6.3, h: 0.3,
          fontSize: 10.5, fontFace: 'Arial', bold: true, color: c.color
        });
        s20.addText(c.rules, {
          x: 6.0, y: yPos + 0.45, w: 6.3, h: 0.6,
          fontSize: 11, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 15
        });
      });

      attachNotes(s20, slidesData[19].speakerNotes);
    }

    // --- SLIDE 21: PART 05 PRACTICE - QUESTION 09 ---
    buildQuizSlide(slidesData[20], "CHECK 09",
      "\"Due to the unexpected absence of her assistant, Ms. Gable had to organize the entire quarterly conference on _______.\"",
      ["A. her own", "B. her", "C. hers", "D. herself"], "A",
      "ThÃ nh ngá»¯ cá»‘ Ä‘á»‹nh 'on one's own' = tá»± mÃ¬nh lÃ m, khÃ´ng cÃ³ sá»± trá»£ giÃºp (independently / without help). Giá»›i tá»« 'on' báº¯t buá»™c Ä‘i vá»›i 'her own'. Náº¿u dÃ¹ng Ä‘áº¡i tá»« pháº£n thÃ¢n thÃ¬ pháº£i lÃ  'by herself' chá»© khÃ´ng dÃ¹ng 'on herself'."
    );

    // --- SLIDE 22: PART 05 PRACTICE - QUESTION 10 ---
    buildQuizSlide(slidesData[21], "CHECK 10",
      "\"Before submitting the financial auditing report to the board, Ms. Patel checked the spreadsheet _______ to ensure zero calculation errors.\"",
      ["A. hers", "B. herself", "C. she", "D. her"], "B",
      "Äáº¡i tá»« pháº£n thÃ¢n 'herself' Ä‘á»©ng cuá»‘i má»‡nh Ä‘á» Ä‘Ã³ng vai trÃ² Ä‘áº¡i tá»« nháº¥n máº¡nh (Intensive Pronoun) Ä‘á»ƒ nháº¥n máº¡nh Ä‘Ã­ch thÃ¢n cÃ´ Patel Ä‘Ã£ tá»± tay kiá»ƒm tra báº£ng tÃ­nh. Vá»›i chá»§ ngá»¯ ná»¯ 'Ms. Patel' âž” chá»n 'herself'."
    );

    // --- SLIDE 23: EXECUTIVE DECISION FRAMEWORK ---
    {
      const s23 = createSlide();
      addHeader(s23, slidesData[22].category, slidesData[22].title, slidesData[22].subtitle);

      const questions = [
        {
          num: "QUESTION 1",
          q: "Who does the action?",
          sub: "â€¢ Initiator before Verb âž” Subject Pronoun\n(He plays football)",
          color: '6366F1'
        },
        {
          num: "QUESTION 2",
          q: "Who/What receives?",
          sub: "â€¢ Target after Verb/Preposition âž” Object Pronoun\n(I like him / between you and me)",
          color: '38BDF8'
        },
        {
          num: "QUESTION 3",
          q: "Bounces back or alone?",
          sub: "â€¢ Action reflects on doer / Solo âž” Reflexive\n(He hurt himself / by himself)",
          color: 'EC4899'
        }
      ];

      questions.forEach((item, idx) => {
        const xPos = 0.8 + (idx * 4.0);
        s23.addShape(pptx.ShapeType.roundRect, {
          x: xPos, y: 1.8, w: 3.73, h: 4.8,
          fill: { color: '1E293B' },
          line: { color: item.color, width: 1.5 },
          rectRadius: 0.12
        });
        s23.addText(item.num, {
          x: xPos + 0.2, y: 2.1, w: 3.33, h: 0.35,
          fontSize: 11, fontFace: 'Arial', bold: true, color: item.color, align: 'center'
        });
        s23.addText(item.q, {
          x: xPos + 0.2, y: 2.6, w: 3.33, h: 0.8,
          fontSize: 13, fontFace: 'Arial', bold: true, color: 'FFFFFF', align: 'center'
        });
        s23.addText(item.sub, {
          x: xPos + 0.2, y: 3.6, w: 3.33, h: 2.6,
          fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 22
        });
      });

      attachNotes(s23, slidesData[22].speakerNotes);
    }

    // --- SLIDE 24: MASTER TAXONOMY MATRIX ---
    {
      const s24 = createSlide();
      addHeader(s24, slidesData[23].category, slidesData[23].title, slidesData[23].subtitle);

      s24.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.8, w: 11.73, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '38BDF8', width: 1 },
        rectRadius: 0.12
      });

      const matrixRows = [
        { entity: "I (TÃ´i)", s: "I", o: "me", a: "my + N", p: "mine", r: "myself" },
        { entity: "You (Báº¡n)", s: "you", o: "you", a: "your + N", p: "yours", r: "yourself / -selves" },
        { entity: "He (Anh áº¥y)", s: "he", o: "him", a: "his + N", p: "his", r: "himself" },
        { entity: "She (CÃ´ áº¥y)", s: "she", o: "her", a: "her + N", p: "hers", r: "herself" },
        { entity: "It (NÃ³)", s: "it", o: "it", a: "its + N", p: "its", r: "itself" },
        { entity: "We (ChÃºng tÃ´i)", s: "we", o: "us", a: "our + N", p: "ours", r: "ourselves" },
        { entity: "They (Há»)", s: "they", o: "them", a: "their + N", p: "theirs", r: "themselves" }
      ];

      s24.addText("ENTITY", { x: 1.0, y: 2.05, w: 2.4, h: 0.3, fontSize: 10, fontFace: 'Arial', bold: true, color: '94A3B8' });
      s24.addText("SUBJECT", { x: 3.5, y: 2.05, w: 1.6, h: 0.3, fontSize: 10, fontFace: 'Arial', bold: true, color: '818CF8' });
      s24.addText("OBJECT", { x: 5.2, y: 2.05, w: 1.6, h: 0.3, fontSize: 10, fontFace: 'Arial', bold: true, color: '22D3EE' });
      s24.addText("POSS. ADJ", { x: 6.9, y: 2.05, w: 1.8, h: 0.3, fontSize: 10, fontFace: 'Arial', bold: true, color: 'FBBF24' });
      s24.addText("POSS. PRON", { x: 8.8, y: 2.05, w: 1.8, h: 0.3, fontSize: 10, fontFace: 'Arial', bold: true, color: '34D399' });
      s24.addText("REFLEXIVE", { x: 10.7, y: 2.05, w: 1.6, h: 0.3, fontSize: 10, fontFace: 'Arial', bold: true, color: 'F472B6' });

      matrixRows.forEach((r, idx) => {
        const yPos = 2.45 + (idx * 0.58);
        s24.addText(r.entity, { x: 1.0, y: yPos, w: 2.4, h: 0.35, fontSize: 11, fontFace: 'Calibri', color: 'FFFFFF' });
        s24.addText(r.s, { x: 3.5, y: yPos, w: 1.6, h: 0.35, fontSize: 11, fontFace: 'Calibri', bold: true, color: 'A5B4FC' });
        s24.addText(r.o, { x: 5.2, y: yPos, w: 1.6, h: 0.35, fontSize: 11, fontFace: 'Calibri', bold: true, color: '67E8F9' });
        s24.addText(r.a, { x: 6.9, y: yPos, w: 1.8, h: 0.35, fontSize: 11, fontFace: 'Calibri', bold: true, color: 'FDE68A' });
        s24.addText(r.p, { x: 8.8, y: yPos, w: 1.8, h: 0.35, fontSize: 11, fontFace: 'Calibri', bold: true, color: '6EE7B7' });
        s24.addText(r.r, { x: 10.7, y: yPos, w: 1.6, h: 0.35, fontSize: 11, fontFace: 'Calibri', bold: true, color: 'FBCFE8' });
      });

      attachNotes(s24, slidesData[23].speakerNotes);
    }

    // --- SLIDE 25: GRAND CHALLENGE ---
    {
      const s25 = createSlide();
      addHeader(s25, slidesData[24].category, slidesData[24].title, slidesData[24].subtitle);

      // Left: Sarah Laptop Image
      if (assets.sarah_laptop) {
        s25.addImage({ data: assets.sarah_laptop, x: 0.8, y: 1.8, w: 5.2, h: 4.9, round: true });
      }

      // Right: Challenge Question Box
      s25.addShape(pptx.ShapeType.roundRect, {
        x: 6.3, y: 1.8, w: 6.23, h: 4.9,
        fill: { color: '1E293B' },
        line: { color: 'EF4444', width: 1.5 },
        rectRadius: 0.12
      });
      s25.addShape(pptx.ShapeType.roundRect, {
        x: 6.6, y: 2.05, w: 2.0, h: 0.35,
        fill: { color: 'EF4444' },
        line: { color: 'EF4444' },
        rectRadius: 0.08
      });
      s25.addText("GRAND CHALLENGE", {
        x: 6.6, y: 2.05, w: 2.0, h: 0.35,
        fontSize: 10, fontFace: 'Arial', bold: true, color: 'FFFFFF', align: 'center', valign: 'middle'
      });
      s25.addText("â€œSarah has a new laptop. ___ laptop is very expensive, but the laptop is not ___.â€", {
        x: 6.6, y: 2.55, w: 5.6, h: 1.1,
        fontSize: 13.5, fontFace: 'Arial', bold: true, color: 'FFFFFF', lineSpacing: 20
      });

      const grandOpts = [
        "A. Hers / her",
        "B. Her / hers",
        "C. She / her",
        "D. Her / she"
      ];
      grandOpts.forEach((opt, idx) => {
        const isCorrect = idx === 1;
        const yPos = 3.8 + (idx * 0.48);
        s25.addShape(pptx.ShapeType.roundRect, {
          x: 6.6, y: yPos, w: 5.6, h: 0.4,
          fill: { color: isCorrect ? '064E3B' : '0F172A' },
          line: { color: isCorrect ? '10B981' : '334155', width: 1 },
          rectRadius: 0.06
        });
        s25.addText(opt + (isCorrect ? "  âœ“ [CORRECT ANSWER]" : ""), {
          x: 6.8, y: yPos, w: 5.2, h: 0.4,
          fontSize: 11, fontFace: 'Calibri', bold: isCorrect,
          color: isCorrect ? '34D399' : 'E2E8F0', valign: 'middle'
        });
      });

      s25.addText("Breakdown: 1. 'Her laptop' (precedes noun) | 2. 'not hers' (stands alone at end)", {
        x: 6.6, y: 5.85, w: 5.6, h: 0.6,
        fontSize: 10.5, fontFace: 'Calibri', italic: true, color: '94A3B8'
      });

      attachNotes(s25, slidesData[24].speakerNotes);
    }

    // --- SLIDE 26: CONCLUSION & DISCUSSION ---
    {
      const s26 = createSlide();
      addHeader(s26, slidesData[25].category, slidesData[25].title, slidesData[25].subtitle);

      s26.addShape(pptx.ShapeType.roundRect, {
        x: 1.5, y: 2.0, w: 10.33, h: 4.5,
        fill: { color: '1E293B' },
        line: { color: '6366F1', width: 1.5 },
        rectRadius: 0.15
      });
      s26.addText("â€œPronouns may be small words, but they are very important in English.â€", {
        x: 2.0, y: 2.5, w: 9.33, h: 1.2,
        fontSize: 22, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        align: 'center', lineSpacing: 30
      });
      s26.addText("Pronouns appear in almost every sentence. By mastering who/what they replace, their grammatical position, and their standalone vs. modifier roles, natural English becomes effortless.", {
        x: 2.0, y: 3.8, w: 9.33, h: 1.2,
        fontSize: 14, fontFace: 'Calibri', color: '94A3B8',
        align: 'center', lineSpacing: 22
      });
      s26.addShape(pptx.ShapeType.roundRect, {
        x: 4.8, y: 5.2, w: 3.73, h: 0.6,
        fill: { color: '6366F1' },
        line: { color: '6366F1' },
        rectRadius: 0.1
      });
      s26.addText("Thank You! Opening Floor for Q&A", {
        x: 4.8, y: 5.2, w: 3.73, h: 0.6,
        fontSize: 12, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        align: 'center', valign: 'middle'
      });

      attachNotes(s26, slidesData[25].speakerNotes);
    }

    // Write file to client browser
    pptx.writeFile({ fileName: "Pronouns_Keynote_Masterclass_26_Slides.pptx" })
      .then(() => {
        if (btn) {
          btn.innerHTML = originalText;
          btn.disabled = false;
        }
      })
      .catch((err) => {
        console.error("PPTX Generation Error:", err);
        if (btn) {
          btn.innerHTML = `<span>Lá»—i xuáº¥t PPTX</span>`;
          setTimeout(() => {
            btn.innerHTML = originalText;
            btn.disabled = false;
          }, 3000);
        }
      });

  } catch (err) {
    console.error("PPTX Initialization Error:", err);
    if (btn) {
      btn.innerHTML = `<span>Lá»—i khá»Ÿi táº¡o</span>`;
      setTimeout(() => {
        btn.innerHTML = originalText;
        btn.disabled = false;
      }, 3000);
    }
  }
};