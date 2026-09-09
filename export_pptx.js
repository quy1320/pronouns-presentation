// Dedicated Executive-Grade PowerPoint (.pptx) Exporter
// Produces 16:9 widescreen slides with modern card layouts, embedded 3D illustrations, styled quiz cards, and verbatim Vietnamese speaker notes.

window.exportPPTX = function() {
  const btn = document.getElementById('exportPptxBtn');
  const originalText = btn.innerHTML;
  btn.innerHTML = `<span>Đang tạo PowerPoint cao cấp...</span>`;
  btn.disabled = true;

  try {
    const pptx = new PptxGenJS();
    // Executive 16:9 widescreen layout (13.333 in x 7.5 in - standard PowerPoint 365)
    pptx.defineLayout({ name: 'WIDE169', width: 13.333, height: 7.5 });
    pptx.layout = 'WIDE169';
    pptx.author = 'English Masterclass Keynote';
    pptx.company = 'English Mastery';
    pptx.title = 'English Pronouns Keynote Presentation';

    const assets = window.SLIDE_ASSETS || {};

    // Helper: Create slide with 100% full-bleed dark navy background (guarantees zero white borders)
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
      // Category pill
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 0.35, w: 2.8, h: 0.3,
        fill: { color: '1E293B' },
        line: { color: '38BDF8', width: 1 },
        rectRadius: 0.1
      });
      slide.addText(category.toUpperCase(), {
        x: 0.8, y: 0.35, w: 2.8, h: 0.3,
        fontSize: 9, fontFace: 'Calibri', bold: true, color: '38BDF8',
        align: 'center', valign: 'middle'
      });

      // Title
      slide.addText(title, {
        x: 0.8, y: 0.75, w: 11.7, h: 0.55,
        fontSize: 22, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        valign: 'middle'
      });

      // Subtitle
      slide.addText(subtitle, {
        x: 0.8, y: 1.3, w: 11.7, h: 0.3,
        fontSize: 11, fontFace: 'Calibri', color: '94A3B8',
        valign: 'top'
      });
    }

    // Helper: Add clean Vietnamese speaker notes
    function attachNotes(slide, rawNotes) {
      const cleanNotes = rawNotes
        .replace(/<[^>]+>/g, ' ')
        .replace(/&quot;/g, '"')
        .replace(/&amp;/g, '&')
        .replace(/\s+/g, ' ')
        .trim();
      slide.addNotes(cleanNotes);
    }

    // Helper: Generic Quiz Slide Builder
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
        fontSize: 15, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        valign: 'middle'
      });

      // 4 Option Cards in a 2x2 Grid
      const positions = [
        { x: 0.8, y: 3.0 },
        { x: 6.83, y: 3.0 },
        { x: 0.8, y: 4.0 },
        { x: 6.83, y: 4.0 }
      ];

      optionsArr.forEach((opt, i) => {
        const isCorrect = (opt.charAt(0) === correctOpt);
        const pos = positions[i];
        
        slide.addShape(pptx.ShapeType.roundRect, {
          x: pos.x, y: pos.y, w: 5.7, h: 0.85,
          fill: { color: isCorrect ? '064E3B' : '1E293B' },
          line: { color: isCorrect ? '10B981' : '334155', width: isCorrect ? 2 : 1 },
          rectRadius: 0.1
        });

        // Letter badge
        slide.addShape(pptx.ShapeType.roundRect, {
          x: pos.x + 0.25, y: pos.y + 0.18, w: 0.48, h: 0.48,
          fill: { color: isCorrect ? '10B981' : '334155' },
          line: { color: isCorrect ? '10B981' : '475569' },
          rectRadius: 0.24
        });
        slide.addText(opt.charAt(0), {
          x: pos.x + 0.25, y: pos.y + 0.18, w: 0.48, h: 0.48,
          fontSize: 11, fontFace: 'Arial', bold: true, color: 'FFFFFF',
          align: 'center', valign: 'middle'
        });

        // Option text
        slide.addText(opt.substring(3) + (isCorrect ? '  ✓ (CORRECT)' : ''), {
          x: pos.x + 0.9, y: pos.y, w: 4.6, h: 0.85,
          fontSize: 13, fontFace: 'Arial', bold: isCorrect,
          color: isCorrect ? '34D399' : 'E2E8F0',
          valign: 'middle'
        });
      });

      // Explanation Box
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 5.05, w: 11.73, h: 1.7,
        fill: { color: '0F2922' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.12
      });
      slide.addText(`CORRECT ANSWER: ${correctOpt}  •  LINGUISTIC RATIONALE`, {
        x: 1.1, y: 5.2, w: 11.1, h: 0.35,
        fontSize: 12, fontFace: 'Arial', bold: true, color: '34D399'
      });
      slide.addText(explanation, {
        x: 1.1, y: 5.6, w: 11.1, h: 1.0,
        fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1',
        lineSpacing: 18
      });

      attachNotes(slide, slideItem.speakerNotes);
    }

    // ==========================================
    // BUILD ALL 25 SLIDES
    // ==========================================

    // --- SLIDE 1: COVER ---
    {
      const s1 = createSlide();
      
      // Left Hero Card
      s1.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 0.9, w: 6.2, h: 5.7,
        fill: { color: '1E293B' },
        line: { color: '334155', width: 1.5 },
        rectRadius: 0.15
      });
      // Pill
      s1.addShape(pptx.ShapeType.roundRect, {
        x: 1.2, y: 1.3, w: 2.8, h: 0.35,
        fill: { color: '312E81' },
        line: { color: '6366F1', width: 1 },
        rectRadius: 0.08
      });
      s1.addText("KEYNOTE PRESENTATION", {
        x: 1.2, y: 1.3, w: 2.8, h: 0.35,
        fontSize: 9, fontFace: 'Arial', bold: true, color: 'A5B4FC',
        align: 'center', valign: 'middle'
      });
      // Title
      s1.addText("PRONOUNS IN ENGLISH", {
        x: 1.2, y: 1.85, w: 5.4, h: 1.1,
        fontSize: 30, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        lineSpacing: 34
      });
      s1.addText("How Small Words Drive Natural, Fluent Communication", {
        x: 1.2, y: 3.05, w: 5.4, h: 0.6,
        fontSize: 13, fontFace: 'Calibri', color: '94A3B8'
      });
      // Bullets
      s1.addText("• 5 Essential Pronoun Families (Subject, Object, Possessive, Reflexive, Demonstrative)\n• Eliminating awkward repetition & sounding like an executive\n• 10 Live interactive audience polls + 1 Rapid-Fire Challenge", {
        x: 1.2, y: 3.75, w: 5.4, h: 1.5,
        fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1',
        lineSpacing: 22
      });
      // Badge
      s1.addShape(pptx.ShapeType.roundRect, {
        x: 1.2, y: 5.45, w: 2.9, h: 0.4,
        fill: { color: '064E3B' },
        line: { color: '10B981', width: 1 },
        rectRadius: 0.08
      });
      s1.addText("30-Minute Executive Session", {
        x: 1.2, y: 5.45, w: 2.9, h: 0.4,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '6EE7B7',
        align: 'center', valign: 'middle'
      });

      // Right Image Card
      if (assets.hero_cover) {
        s1.addImage({
          data: assets.hero_cover,
          x: 7.3, y: 0.9, w: 5.23, h: 5.7,
          round: true
        });
      }
      attachNotes(s1, slidesData[0].speakerNotes);
    }

    // --- SLIDE 2: THE REPETITION STORY ---
    {
      const s2 = createSlide();
      addHeader(s2, slidesData[1].category, slidesData[1].title, slidesData[1].subtitle);

      // Left Column: Story & Solution
      // Story Card
      s2.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.7, w: 6.4, h: 2.0,
        fill: { color: '1E293B' },
        line: { color: 'F87171', width: 1.5 },
        rectRadius: 0.12
      });
      s2.addText("WITHOUT PRONOUNS (Exhausted & Robotic):", {
        x: 1.1, y: 1.85, w: 5.8, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: 'F87171'
      });
      s2.addText("“If I say: Tom is a student. Tom likes football. Tom plays football every day. Tom lives near my house...”", {
        x: 1.1, y: 2.2, w: 5.8, h: 0.9,
        fontSize: 13, fontFace: 'Calibri', italic: true, color: 'E2E8F0',
        lineSpacing: 20
      });
      s2.addText("❓ Observation: Notice how repeating 'Tom' makes speech stiff and repetitive.", {
        x: 1.1, y: 3.15, w: 5.8, h: 0.4,
        fontSize: 11, fontFace: 'Calibri', bold: true, color: 'FBBF24'
      });

      // Fluent Delivery Card
      s2.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 3.9, w: 6.4, h: 2.8,
        fill: { color: '064E3B' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.12
      });
      s2.addText("THE FLUENT NATIVE DELIVERY (With Pronouns):", {
        x: 1.1, y: 4.1, w: 5.8, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: '34D399'
      });
      s2.addText("“Tom is a student. He likes football. He plays football every day.”", {
        x: 1.1, y: 4.5, w: 5.8, h: 0.8,
        fontSize: 16, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        lineSpacing: 24
      });
      s2.addText("👉 The pronoun 'He' replaces the noun 'Tom' to keep speech smooth, concise, and natural.", {
        x: 1.1, y: 5.4, w: 5.8, h: 1.0,
        fontSize: 12, fontFace: 'Calibri', color: 'A7F3D0',
        lineSpacing: 18
      });

      // Right Image: Tom
      if (assets.tom_football) {
        s2.addImage({
          data: assets.tom_football,
          x: 7.5, y: 1.7, w: 5.03, h: 5.0,
          round: true
        });
      }
      attachNotes(s2, slidesData[1].speakerNotes);
    }

    // --- SLIDE 3: DEFINITION & STRATEGIC VALUE ---
    {
      const s3 = createSlide();
      addHeader(s3, slidesData[2].category, slidesData[2].title, slidesData[2].subtitle);

      // Left Card: Definition
      s3.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.7, w: 5.6, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '6366F1', width: 1.5 },
        rectRadius: 0.12
      });
      s3.addText("LINGUISTIC DEFINITION", {
        x: 1.1, y: 1.95, w: 5.0, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: '818CF8'
      });
      s3.addText("Pronoun = Replaces a Noun", {
        x: 1.1, y: 2.3, w: 5.0, h: 0.5,
        fontSize: 18, fontFace: 'Arial', bold: true, color: 'FFFFFF'
      });
      s3.addText("A pronoun is a word used in place of a noun or noun phrase to refer to people, objects, or concepts without naming them repeatedly.", {
        x: 1.1, y: 2.85, w: 5.0, h: 1.1,
        fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1',
        lineSpacing: 18
      });
      // Examples Box
      s3.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: 4.15, w: 5.0, h: 2.2,
        fill: { color: '0F172A' },
        line: { color: '334155', width: 1 },
        rectRadius: 0.08
      });
      s3.addText("• Tom  ➔  he\n• Anna  ➔  she\n• The book  ➔  it\n• Tom and Anna  ➔  they", {
        x: 1.3, y: 4.3, w: 4.6, h: 1.9,
        fontSize: 14, fontFace: 'Arial', bold: true, color: '38BDF8',
        lineSpacing: 26
      });

      // Right Card: 4 Strategic Benefits
      s3.addShape(pptx.ShapeType.roundRect, {
        x: 6.8, y: 1.7, w: 5.73, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: 'F59E0B', width: 1.5 },
        rectRadius: 0.12
      });
      s3.addText("EXECUTIVE COMMUNICATION VALUE", {
        x: 7.1, y: 1.95, w: 5.1, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s3.addText("Why Do We Need Pronouns?", {
        x: 7.1, y: 2.3, w: 5.1, h: 0.5,
        fontSize: 18, fontFace: 'Arial', bold: true, color: 'FFFFFF'
      });

      const benefits = [
        { num: "01", title: "Eliminate Word Repetition", desc: "Keeps professional speech clean, sharp, and executive." },
        { num: "02", title: "Concise Delivery", desc: "Conveys high-impact ideas with fewer syllables and zero drag." },
        { num: "03", title: "Natural Speech Rhythm", desc: "Mimics native speaker cadence, tempo, and acoustic flow." },
        { num: "04", title: "Context & Referent Clarity", desc: "Clearly identifies who or what is being discussed with zero doubt." }
      ];

      let bY = 2.95;
      benefits.forEach(b => {
        s3.addText(b.num + ".  " + b.title, {
          x: 7.1, y: bY, w: 5.1, h: 0.35,
          fontSize: 13, fontFace: 'Arial', bold: true, color: 'FBBF24'
        });
        s3.addText(b.desc, {
          x: 7.6, y: bY + 0.3, w: 4.6, h: 0.5,
          fontSize: 11, fontFace: 'Calibri', color: 'CBD5E1'
        });
        bY += 0.85;
      });

      attachNotes(s3, slidesData[2].speakerNotes);
    }

    // --- SLIDE 4: ROADMAP - 8 ESSENTIAL FAMILIES ---
    {
      const s4 = createSlide();
      addHeader(s4, slidesData[3].category, slidesData[3].title, slidesData[3].subtitle);

      const cards = [
        { x: 0.8, y: 1.7, w: 2.7, h: 2.3, color: '818CF8', title: "1. Personal", sub: "Subject & Object roles", items: "I, he, she, him, them..." },
        { x: 3.8, y: 1.7, w: 2.7, h: 2.3, color: '38BDF8', title: "2. Possessive", sub: "Ownership & Equation", items: "mine = my books..." },
        { x: 6.8, y: 1.7, w: 2.7, h: 2.3, color: 'C084FC', title: "3. Reflexive", sub: "Self-action & by myself", items: "myself, herself..." },
        { x: 9.8, y: 1.7, w: 2.7, h: 2.3, color: 'FBBF24', title: "4. Demonstrative", sub: "Distance & Uncountables", items: "this, that, these, those" },
        { x: 0.8, y: 4.25, w: 2.7, h: 2.3, color: '34D399', title: "5. Indefinite", sub: "Singular Verb Law", items: "everyone, someone..." },
        { x: 3.8, y: 4.25, w: 2.7, h: 2.3, color: 'F472B6', title: "6. Interrogative", sub: "Precise Inquiries", items: "who, whom, whose..." },
        { x: 6.8, y: 4.25, w: 2.7, h: 2.3, color: '60A5FA', title: "7. Relative", sub: "TOEIC Clause Connectors", items: "who, which, that..." },
        { x: 9.8, y: 4.25, w: 2.7, h: 2.3, color: 'F59E0B', title: "8. Compound", sub: "Open Choice Mastery", items: "whoever, whichever..." }
      ];

      cards.forEach(c => {
        s4.addShape(pptx.ShapeType.roundRect, {
          x: c.x, y: c.y, w: c.w, h: c.h,
          fill: { color: '1E293B' },
          line: { color: c.color, width: 1.5 },
          rectRadius: 0.12
        });
        s4.addText(c.title, {
          x: c.x + 0.2, y: c.y + 0.2, w: c.w - 0.4, h: 0.35,
          fontSize: 14, fontFace: 'Arial', bold: true, color: c.color
        });
        s4.addText(c.sub, {
          x: c.x + 0.2, y: c.y + 0.6, w: c.w - 0.4, h: 0.35,
          fontSize: 10, fontFace: 'Calibri', color: '94A3B8'
        });
        s4.addText(c.items, {
          x: c.x + 0.2, y: c.y + 1.1, w: c.w - 0.4, h: 0.9,
          fontSize: 12, fontFace: 'Arial', bold: true, color: 'FFFFFF',
          lineSpacing: 16
        });
      });

      attachNotes(s4, slidesData[3].speakerNotes);
    }

    // --- SLIDE 5: SUBJECT PRONOUNS ---
    {
      const s5 = createSlide();
      addHeader(s5, slidesData[4].category, slidesData[4].title, slidesData[4].subtitle);

      // Left Side: Table
      const subTable = [
        [{ text: "Grammatical Role / Person", options: { bold: true, fill: "312E81", color: "FFFFFF" } }, { text: "Subject Pronoun", options: { bold: true, fill: "312E81", color: "38BDF8" } }],
        ["1st Person Singular", "I"],
        ["2nd Person (Singular / Plural)", "You"],
        ["3rd Person Masculine", "He"],
        ["3rd Person Feminine", "She"],
        ["3rd Person Inanimate / Animal", "It"],
        ["1st Person Plural", "We"],
        ["3rd Person Plural", "They"]
      ];
      s5.addTable(subTable, {
        x: 0.8, y: 1.7, w: 5.8,
        colW: [3.8, 2.0],
        rowH: 0.5,
        fontSize: 12, fontFace: 'Arial', color: 'E2E8F0',
        border: { pt: '1', color: '334155' },
        fill: '1E293B'
      });

      // Right Side: Rule & Examples
      s5.addShape(pptx.ShapeType.roundRect, {
        x: 6.9, y: 1.7, w: 5.63, h: 2.2,
        fill: { color: '1E293B' },
        line: { color: 'F59E0B', width: 1.5 },
        rectRadius: 0.12
      });
      s5.addText("⚡ THE STRUCTURAL SYNTAX RULE:", {
        x: 7.2, y: 1.9, w: 5.0, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s5.addText("Subject Pronoun + Verb", {
        x: 7.2, y: 2.25, w: 5.0, h: 0.6,
        fontSize: 22, fontFace: 'Arial', bold: true, color: 'FFFFFF'
      });
      s5.addText("The Subject Pronoun always initiates the action and is positioned BEFORE the main verb in statements.\n\nClassic Sentence Anchor: “I love you” ➔ ‘I’ is the Subject initiating the verb ‘love’.", {
        x: 7.2, y: 2.9, w: 5.0, h: 1.1,
        fontSize: 11, fontFace: 'Calibri', color: 'CBD5E1'
      });

      // Examples Card
      s5.addShape(pptx.ShapeType.roundRect, {
        x: 6.9, y: 4.15, w: 5.63, h: 2.55,
        fill: { color: '1E293B' },
        line: { color: '38BDF8', width: 1.5 },
        rectRadius: 0.12
      });
      s5.addText("NATURAL SENTENCE DEMONSTRATIONS:", {
        x: 7.2, y: 4.35, w: 5.0, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: '38BDF8'
      });
      s5.addText("• I am a student.\n• She is my colleague.\n• He plays football every afternoon.\n• They work at our global office.", {
        x: 7.2, y: 4.75, w: 5.0, h: 1.7,
        fontSize: 13, fontFace: 'Calibri', bold: true, color: 'FFFFFF',
        lineSpacing: 22
      });

      attachNotes(s5, slidesData[4].speakerNotes);
    }

    // --- SLIDE 6: OBJECT PRONOUNS ---
    {
      const s6 = createSlide();
      addHeader(s6, slidesData[5].category, slidesData[5].title, slidesData[5].subtitle);

      // Left Side: Table & Principle
      const objTable = [
        [{ text: "Subject Form (Doer)", options: { bold: true, fill: "312E81", color: "FFFFFF" } }, { text: "Object Form (Receiver)", options: { bold: true, fill: "064E3B", color: "34D399" } }],
        ["I", "me"],
        ["You", "you"],
        ["He", "him"],
        ["She", "her"],
        ["It", "it"],
        ["We", "us"],
        ["They", "them"]
      ];
      s6.addTable(objTable, {
        x: 0.8, y: 1.7, w: 5.6,
        colW: [2.8, 2.8],
        rowH: 0.42,
        fontSize: 11, fontFace: 'Arial', color: 'E2E8F0',
        border: { pt: '1', color: '334155' },
        fill: '1E293B'
      });

      s6.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 5.35, w: 5.6, h: 1.35,
        fill: { color: '1E293B' },
        line: { color: '818CF8', width: 1.5 },
        rectRadius: 0.1
      });
      s6.addText("🎯 THE CONTRAST & PLACEMENT PRINCIPLE:", {
        x: 1.0, y: 5.45, w: 5.2, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '818CF8'
      });
      s6.addText("• Subject = The one who DOES the action (Before Verb: S + V)\n• Object = The one who RECEIVES the action (After Verb / Preposition: V / Prep + O)\n• Sentence Anchor: “I love you” ➔ ‘you’ sits after the verb ‘love’ as Object.", {
        x: 1.0, y: 5.75, w: 5.2, h: 0.9,
        fontSize: 10.5, fontFace: 'Calibri', color: 'E2E8F0', lineSpacing: 15
      });

      // Right Side: Image Anna & Peter + Case study
      if (assets.anna_and_peter) {
        s6.addImage({
          data: assets.anna_and_peter,
          x: 6.7, y: 1.7, w: 5.83, h: 3.2,
          round: true
        });
      }
      s6.addShape(pptx.ShapeType.roundRect, {
        x: 6.7, y: 5.05, w: 5.83, h: 1.65,
        fill: { color: '064E3B' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.12
      });
      s6.addText("“She likes him.”", {
        x: 6.9, y: 5.2, w: 5.4, h: 0.5,
        fontSize: 18, fontFace: 'Arial', bold: true, color: 'FFFFFF'
      });
      s6.addText("• She = Initiator of affection (Subject, placed before verb)\n• him = Recipient of affection (Object, placed after transitive verb 'likes')", {
        x: 6.9, y: 5.75, w: 5.4, h: 0.8,
        fontSize: 12, fontFace: 'Calibri', color: 'A7F3D0'
      });

      attachNotes(s6, slidesData[5].speakerNotes);
    }

    // --- QUIZ SLIDES ---
    // Slide 7: Q1
    buildQuizSlide(slidesData[6], "CHECK 01", "Tom is my friend. ___ is very nice.", ["A. He", "B. Him", "C. His", "D. Himself"], "A", "'Tom' is the initiator / subject performing before the verb 'is' in the second clause. Therefore, Tom = Subject Pronoun 'He'.");

    // Slide 8: Q2
    buildQuizSlide(slidesData[7], "CHECK 02", "“Anna likes Peter.” — If we replace “Peter”, which sentence is correct?", ["A. Anna likes he.", "B. Anna likes him.", "C. Anna likes his.", "D. Anna likes himself."], "B", "Peter receives Anna's affection and sits after the transitive verb 'likes'. Therefore, Peter = Object Pronoun 'him'.");

    // --- SLIDE 9: POSSESSIVE PRONOUNS ---
    {
      const s9 = createSlide();
      addHeader(s9, slidesData[8].category, slidesData[8].title, slidesData[8].subtitle);

      const posTable = [
        [{ text: "Owner (Person)", options: { bold: true, fill: "312E81", color: "FFFFFF" } }, { text: "Possessive Pronoun", options: { bold: true, fill: "581C87", color: "F3E8FF" } }],
        ["I", "mine"],
        ["You", "yours"],
        ["He", "his"],
        ["She", "hers"],
        ["It", "its"],
        ["We", "ours"],
        ["They", "theirs"]
      ];
      s9.addTable(posTable, {
        x: 0.8, y: 1.7, w: 5.8,
        colW: [3.0, 2.8],
        rowH: 0.5,
        fontSize: 12, fontFace: 'Arial', color: 'E2E8F0',
        border: { pt: '1', color: '334155' },
        fill: '1E293B'
      });

      // Right Side Cards
      s9.addShape(pptx.ShapeType.roundRect, {
        x: 6.9, y: 1.7, w: 5.63, h: 2.2,
        fill: { color: '1E293B' },
        line: { color: 'C084FC', width: 1.5 },
        rectRadius: 0.12
      });
      s9.addText("🔑 THE STRATEGIC FUNCTION:", {
        x: 7.2, y: 1.9, w: 5.0, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: 'C084FC'
      });
      s9.addText("Directly Answers: “Whose is this?”", {
        x: 7.2, y: 2.25, w: 5.0, h: 0.5,
        fontSize: 18, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s9.addText("Used to declare ownership definitively without having to mention the noun a second time.", {
        x: 7.2, y: 2.85, w: 5.0, h: 0.8,
        fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1'
      });

      s9.addShape(pptx.ShapeType.roundRect, {
        x: 6.9, y: 4.15, w: 5.63, h: 2.55,
        fill: { color: '1E293B' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.12
      });
      s9.addText("EXECUTIVE DEMONSTRATION:", {
        x: 7.2, y: 4.35, w: 5.0, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: '34D399'
      });
      s9.addText("“This book is mine.”", {
        x: 7.2, y: 4.75, w: 5.0, h: 0.6,
        fontSize: 22, fontFace: 'Arial', bold: true, color: 'FFFFFF'
      });
      s9.addText("= This book belongs to me.\nThe word 'mine' stands completely alone at the end with NO following noun.", {
        x: 7.2, y: 5.4, w: 5.0, h: 1.0,
        fontSize: 12, fontFace: 'Calibri', color: 'CBD5E1'
      });

      attachNotes(s9, slidesData[8].speakerNotes);
    }

    // --- SLIDE 10: ADJECTIVE VS PRONOUN ---
    {
      const s10 = createSlide();
      addHeader(s10, slidesData[9].category, slidesData[9].title, slidesData[9].subtitle);

      // Left: Possessive Adjective
      s10.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.7, w: 5.7, h: 4.1,
        fill: { color: '1E293B' },
        line: { color: 'F59E0B', width: 1.5 },
        rectRadius: 0.12
      });
      s10.addText("CATEGORY A: POSSESSIVE ADJECTIVES", {
        x: 1.1, y: 1.95, w: 5.1, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s10.addText("my, your, his, her, its, our, their", {
        x: 1.1, y: 2.3, w: 5.1, h: 0.4,
        fontSize: 14, fontFace: 'Arial', bold: true, color: 'FFFFFF'
      });
      s10.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: 2.85, w: 5.1, h: 2.6,
        fill: { color: '0F172A' },
        line: { color: '334155', width: 1 },
        rectRadius: 0.08
      });
      s10.addText("⚠️ MANDATORY NOUN COLLOCATION:\n\nFormula:  TTSH + Noun / Noun Phrase\n\nExample:  This is my book.\n               My house is beautiful.\n\n➔ CANNOT stand alone without a noun!", {
        x: 1.3, y: 3.0, w: 4.7, h: 2.3,
        fontSize: 12.5, fontFace: 'Calibri', color: 'E2E8F0', lineSpacing: 17
      });

      // Right: Possessive Pronoun
      s10.addShape(pptx.ShapeType.roundRect, {
        x: 6.83, y: 1.7, w: 5.7, h: 4.1,
        fill: { color: '1E293B' },
        line: { color: '818CF8', width: 1.5 },
        rectRadius: 0.12
      });
      s10.addText("CATEGORY B: POSSESSIVE PRONOUNS", {
        x: 7.1, y: 1.95, w: 5.1, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: '818CF8'
      });
      s10.addText("mine, yours, his, hers, ours, theirs", {
        x: 7.1, y: 2.3, w: 5.1, h: 0.4,
        fontSize: 14, fontFace: 'Arial', bold: true, color: 'FFFFFF'
      });
      s10.addShape(pptx.ShapeType.roundRect, {
        x: 7.1, y: 2.85, w: 5.1, h: 2.6,
        fill: { color: '0F172A' },
        line: { color: '334155', width: 1 },
        rectRadius: 0.08
      });
      s10.addText("✅ STANDALONE SYNTACTIC ROLE:\n\nFormula:  STANDALONE (Replaces Adjective + Noun)\n\nExample:  The book is mine. (= my book)\n               The car is hers. (= her car)\n\n➔ Already replaces the noun completely!", {
        x: 7.3, y: 3.0, w: 4.7, h: 2.3,
        fontSize: 12.5, fontFace: 'Calibri', color: 'E2E8F0', lineSpacing: 17
      });

      // Bottom Rule Banner: Algebraic Equation
      s10.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 6.0, w: 11.73, h: 0.85,
        fill: { color: '1E1B4B' },
        line: { color: '6366F1', width: 1.5 },
        rectRadius: 0.1
      });
      s10.addText("💡 THE ALGEBRAIC LAW:   Possessive Pronoun = Possessive Adjective + Noun   (mine = my books)   |   TTSH + N/NP (My house is beautiful)", {
        x: 0.8, y: 6.0, w: 11.73, h: 0.85,
        fontSize: 12.5, fontFace: 'Arial', bold: true, color: 'FBBF24',
        align: 'center', valign: 'middle'
      });

      attachNotes(s10, slidesData[9].speakerNotes);
    }

    // Slide 11: Q3
    buildQuizSlide(slidesData[10], "CHECK 03", "This is ___ book.", ["A. mine", "B. my", "C. me", "D. I"], "B", "The noun 'book' immediately follows the blank. By the rule [Possessive Adjective + Noun], we must choose 'my'.");

    // Slide 12: Q4
    buildQuizSlide(slidesData[11], "CHECK 04", "This book is ___.", ["A. my", "B. me", "C. mine", "D. I"], "C", "There is NO noun after the blank at the end of the sentence. Therefore, we use the standalone possessive pronoun 'mine'.");

    // --- SLIDE 13: REFLEXIVE PRONOUNS ---
    {
      const s13 = createSlide();
      addHeader(s13, slidesData[12].category, slidesData[12].title, slidesData[12].subtitle);

      // Left Column: 3 Core Functions Card
      s13.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.7, w: 6.0, h: 2.35,
        fill: { color: '1E293B' },
        line: { color: 'C084FC', width: 1.5 },
        rectRadius: 0.12
      });
      s13.addText("🪞 3 DISTINCT CORE FUNCTIONS:", {
        x: 1.1, y: 1.85, w: 5.4, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: 'C084FC'
      });
      s13.addText("1. Subject = Object: “I love myself” (Action returns to doer)\n2. Emphatic: “She carries these books herself” (Intensifies doer)\n3. Idiom with ‘by’: “by myself / by yourself” = alone (unassisted)", {
        x: 1.1, y: 2.2, w: 5.4, h: 1.7,
        fontSize: 11.5, fontFace: 'Calibri', bold: true, color: 'FFFFFF', lineSpacing: 17
      });

      // Singular vs Plural Table
      const refTable = [
        [{ text: "Singular (-self)", options: { bold: true, fill: "312E81", color: "FFFFFF" } }, { text: "Plural (-selves)", options: { bold: true, fill: "581C87", color: "F3E8FF" } }],
        ["myself, yourself", "ourselves"],
        ["himself, herself, itself", "yourselves, themselves"]
      ];
      s13.addTable(refTable, {
        x: 0.8, y: 4.25, w: 6.0,
        colW: [3.0, 3.0],
        rowH: 0.55,
        fontSize: 12, fontFace: 'Arial', color: 'E2E8F0',
        border: { pt: '1', color: '334155' },
        fill: '1E293B'
      });

      // Right Column: Image
      if (assets.reflexive_mirror) {
        s13.addImage({
          data: assets.reflexive_mirror,
          x: 7.1, y: 1.7, w: 5.43, h: 5.0,
          round: true
        });
      }
      attachNotes(s13, slidesData[12].speakerNotes);
    }

    // Slide 14: Q5
    buildQuizSlide(slidesData[13], "CHECK 05", "John cut ___ while cooking.", ["A. him", "B. his", "C. himself", "D. he"], "C", "John performed the slicing action and John suffered the injury. The doer and receiver are the exact same person ➔ 'himself'.");

    // --- SLIDE 15: DEMONSTRATIVE PRONOUNS ---
    {
      const s15 = createSlide();
      addHeader(s15, slidesData[14].category, slidesData[14].title, slidesData[14].subtitle);

      // Left Column: 2x2 Matrix Table & Rules
      const demTable = [
        [{ text: "Quantity / Distance", options: { bold: true, fill: "312E81", color: "FFFFFF" } }, { text: "NEAR (Close to Speaker)", options: { bold: true, fill: "0369A1", color: "38BDF8" } }, { text: "FAR (Distant from Speaker)", options: { bold: true, fill: "B45309", color: "FBBF24" } }],
        ["Singular (1 item)", "THIS  (This is my phone)", "THAT  (That is my car)"],
        ["Plural (Multiple items)", "THESE  (These are my books)", "THOSE  (Those are my shoes)"]
      ];
      s15.addTable(demTable, {
        x: 0.8, y: 1.7, w: 6.0,
        colW: [2.0, 2.0, 2.0],
        rowH: 0.75,
        fontSize: 11, fontFace: 'Arial', color: 'E2E8F0',
        border: { pt: '1', color: '334155' },
        fill: '1E293B'
      });

      s15.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 4.4, w: 6.0, h: 2.3,
        fill: { color: '1E293B' },
        line: { color: '38BDF8', width: 1.5 },
        rectRadius: 0.12
      });
      s15.addText("⚡ THE TWO-AXIS COORDINATE & ADVANCED RULES:", {
        x: 1.1, y: 4.55, w: 5.4, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '38BDF8'
      });
      s15.addText("• Axis 1 (Distance): This / These = NEAR  |  That / Those = FAR\n• Axis 2 (Quantity): This / That = 1 (Singular)  |  These / Those = MANY (Plural)\n• Uncountable Nouns: Use THIS / THAT (e.g. this water, that advice)\n• Contextual Reference: 'That' frequently refers back to an identified idea.", {
        x: 1.1, y: 4.85, w: 5.4, h: 1.7,
        fontSize: 10.5, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 15
      });

      // Right Column: Image
      if (assets.spatial_pointers) {
        s15.addImage({
          data: assets.spatial_pointers,
          x: 7.1, y: 1.7, w: 5.43, h: 5.0,
          round: true
        });
      }
      attachNotes(s15, slidesData[14].speakerNotes);
    }

    // Slide 16: Q6
    buildQuizSlide(slidesData[15], "CHECK 06", "You are holding a book in your hand. You say: “___ is my book.”", ["A. Those", "B. These", "C. This", "D. That"], "C", "The book is held right in your hand (NEAR) and is 1 item (SINGULAR) ➔ 'This'.");

    // Slide 17: Q7
    buildQuizSlide(slidesData[16], "CHECK 07", "You see multiple cars parked far down the road. You say: “___ are cars.”", ["A. This", "B. That", "C. These", "D. Those"], "D", "Multiple vehicles (PLURAL) positioned in the distance (FAR) ➔ 'Those'.");

    // --- SLIDE 18: INDEFINITE PRONOUNS & SINGULAR VERB LAW ---
    {
      const s18 = createSlide();
      addHeader(s18, slidesData[17].category, slidesData[17].title, slidesData[17].subtitle);

      // Left Table: Indefinite Pronouns Matrix
      const indTable = [
        [{ text: "Suffix / Category", options: { bold: true, fill: "312E81", color: "FFFFFF" } }, { text: "Indefinite Pronouns", options: { bold: true, fill: "312E81", color: "38BDF8" } }, { text: "Agreement Law", options: { bold: true, fill: "312E81", color: "FBBF24" } }],
        ["-body / -one (People)", "everyone, someone, anyone, no one, everybody, somebody, nobody", "SINGULAR VERB"],
        ["-thing (Objects/Events)", "everything, something, anything, nothing", "SINGULAR VERB"],
        ["Distribution / Choice", "each, either, neither, another, each other", "SINGULAR VERB"]
      ];
      s18.addTable(indTable, {
        x: 0.8, y: 1.7, w: 6.2,
        colW: [1.8, 3.2, 1.2],
        rowH: 0.75,
        fontSize: 10.5, fontFace: 'Arial', color: 'E2E8F0',
        border: { pt: '1', color: '334155' },
        fill: '1E293B'
      });

      // Right Top: Singular Verb Law Banner
      s18.addShape(pptx.ShapeType.roundRect, {
        x: 7.3, y: 1.7, w: 5.23, h: 2.3,
        fill: { color: '1E293B' },
        line: { color: 'EF4444', width: 2 },
        rectRadius: 0.12
      });
      s18.addText("⚠️ THE CRUCIAL TOEIC LAW:", {
        x: 7.5, y: 1.85, w: 4.8, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: 'F87171'
      });
      s18.addText("Indefinite Pronoun + Singular Verb", {
        x: 7.5, y: 2.15, w: 4.8, h: 0.45,
        fontSize: 16, fontFace: 'Arial', bold: true, color: 'FFFFFF'
      });
      s18.addText("Formula: Vs/es  |  is  |  was  |  has\n\nEven though 'everyone/everybody' represents all people conceptually, English grammar commands a SINGULAR verb without exception!", {
        x: 7.5, y: 2.65, w: 4.8, h: 1.2,
        fontSize: 11, fontFace: 'Calibri', color: 'FCD34D', lineSpacing: 16
      });

      // Right Bottom: Example Demonstrations
      s18.addShape(pptx.ShapeType.roundRect, {
        x: 7.3, y: 4.15, w: 5.23, h: 2.55,
        fill: { color: '1E293B' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.12
      });
      s18.addText("EXECUTIVE SENTENCE DEMONSTRATIONS:", {
        x: 7.5, y: 4.35, w: 4.8, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: '34D399'
      });
      s18.addText("• Everyone has a vital role in our department.\n• Somebody is waiting at the executive lounge.\n• Nothing is impossible if we collaborate.\n• Each of the proposals was approved.", {
        x: 7.5, y: 4.7, w: 4.8, h: 1.8,
        fontSize: 11.5, fontFace: 'Calibri', bold: true, color: 'FFFFFF', lineSpacing: 19
      });

      attachNotes(s18, slidesData[17].speakerNotes);
    }

    // Slide 19: Check 08
    buildQuizSlide(slidesData[18], "CHECK 08", "“Everyone on the executive board ___ agreed to the proposal.”", ["A. have", "B. has", "C. are", "D. were"], "B", "Even though 'everyone' refers to multiple members conceptually, indefinite pronouns ending in -one, -body, -thing strictly take a SINGULAR verb. Between 'has' and 'have', 'has' is the singular form.");

    // --- SLIDE 20: INTERROGATIVE PRONOUNS ---
    {
      const s20 = createSlide();
      addHeader(s20, slidesData[19].category, slidesData[19].title, slidesData[19].subtitle);

      const intTable = [
        [{ text: "Pronoun", options: { bold: true, fill: "312E81", color: "FFFFFF" } }, { text: "Inquiry Target & Meaning", options: { bold: true, fill: "312E81", color: "38BDF8" } }],
        ["Who", "Person (Subject)  ➔  “Who is she?”"],
        ["Whom", "Person (Object, formal)  ➔  “Whom did you meet?”"],
        ["Whose", "Possession / Ownership  ➔  “Whose book is this?”"],
        ["What", "Thing, Event, Idea  ➔  “What is this?”"],
        ["Which", "Specific Choice among options  ➔  “Which do you like?”"]
      ];
      s20.addTable(intTable, {
        x: 0.8, y: 1.7, w: 6.2,
        colW: [1.8, 4.4],
        rowH: 0.65,
        fontSize: 11, fontFace: 'Arial', color: 'E2E8F0',
        border: { pt: '1', color: '334155' },
        fill: '1E293B'
      });

      // Right 4 Cards
      const intCards = [
        { y: 1.7, color: '38BDF8', word: "WHO", ex: "“Who is she?” (Person as Subject)" },
        { y: 2.95, color: 'C084FC', word: "WHOSE", ex: "“Whose phone is this?” (Investigating Ownership)" },
        { y: 4.2, color: '34D399', word: "WHAT", ex: "“What happened?” (Inquiring on Events/Things)" },
        { y: 5.45, color: 'FBBF24', word: "WHICH", ex: "“Which one do you prefer?” (Definite Choice)" }
      ];

      intCards.forEach(c => {
        s20.addShape(pptx.ShapeType.roundRect, {
          x: 7.3, y: c.y, w: 5.23, h: 1.1,
          fill: { color: '1E293B' },
          line: { color: c.color, width: 1.5 },
          rectRadius: 0.1
        });
        s20.addText(c.word, {
          x: 7.5, y: c.y + 0.15, w: 4.8, h: 0.3,
          fontSize: 13, fontFace: 'Arial', bold: true, color: c.color
        });
        s20.addText(c.ex, {
          x: 7.5, y: c.y + 0.45, w: 4.8, h: 0.5,
          fontSize: 12, fontFace: 'Calibri', color: 'FFFFFF'
        });
      });

      attachNotes(s20, slidesData[19].speakerNotes);
    }

    // Slide 21: Check 09
    buildQuizSlide(slidesData[20], "CHECK 09", "“___ is your best friend?”", ["A. What", "B. Who", "C. Whose", "D. Which"], "B", "The query investigates a human being ('best friend') as the subject of the clause ➔ 'Who'.");

    // Slide 22: Check 10
    buildQuizSlide(slidesData[21], "CHECK 10", "“___ phone is this?”", ["A. Who", "B. What", "C. Whose", "D. Which"], "C", "The query asks who owns the phone (Whose = Belonging to whom) ➔ 'Whose'.");

    // --- SLIDE 23: RELATIVE PRONOUNS CLAUSE CONNECTOR MATRIX ---
    {
      const s23 = createSlide();
      addHeader(s23, slidesData[22].category, slidesData[22].title, slidesData[22].subtitle);

      // Relative Pronouns Formula Matrix Table
      const relTable = [
        [{ text: "Antecedent (Noun)", options: { bold: true, fill: "312E81", color: "FFFFFF" } }, { text: "Connector", options: { bold: true, fill: "312E81", color: "38BDF8" } }, { text: "Followed By", options: { bold: true, fill: "312E81", color: "FBBF24" } }, { text: "Exam Pattern & Natural Demonstration", options: { bold: true, fill: "312E81", color: "34D399" } }],
        ["N (Person)", "WHO", "+ Verb (V)", "N(person) + WHO + V  ➔  The engineer who designed this app"],
        ["N (Person)", "WHOM", "+ Clause (S + V)", "N(person) + WHOM + S + V  ➔  The client whom we met yesterday"],
        ["N (Thing / Object)", "WHICH", "+ V / S + V", "N(thing) + WHICH + V / S + V  ➔  The contract which was signed"],
        ["N (Person / Thing)", "WHOSE", "+ Noun + V", "N + WHOSE + N + V  ➔  The author whose book won the global award"]
      ];
      s23.addTable(relTable, {
        x: 0.8, y: 1.7, w: 11.73,
        colW: [2.2, 1.5, 2.2, 5.83],
        rowH: 0.72,
        fontSize: 11, fontFace: 'Arial', color: 'E2E8F0',
        border: { pt: '1', color: '334155' },
        fill: '1E293B'
      });

      // Bottom Banner: Rapid TOEIC Formula
      s23.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 5.05, w: 11.73, h: 1.7,
        fill: { color: '1E1B4B' },
        line: { color: '6366F1', width: 2 },
        rectRadius: 0.12
      });
      s23.addText("⚡ TOEIC RAPID RECOGNITION FORMULA MATRIX:", {
        x: 1.1, y: 5.25, w: 11.13, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s23.addText("• Blank after Person + followed by Verb ➔ Pick WHO immediately!\n• Blank after Person + followed by Subject + Verb ➔ Pick WHOM (Object position)!\n• Blank between two Nouns declaring possession ➔ Pick WHOSE without hesitation!", {
        x: 1.1, y: 5.6, w: 11.13, h: 1.0,
        fontSize: 12, fontFace: 'Calibri', color: 'FFFFFF', lineSpacing: 18
      });

      attachNotes(s23, slidesData[22].speakerNotes);
    }

    // --- SLIDE 24: THAT RESTRICTIONS & PARTICIPLE REDUCTION ---
    {
      const s24 = createSlide();
      addHeader(s24, slidesData[23].category, slidesData[23].title, slidesData[23].subtitle);

      // Left: THAT Restrictions
      s24.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.7, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: 'EF4444', width: 2 },
        rectRadius: 0.12
      });
      s24.addText("🚫 THE 2 STRICT 'THAT' RESTRICTIONS", {
        x: 1.1, y: 1.9, w: 5.1, h: 0.3,
        fontSize: 12, fontFace: 'Arial', bold: true, color: 'F87171'
      });
      s24.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: 2.3, w: 5.1, h: 1.8,
        fill: { color: '0F172A' },
        line: { color: '334155', width: 1 },
        rectRadius: 0.08
      });
      s24.addText("1. COMMA PROHIBITION (Non-defining clauses):\n   ❌ Mr. David, that is our CFO, spoke today.\n   ✅ Mr. David, who is our CFO, spoke today.\n\n2. PREPOSITION PROHIBITION:\n   ❌ The company in that we invested...\n   ✅ The company in which we invested...", {
        x: 1.25, y: 2.4, w: 4.8, h: 1.6,
        fontSize: 11, fontFace: 'Calibri', color: 'E2E8F0', lineSpacing: 16
      });
      s24.addText("RULE: 'THAT' is versatile, but NEVER appears after a comma or a preposition in formal English/TOEIC!", {
        x: 1.1, y: 4.3, w: 5.1, h: 2.2,
        fontSize: 11.5, fontFace: 'Calibri', bold: true, color: 'FCD34D', lineSpacing: 18
      });

      // Right: Participle Clause Reduction
      s24.addShape(pptx.ShapeType.roundRect, {
        x: 6.83, y: 1.7, w: 5.7, h: 5.0,
        fill: { color: '1E293B' },
        line: { color: '38BDF8', width: 2 },
        rectRadius: 0.12
      });
      s24.addText("⚡ ADVANCED PARTICIPLE REDUCTION TRAP", {
        x: 7.1, y: 1.9, w: 5.1, h: 0.3,
        fontSize: 12, fontFace: 'Arial', bold: true, color: '38BDF8'
      });
      s24.addShape(pptx.ShapeType.roundRect, {
        x: 7.1, y: 2.3, w: 5.1, h: 1.8,
        fill: { color: '0F172A' },
        line: { color: '334155', width: 1 },
        rectRadius: 0.08
      });
      s24.addText("1. ACTIVE VOICE REDUCTION (V-ing):\n   “The man who lives next door...”\n   ➔ “The man living next door...”\n\n2. PASSIVE VOICE REDUCTION (V3/ed):\n   “The proposal which was submitted yesterday...”\n   ➔ “The proposal submitted yesterday...”", {
        x: 7.25, y: 2.4, w: 4.8, h: 1.6,
        fontSize: 11, fontFace: 'Calibri', color: 'E2E8F0', lineSpacing: 16
      });
      s24.addText("EXAM TRAP: When the relative pronoun is eliminated, the verb transforms into a participle (V-ing / V3), NOT an active finite verb!", {
        x: 7.1, y: 4.3, w: 5.1, h: 2.2,
        fontSize: 11.5, fontFace: 'Calibri', bold: true, color: '34D399', lineSpacing: 18
      });

      attachNotes(s24, slidesData[23].speakerNotes);
    }

    // Slide 25: Check 11
    buildQuizSlide(slidesData[24], "CHECK 11", "“The senior manager, ___ we met at the summit, approved the budget.”", ["A. that", "B. whom", "C. which", "D. whose"], "B", "Two critical clues: 1) Antecedent is a person ('senior manager'); 2) Following is a clause 'we met' (S + V) with a preceding comma. 'that' is strictly prohibited after commas, making 'whom' the correct object relative pronoun.");

    // --- SLIDE 26: ADVANCED COMPOUND PRONOUNS (WHOEVER & WHICHEVER) ---
    {
      const s26 = createSlide();
      addHeader(s26, slidesData[25].category, slidesData[25].title, slidesData[25].subtitle);

      // Left: WHOEVER Equation
      s26.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.7, w: 5.7, h: 4.0,
        fill: { color: '1E293B' },
        line: { color: 'F59E0B', width: 1.5 },
        rectRadius: 0.12
      });
      s26.addText("🌟 WHOEVER — THE TOEIC EQUATION", {
        x: 1.1, y: 1.9, w: 5.1, h: 0.3,
        fontSize: 12, fontFace: 'Arial', bold: true, color: 'FBBF24'
      });
      s26.addShape(pptx.ShapeType.roundRect, {
        x: 1.1, y: 2.3, w: 5.1, h: 1.1,
        fill: { color: '0F172A' },
        line: { color: 'F59E0B', width: 1 },
        rectRadius: 0.08
      });
      s26.addText("Whoever + V(sing)  =  Anyone who + V(sing)", {
        x: 1.1, y: 2.3, w: 5.1, h: 1.1,
        fontSize: 13, fontFace: 'Arial', bold: true, color: 'FFFFFF',
        align: 'center', valign: 'middle'
      });
      s26.addText("• Meaning: “Bất kỳ ai / Bất cứ ai mà...”\n• Example: “Whoever arrives first receives the VIP handbook.”\n• TOEIC Trap: 'Whoever' already embeds 'Anyone who'. Never say 'Anyone whoever'!", {
        x: 1.1, y: 3.6, w: 5.1, h: 1.9,
        fontSize: 11.5, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 18
      });

      // Right: WHICHEVER Choices
      s26.addShape(pptx.ShapeType.roundRect, {
        x: 6.83, y: 1.7, w: 5.7, h: 4.0,
        fill: { color: '1E293B' },
        line: { color: '38BDF8', width: 1.5 },
        rectRadius: 0.12
      });
      s26.addText("🌟 WHICHEVER — THE LIMITED CHOICE CONNECTOR", {
        x: 7.1, y: 1.9, w: 5.1, h: 0.3,
        fontSize: 12, fontFace: 'Arial', bold: true, color: '38BDF8'
      });
      s26.addShape(pptx.ShapeType.roundRect, {
        x: 7.1, y: 2.3, w: 5.1, h: 1.1,
        fill: { color: '0F172A' },
        line: { color: '38BDF8', width: 1 },
        rectRadius: 0.08
      });
      s26.addText("Selection from a Known Limited Set of Options", {
        x: 7.1, y: 2.3, w: 5.1, h: 1.1,
        fontSize: 12.5, fontFace: 'Arial', bold: true, color: '38BDF8',
        align: 'center', valign: 'middle'
      });
      s26.addText("• As Subject: “Whichever is cheaper will be selected.”\n• As Object: “Choose whichever you prefer.”\n• As Determiner: “Take whichever flight suits your schedule.”", {
        x: 7.1, y: 3.6, w: 5.1, h: 1.9,
        fontSize: 11.5, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 18
      });

      // Bottom Card
      s26.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 5.9, w: 11.73, h: 0.95,
        fill: { color: '1E1B4B' },
        line: { color: '6366F1', width: 1.5 },
        rectRadius: 0.1
      });
      s26.addText("⚡ GRAMMATICAL AGREEMENT: Both 'Whoever' and 'Whichever' resolve open conditional clauses and strictly command SINGULAR verb agreement!", {
        x: 0.8, y: 5.9, w: 11.73, h: 0.95,
        fontSize: 12.5, fontFace: 'Arial', bold: true, color: 'FBBF24',
        align: 'center', valign: 'middle'
      });

      attachNotes(s26, slidesData[25].speakerNotes);
    }

    // Slide 27: Check 12
    buildQuizSlide(slidesData[26], "CHECK 12", "“___ arrives at the conference venue first will receive a complimentary VIP pass.”", ["A. Anyone", "B. Whomever", "C. Whoever", "D. Which"], "C", "The blank acts as the subject of the clause 'arrives...'. Under the TOEIC equation, 'Whoever + V' = 'Anyone who + V'. 'Anyone' alone lacks the connecting relative pronoun.");

    // --- SLIDE 28: THE 3 GOLDEN QUESTIONS ---
    {
      const s28 = createSlide();
      addHeader(s28, slidesData[27].category, slidesData[27].title, slidesData[27].subtitle);

      const q3Cards = [
        { x: 0.8, color: '818CF8', qNum: "QUESTION 1", qTitle: "Who does the action?", res: "➔ SUBJECT PRONOUN", ex: "He plays football • She reads", desc: "Positioned before the verb as the initiator." },
        { x: 4.8, color: '38BDF8', qNum: "QUESTION 2", qTitle: "Who or what receives?", res: "➔ OBJECT PRONOUN", ex: "I like him • They help us", desc: "Positioned after the verb or preposition as target." },
        { x: 8.8, color: 'C084FC', qNum: "QUESTION 3", qTitle: "Bounces back to doer?", res: "➔ REFLEXIVE PRONOUN", ex: "He hurt himself • Cooked by myself", desc: "Used when doer and receiver are identical." }
      ];

      q3Cards.forEach(c => {
        s28.addShape(pptx.ShapeType.roundRect, {
          x: c.x, y: 1.7, w: 3.7, h: 5.0,
          fill: { color: '1E293B' },
          line: { color: c.color, width: 2 },
          rectRadius: 0.15
        });
        s28.addText(c.qNum, {
          x: c.x + 0.3, y: 2.0, w: 3.1, h: 0.3,
          fontSize: 11, fontFace: 'Arial', bold: true, color: c.color
        });
        s28.addText(c.qTitle, {
          x: c.x + 0.3, y: 2.4, w: 3.1, h: 0.8,
          fontSize: 18, fontFace: 'Arial', bold: true, color: 'FFFFFF', lineSpacing: 22
        });
        s28.addShape(pptx.ShapeType.roundRect, {
          x: c.x + 0.3, y: 3.4, w: 3.1, h: 0.6,
          fill: { color: '0F172A' },
          line: { color: c.color, width: 1 },
          rectRadius: 0.08
        });
        s28.addText(c.res, {
          x: c.x + 0.3, y: 3.4, w: 3.1, h: 0.6,
          fontSize: 12, fontFace: 'Arial', bold: true, color: c.color,
          align: 'center', valign: 'middle'
        });
        s28.addText(c.ex, {
          x: c.x + 0.3, y: 4.3, w: 3.1, h: 0.7,
          fontSize: 13, fontFace: 'Calibri', bold: true, color: 'FFFFFF', lineSpacing: 18
        });
        s28.addText(c.desc, {
          x: c.x + 0.3, y: 5.2, w: 3.1, h: 1.0,
          fontSize: 11, fontFace: 'Calibri', color: '94A3B8', lineSpacing: 16
        });
      });

      attachNotes(s28, slidesData[27].speakerNotes);
    }

    // --- SLIDE 29: ALL-IN-ONE MASTER MATRIX ---
    {
      const s29 = createSlide();
      addHeader(s29, slidesData[28].category, slidesData[28].title, slidesData[28].subtitle);

      const allTable = [
        [{ text: "Grammar Category", options: { bold: true, fill: "312E81", color: "FFFFFF" } }, { text: "Key English Pronouns", options: { bold: true, fill: "312E81", color: "38BDF8" } }, { text: "Sentence Function & Placement Rules", options: { bold: true, fill: "312E81", color: "FBBF24" } }],
        ["1. Subject Pronoun", "I, you, he, she, it, we, they", "The doer; stands before the main verb (S + V)."],
        ["2. Object Pronoun", "me, you, him, her, it, us, them", "The receiver; stands after verb or preposition (V/Prep + O)."],
        ["3. Possessive Pronoun", "mine, yours, his, hers, ours, theirs", "ĐTSH = TTSH + N; stands alone without noun."],
        ["4. Reflexive Pronoun", "myself, yourself, himself, ourselves, themselves", "S = O; emphatic; or 'by myself' = alone."],
        ["5. Demonstrative", "this, that, these, those", "Distance (Near/Far) & Quantity (Singular/Plural)."],
        ["6. Indefinite Pronoun", "everyone, someone, anything, each, another...", "Commands strictly SINGULAR verb agreement (Vs/es, is, has)."],
        ["7. Relative Pronoun", "who, whom, which, that, whose", "Connects clauses; 'THAT' banned after commas/prepositions."],
        ["8. Compound Pronoun", "whoever, whichever...", "Whoever = Anyone who + V(sing); Whichever for choices."]
      ];

      s29.addTable(allTable, {
        x: 0.8, y: 1.7, w: 11.73,
        colW: [2.5, 4.2, 5.03],
        rowH: 0.52,
        fontSize: 11, fontFace: 'Arial', color: 'E2E8F0',
        border: { pt: '1', color: '334155' },
        fill: '1E293B'
      });

      attachNotes(s29, slidesData[28].speakerNotes);
    }

    // --- SLIDE 30: GRAND CHALLENGE (SARAH'S LAPTOP) ---
    {
      const s30 = createSlide();
      addHeader(s30, slidesData[29].category, slidesData[29].title, slidesData[29].subtitle);

      // Left Image: Sarah
      if (assets.sarah_laptop) {
        s30.addImage({
          data: assets.sarah_laptop,
          x: 0.8, y: 1.7, w: 4.8, h: 5.0,
          round: true
        });
      }

      // Right Challenge Container
      // Question Card
      s30.addShape(pptx.ShapeType.roundRect, {
        x: 5.8, y: 1.7, w: 6.73, h: 1.25,
        fill: { color: '1E293B' },
        line: { color: 'EF4444', width: 2 },
        rectRadius: 0.12
      });
      s30.addText("GRAND CHALLENGE:", {
        x: 6.0, y: 1.85, w: 6.3, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: 'F87171'
      });
      s30.addText("Sarah has a new laptop. ___ laptop is very expensive, but the laptop is not ___.", {
        x: 6.0, y: 2.15, w: 6.3, h: 0.7,
        fontSize: 13, fontFace: 'Arial', bold: true, color: 'FFFFFF', lineSpacing: 20
      });

      // 4 Options in 2x2
      const sarahOpts = [
        { x: 5.8, y: 3.1, txt: "A.  Hers / her", corr: false },
        { x: 9.25, y: 3.1, txt: "B.  Her / hers  ✓", corr: true },
        { x: 5.8, y: 4.0, txt: "C.  She / her", corr: false },
        { x: 9.25, y: 4.0, txt: "D.  Her / she", corr: false }
      ];

      sarahOpts.forEach(opt => {
        s30.addShape(pptx.ShapeType.roundRect, {
          x: opt.x, y: opt.y, w: 3.28, h: 0.75,
          fill: { color: opt.corr ? '064E3B' : '1E293B' },
          line: { color: opt.corr ? '10B981' : '334155', width: opt.corr ? 2 : 1 },
          rectRadius: 0.08
        });
        s30.addText(opt.txt, {
          x: opt.x + 0.2, y: opt.y, w: 2.88, h: 0.75,
          fontSize: 12, fontFace: 'Arial', bold: opt.corr,
          color: opt.corr ? '34D399' : 'E2E8F0', valign: 'middle'
        });
      });

      // Explanation Box
      s30.addShape(pptx.ShapeType.roundRect, {
        x: 5.8, y: 4.95, w: 6.73, h: 1.75,
        fill: { color: '0F2922' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.1
      });
      s30.addText("ANSWER: B (Her / hers)  •  DUAL PARADIGM BREAKDOWN", {
        x: 6.0, y: 5.1, w: 6.3, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: '34D399'
      });
      s30.addText("1. 'Her laptop' ➔ Precedes noun 'laptop' ➔ Possessive Adjective (TTSH + N).\n2. 'not hers' ➔ Stands completely alone at sentence end ➔ Possessive Pronoun (hers = her laptop).", {
        x: 6.0, y: 5.45, w: 6.3, h: 1.1,
        fontSize: 11, fontFace: 'Calibri', color: 'CBD5E1', lineSpacing: 18
      });

      attachNotes(s30, slidesData[29].speakerNotes);
    }

    // --- SLIDE 31: RAPID FIRE ROUND ---
    {
      const s31 = createSlide();
      addHeader(s31, slidesData[30].category, slidesData[30].title, slidesData[30].subtitle);

      const rapids = [
        { q: "1.  Tom is a student. [  ___  ] studies English.", ans: "HE" },
        { q: "2.  I like Anna. I often talk to [  ___  ].", ans: "HER" },
        { q: "3.  This is my pen. The pen is [  ___  ].", ans: "MINE" },
        { q: "4.  He made the cake by [  ___  ].", ans: "HIMSELF" },
        { q: "5.  [  ___  ] are my shoes (pointing to shoes on feet).", ans: "THESE" }
      ];

      let rY = 1.7;
      rapids.forEach(r => {
        // Bar container
        s31.addShape(pptx.ShapeType.roundRect, {
          x: 0.8, y: rY, w: 11.73, h: 0.85,
          fill: { color: '1E293B' },
          line: { color: '334155', width: 1 },
          rectRadius: 0.1
        });
        // Sentence
        s31.addText(r.q, {
          x: 1.1, y: rY, w: 8.5, h: 0.85,
          fontSize: 13, fontFace: 'Arial', bold: true, color: 'FFFFFF',
          valign: 'middle'
        });
        // Answer Pill
        s31.addShape(pptx.ShapeType.roundRect, {
          x: 10.2, y: rY + 0.15, w: 2.0, h: 0.55,
          fill: { color: '064E3B' },
          line: { color: '10B981', width: 1.5 },
          rectRadius: 0.28
        });
        s31.addText("➔  " + r.ans, {
          x: 10.2, y: rY + 0.15, w: 2.0, h: 0.55,
          fontSize: 12, fontFace: 'Arial', bold: true, color: '34D399',
          align: 'center', valign: 'middle'
        });

        rY += 1.0;
      });

      attachNotes(s31, slidesData[30].speakerNotes);
    }

    // --- SLIDE 32: CONCLUSION & Q&A ---
    {
      const s32 = createSlide();
      addHeader(s32, slidesData[31].category, slidesData[31].title, slidesData[31].subtitle);

      // Quote Card
      s32.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.7, w: 11.73, h: 1.4,
        fill: { color: '1E1B4B' },
        line: { color: '6366F1', width: 2 },
        rectRadius: 0.15
      });
      s32.addText("“Pronouns may be small words, but they form the backbone of natural English.”", {
        x: 1.1, y: 1.85, w: 11.13, h: 0.5,
        fontSize: 18, fontFace: 'Arial', bold: true, color: 'FBBF24',
        align: 'center', valign: 'middle'
      });
      s32.addText("Pronouns appear in almost every sentence. By identifying who/what they replace and their grammatical position, natural English becomes effortless.", {
        x: 1.1, y: 2.45, w: 11.13, h: 0.5,
        fontSize: 11.5, fontFace: 'Calibri', color: 'CBD5E1',
        align: 'center'
      });

      // 6 Golden Takeaways
      s32.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 3.25, w: 11.73, h: 2.65,
        fill: { color: '1E293B' },
        line: { color: '38BDF8', width: 1.5 },
        rectRadius: 0.12
      });
      s32.addText("THE 6 GOLDEN RULES TO TAKE HOME:", {
        x: 1.1, y: 3.4, w: 11.13, h: 0.25,
        fontSize: 11.5, fontFace: 'Arial', bold: true, color: '38BDF8'
      });
      s32.addText("1. Pronoun = Replaces a noun to eliminate repetitive speech and sound executive.\n2. Subject Pronoun = The initiator positioned before the verb (He plays football).\n3. Object Pronoun = The recipient positioned after the verb or preposition (She likes him).\n4. Possessive Pronoun = Standalone ownership without a noun (mine = my books).\n5. Indefinite Pronoun = Strictly conjugated with SINGULAR verbs (Everyone has...).\n6. Relative Pronouns = Clause connectors; 'THAT' strictly forbidden after commas and prepositions.", {
        x: 1.1, y: 3.75, w: 11.13, h: 2.0,
        fontSize: 11.5, fontFace: 'Calibri', color: 'FFFFFF', lineSpacing: 18
      });

      // Q&A Banner
      s32.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 6.05, w: 11.73, h: 0.75,
        fill: { color: '064E3B' },
        line: { color: '10B981', width: 1.5 },
        rectRadius: 0.1
      });
      s32.addText("THANK YOU FOR YOUR ATTENTION!  •  Q&A SESSION NOW OPEN", {
        x: 0.8, y: 6.05, w: 11.73, h: 0.75,
        fontSize: 13, fontFace: 'Arial', bold: true, color: '34D399',
        align: 'center', valign: 'middle'
      });

      attachNotes(s32, slidesData[31].speakerNotes);
    }

    // Export PPTX and inject smooth slide transitions (Fade) via JSZip
    pptx.write({ outputType: 'blob' })
      .then(async (blob) => {
        let finalBlob = blob;
        if (window.JSZip) {
          try {
            const zip = await JSZip.loadAsync(blob);
            const slideRegex = /^ppt\/slides\/slide\d+\.xml$/;
            for (const [filename, file] of Object.entries(zip.files)) {
              if (slideRegex.test(filename)) {
                let xml = await file.async("string");
                if (!xml.includes("<p:transition")) {
                  // Standard OpenXML Slide Transition (Fade, medium speed)
                  const transXml = '<p:transition spd="med" advClick="1"><p:fade/></p:transition>';
                  if (xml.includes("</p:clrMapOvr>")) {
                    xml = xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + transXml);
                  } else if (xml.includes("</p:cSld>")) {
                    xml = xml.replace("</p:cSld>", "</p:cSld>" + transXml);
                  }
                  zip.file(filename, xml);
                }
              }
            }
            finalBlob = await zip.generateAsync({ type: "blob" });
          } catch (e) {
            console.warn("Could not inject transitions:", e);
          }
        }

        // Trigger browser download
        const url = URL.createObjectURL(finalBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'English_Pronouns_Masterclass.pptx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        btn.innerHTML = `<span>Đã tải thành công! ✅</span>`;
        setTimeout(() => {
          btn.innerHTML = originalText;
          btn.disabled = false;
        }, 3000);
      })
      .catch(err => {
        alert('Lỗi xuất PPTX: ' + err);
        btn.innerHTML = originalText;
        btn.disabled = false;
      });

  } catch (err) {
    alert('Lỗi tạo PPTX: ' + err.message);
    btn.innerHTML = originalText;
    btn.disabled = false;
  }
};
