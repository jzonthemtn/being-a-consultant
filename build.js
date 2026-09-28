const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const OUT = process.argv[2] || "going-independent.pptx";

// Palette: deep forest, warm amber accent
const C = {
  dark: "16211C", dark2: "1F2D26", ridge1: "1E2C25", ridge2: "243429",
  light: "F7F7F4", card: "EDEFEB", ink: "16211C", muted: "4A5550", mutedDark: "B9C7BE",
  amber: "E0A458", amberDeep: "A5541A", slate: "5C6962", white: "FBFBF8",
};
const HEAD = "Cambria";
const BODY = "Calibri";
const W = 13.333, H = 7.5, M = 0.7;

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

async function ridge() {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="420" viewBox="0 0 1920 420">
<path d="M0 300 L220 170 L380 260 L600 90 L820 240 L980 150 L1180 280 L1400 60 L1620 220 L1780 140 L1920 210 L1920 420 L0 420 Z" fill="#${C.ridge1}"/>
<path d="M0 360 L260 250 L470 330 L720 210 L960 320 L1200 230 L1460 330 L1700 250 L1920 310 L1920 420 L0 420 Z" fill="#${C.ridge2}"/></svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function num(slide, color) {
  slide.slideNumber = { x: W - M - 0.8, y: H - 0.6, w: 0.8, h: 0.3, fontFace: BODY, fontSize: 12, color, align: "right" };
}
function eyebrow(slide, text, color) {
  slide.addText(text.toUpperCase(), {
    x: M, y: 0.55, w: W - 2 * M, h: 0.4, fontFace: BODY, fontSize: 14, bold: true,
    color, charSpacing: 3, margin: 0, isTextBox: true,
  });
}
function title(slide, text, color, opts = {}) {
  slide.addText(text, {
    x: M, y: 0.95, w: W - 2 * M, h: opts.h || 0.9, fontFace: HEAD, fontSize: opts.size || 36,
    bold: true, color, margin: 0, valign: "top", isTextBox: true,
  });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "Jeff Zemerick";
  pres.title = "Going Independent";

  const ridgeImg = await ridge();
  const I = {
    search: await icon(fa.FaSearch, C.amber), chat: await icon(fa.FaComments, C.amber),
    cloud: await icon(fa.FaCloud, C.amber), lock: await icon(fa.FaLock, C.amber),
    sell: await icon(fa.FaHandshake, C.ink), deliver: await icon(fa.FaTools, C.ink),
    run: await icon(fa.FaFileInvoiceDollar, C.ink),
    code: await icon(fa.FaCodeBranch, C.amber), mic: await icon(fa.FaMicrophone, C.amber),
    pen: await icon(fa.FaPenNib, C.amber), trust: await icon(fa.FaHandshake, C.dark),
    work: await icon(fa.FaWrench, C.amberDeep), doc: await icon(fa.FaFileAlt, C.amberDeep),
    teach: await icon(fa.FaChalkboardTeacher, C.amberDeep), leave: await icon(fa.FaSignOutAlt, C.amberDeep),
    check: await icon(fa.FaCheckCircle, C.amberDeep),
    up: await icon(fa.FaThumbsUp, C.amberDeep), warn: await icon(fa.FaExclamationTriangle, C.slate),
  };

  // 1. Cover
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addImage({ data: ridgeImg, x: 0, y: H - 2.92, w: W, h: 2.92, altText: "" });
    s.addText("SEPTEMBER 30, 2026", {
      x: M, y: 1.5, w: 11, h: 0.4, fontFace: BODY, fontSize: 16, bold: true, color: C.amber,
      charSpacing: 3, margin: 0, isTextBox: true,
    });
    s.addText("Going Independent", {
      x: M, y: 2.0, w: 11.5, h: 1.3, fontFace: HEAD, fontSize: 66, bold: true, color: C.white,
      margin: 0, isTextBox: true,
    });
    s.addText("What it means to be an independent consultant", {
      x: M, y: 3.35, w: 10, h: 0.9, fontFace: BODY, fontSize: 24, color: C.mutedDark,
      margin: 0, valign: "top", isTextBox: true,
    });
    s.addText([
      { text: "Jeff Zemerick", options: { bold: true, color: C.white } },
      { text: "   ·   jeffzemerick.dev", options: { color: C.mutedDark } },
    ], { x: M, y: 4.35, w: 11, h: 0.5, fontFace: BODY, fontSize: 18, margin: 0, isTextBox: true });
    s.addNotes("Thanks for having me. Today I'll cover who I am, how I ended up self-employed, what the job actually looks like, how to find your niche, and what you can start doing now if consulting interests you. Then plenty of time for questions.");
  }

  // 2. About
  {
    const s = pres.addSlide();
    s.background = { color: C.light }; num(s, C.muted);
    eyebrow(s, "About me", C.amberDeep);
    title(s, "Engineer before consultant", C.ink);
    s.addText("WVU: B.S. 2006, M.S. 2008", { x: M, y: 1.7, w: 6, h: 0.35, fontFace: BODY, fontSize: 14, color: C.muted, margin: 0, isTextBox: true });
    const stats = [
      ["ASF", "Apache Software Foundation Member and Apache OpenNLP PMC Chair"],
      ["TSC", "Member of the OpenSearch Technical Steering Committee"],
      ["25", "cloud certifications (17 AWS, 8 Google Cloud)"],
    ];
    const gap = 0.35, cw = (W - 2 * M - 2 * gap) / 3, cy = 2.3, ch = 2.75;
    stats.forEach(([n, l], i) => {
      const x = M + i * (cw + gap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, fill: { color: C.card }, rectRadius: 0.12, line: { type: "none" } });
      s.addText(n, { x: x + 0.35, y: cy + 0.3, w: cw - 0.7, h: 1.1, fontFace: HEAD, fontSize: 60, bold: true, color: C.amberDeep, margin: 0, isTextBox: true });
      s.addText(l, { x: x + 0.35, y: cy + 1.45, w: cw - 0.7, h: 0.95, fontFace: BODY, fontSize: 18, color: C.muted, margin: 0, valign: "top", isTextBox: true });
    });
    s.addText([
      { text: "Go looking for the hard problems.", options: { bold: true, color: C.ink } },
      { text: " That\u2019s where you\u2019ll learn the most." },
    ], { x: M, y: 5.5, w: W - 2 * M, h: 1.0, fontFace: BODY, fontSize: 18, color: C.muted, margin: 0, valign: "top", isTextBox: true });
    s.addNotes("Quick background so you know where I'm coming from. I studied at WVU, a B.S. in 2006 and an M.S. in 2008, and I'm a software engineer at heart: distributed systems, search, and natural language processing. Open source has been a big part of my career. I'm a member of the Apache Software Foundation, chair of the Apache OpenNLP project, and on the OpenSearch Technical Steering Committee. The takeaway on this slide is for you: go looking for the hard problems. That's where you'll learn the most, and it's what clients will eventually pay you for.");
  }

  // 3. Journey
  {
    const s = pres.addSlide();
    s.background = { color: C.light }; num(s, C.muted);
    eyebrow(s, "My journey", C.amberDeep);
    title(s, "A non-linear path", C.ink);
    const steps = [
      ["AGE ~15", "Started a company", "A software company I worked through college.", C.amberDeep],
      ["AFTER COLLEGE", "The employee years", "NASA, FBI, Leidos, etc. Always working to get back to self-employed.", C.slate],
      ["2018", "Self-employed again", "Started taking on independent work.", C.amberDeep],
      ["2019", "Fully independent", "Consulting became the whole job.", C.amberDeep],
      ["TODAY", "Consulting today", "Search, NLP, and cloud consulting.", C.amberDeep],
    ];
    const gap = 0.3, cw = (W - 2 * M - 4 * gap) / 5, ly = 2.75, dot = 0.34;
    s.addShape(pres.shapes.LINE, { x: M + dot / 2, y: ly, w: 4 * (cw + gap), h: 0, line: { color: "C9C2B2", width: 2 } });
    steps.forEach(([when, head, body, col], i) => {
      const x = M + i * (cw + gap);
      s.addShape(pres.shapes.OVAL, { x, y: ly - dot / 2, w: dot, h: dot, fill: { color: col }, line: { color: C.light, width: 3 } });
      s.addText(when, { x, y: 3.25, w: cw, h: 0.35, fontFace: BODY, fontSize: 14, bold: true, color: col, charSpacing: 2, margin: 0, isTextBox: true });
      s.addText(head, { x, y: 3.65, w: cw, h: 0.85, fontFace: HEAD, fontSize: 20, bold: true, color: C.ink, margin: 0, valign: "top", isTextBox: true });
      s.addText(body, { x, y: 4.6, w: cw, h: 1.6, fontFace: BODY, fontSize: 15, color: C.muted, margin: 0, valign: "top", isTextBox: true });
    });
    s.addText("Working at \u201Creal\u201D jobs provides valuable experience and knowledge.", {
      x: M, y: 6.15, w: W - 2 * M - 1.0, h: 0.45, fontFace: BODY, fontSize: 16, italic: true, color: C.muted, margin: 0, valign: "middle", isTextBox: true,
    });
    s.addNotes("I caught the self-employment bug early. I started a software company around age 15 and worked it through college. After that I took the employee route: NASA, the FBI, Leidos, and others. Those were great projects, but the goal was always to get back to working for myself. I started doing independent work again in 2018 and went fully independent in 2019. Today that's search, NLP, and cloud consulting. The point for you: your path doesn't have to be linear. Those employee years gave me experience and knowledge I use with clients every day. [Optional: add what made 2018 the right time.]");
  }

  // 4. Definition (accent statement)
  {
    const s = pres.addSlide();
    s.background = { color: C.amber }; num(s, C.ink);
    eyebrow(s, "What it means", C.ink);
    s.addText("An independent consultant is a business of one.", {
      x: M, y: 1.0, w: 11, h: 1.8, fontFace: HEAD, fontSize: 48, bold: true, color: C.ink, margin: 0, valign: "top", isTextBox: true,
    });
    const cols = [
      [I.sell, "You sell it", "Find clients, understand their problem, scope the work, set your rate."],
      [I.deliver, "You deliver it", "Do the expert work, and be the person accountable for the result."],
      [I.run, "You run it", "Contracts, invoices, taxes, insurance, and your own training."],
    ];
    const gap = 0.45, cw = (W - 2 * M - 2 * gap) / 3;
    cols.forEach(([img, head, body], i) => {
      const x = M + i * (cw + gap);
      s.addShape(pres.shapes.OVAL, { x, y: 3.55, w: 0.8, h: 0.8, fill: { color: "F2C98E" }, line: { type: "none" } });
      s.addImage({ data: img, x: x + 0.2, y: 3.75, w: 0.4, h: 0.4, altText: "" });
      s.addText(head, { x, y: 4.6, w: cw, h: 0.55, fontFace: HEAD, fontSize: 26, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(body, { x, y: 5.2, w: cw, h: 1.1, fontFace: BODY, fontSize: 17, color: "2A2A22", margin: 0, valign: "top", isTextBox: true });
    });
    s.addNotes("This is the one idea I want you to take away. When you're independent, being great at the technical work is only one third of the job. You also have to sell it: find clients, understand their problem, scope the work, and set your rate. And you have to run it: contracts, invoices, taxes, insurance, and your own training. Nobody hands you a project, and nobody chases an unpaid invoice for you.");
  }

  // 5. Services
  {
    const s = pres.addSlide();
    s.background = { color: C.dark }; num(s, C.mutedDark);
    eyebrow(s, "What I actually do", C.amber);
    title(s, "Specialize so clients know when to call you", C.white);
    const items = [
      [I.search, "Search engineering", "Relevance tuning, OpenSearch, migrations"],
      [I.chat, "NLP and AI", "Language models, text processing, ML pipelines"],
      [I.cloud, "Cloud and data", "AWS and GCP architecture, streaming data engineering"],
      [I.lock, "Governance and compliance", "PII redaction, HIPAA, secure pipelines"],
    ];
    const gap = 0.35, cw = (W - 2 * M - gap) / 2, ch = 1.9;
    items.forEach(([img, head, body], i) => {
      const x = M + (i % 2) * (cw + gap), y = 2.2 + Math.floor(i / 2) * (ch + gap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: C.dark2 }, rectRadius: 0.12, line: { type: "none" } });
      s.addShape(pres.shapes.OVAL, { x: x + 0.35, y: y + 0.45, w: 0.75, h: 0.75, fill: { color: "2E4038" }, line: { type: "none" } });
      s.addImage({ data: img, x: x + 0.53, y: y + 0.63, w: 0.39, h: 0.39, altText: "" });
      s.addText(head, { x: x + 1.35, y: y + 0.42, w: cw - 1.6, h: 0.5, fontFace: HEAD, fontSize: 22, bold: true, color: C.white, margin: 0, isTextBox: true });
      s.addText(body, { x: x + 1.35, y: y + 0.98, w: cw - 1.6, h: 0.6, fontFace: BODY, fontSize: 16, color: C.mutedDark, margin: 0, valign: "top", isTextBox: true });
    });
    s.addNotes("Here's what I actually do. Notice it's not \"AI\" in general. It's search engineering, NLP and AI, the cloud and data work those depend on, and the governance and compliance around them. Being specific makes it easy for people to know when to call you. If someone has a search relevance problem, they think of me.");
  }

  // 6. Trade-offs
  {
    const s = pres.addSlide();
    s.background = { color: C.light }; num(s, C.muted);
    eyebrow(s, "The trade-offs", C.amberDeep);
    title(s, "Being independent comes with responsibility", C.ink);
    const cols = [
      [I.up, "What's great", C.amberDeep, "FFFFFF", [
        "You choose your clients and projects",
        "You set your own schedule",
        "You learn across many industries",
        "Your name is the brand",
        "Freedom in clients, schedule, and focus",
      ]],
      [I.warn, "What's hard", C.slate, C.card, [
        "Income can be irregular",
        "No paid vacation *",
        "Staying current requires effort",
        "Change can be difficult",
        "Continually looking for work",
      ]],
    ];
    const gap = 0.4, cw = (W - 2 * M - gap) / 2, y = 2.2, ch = 4.0;
    cols.forEach(([img, head, col, fill, lines], i) => {
      const x = M + i * (cw + gap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: fill }, rectRadius: 0.12, line: { color: "DDD8CC", width: 1 } });
      s.addImage({ data: img, x: x + 0.45, y: y + 0.45, w: 0.4, h: 0.4, altText: "" });
      s.addText(head, { x: x + 1.05, y: y + 0.4, w: cw - 1.5, h: 0.5, fontFace: HEAD, fontSize: 24, bold: true, color: col, margin: 0, isTextBox: true });
      s.addText(lines.map((t, j) => ({ text: t, options: { bullet: { indent: 18 }, breakLine: j < lines.length - 1, paraSpaceAfter: 14 } })), {
        x: x + 0.45, y: y + 1.2, w: cw - 0.9, h: ch - 1.5, fontFace: BODY, fontSize: 19, color: C.ink, margin: 0, valign: "top", isTextBox: true,
      });
    });
    s.addText("* You can pay yourself for time off and track PTO, if that\u2019s important to you.", {
      x: M, y: 6.4, w: W - 2 * M - 1.0, h: 0.4, fontFace: BODY, fontSize: 14, italic: true, color: C.muted, margin: 0, valign: "middle", isTextBox: true,
    });
    s.addNotes("I want to be honest about both columns. The good side is real: you choose your clients and projects, you set your schedule, you learn across a lot of industries, and your name is the brand. That flexibility is also how I have time for things like wilderness guiding. But the hard side is real too. Income can be irregular. There's no paid vacation by default, though you can pay yourself for time off and track PTO if that matters to you. Staying current takes deliberate effort, change can be difficult, and you're always looking for the next piece of work, even while delivering the current one.");
  }

  // 7. Visibility
  {
    const s = pres.addSlide();
    s.background = { color: C.dark }; num(s, C.mutedDark);
    eyebrow(s, "Building a reputation", C.amber);
    title(s, "Doing the work in public", C.white);
    const steps = [
      [I.code, "Open source", "Apache OpenNLP PMC Chair, OpenSearch TSC member and maintainer", false],
      [I.mic, "Speaking", "OpenSearchCon, Community Over Code, Open Source Summit, PyData", false],
      [I.pen, "Writing", "Blog posts, plus 25 cloud certifications as proof of skill", false],
      [I.trust, "Trust", "Clients can check your work before they ever call you", true],
    ];
    const arrow = 0.35, gap = 0.12, cw = (W - 2 * M - 3 * (arrow + 2 * gap)) / 4, y = 2.25, ch = 3.0;
    steps.forEach(([img, head, body, hi], i) => {
      const x = M + i * (cw + arrow + 2 * gap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: hi ? C.amber : C.dark2 }, rectRadius: 0.12, line: { type: "none" } });
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fill: { color: hi ? "F2C98E" : "2E4038" }, line: { type: "none" } });
      s.addImage({ data: img, x: x + 0.47, y: y + 0.47, w: 0.36, h: 0.36, altText: "" });
      s.addText(head, { x: x + 0.3, y: y + 1.2, w: cw - 0.6, h: 0.5, fontFace: HEAD, fontSize: 20, bold: true, color: hi ? C.ink : C.white, margin: 0, isTextBox: true });
      s.addText(body, { x: x + 0.3, y: y + 1.75, w: cw - 0.6, h: 1.1, fontFace: BODY, fontSize: 15, color: hi ? C.ink : C.mutedDark, margin: 0, valign: "top", isTextBox: true });
      if (i < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + cw + gap, y: y + ch / 2 - 0.15, w: arrow, h: 0.3, fill: { color: "6C7A72" }, line: { type: "none" } });
    });
    s.addText([
      { text: "As an independent, you have no company brand to lean on. ", options: { color: C.mutedDark } },
      { text: "Public work becomes your portfolio.", options: { color: C.white, bold: true } },
    ], { x: M, y: 5.6, w: W - 2 * M, h: 0.6, fontFace: BODY, fontSize: 19, margin: 0, isTextBox: true });
    s.addText([
      { text: "At this point, my website " , options: { color: C.mutedDark } },
      { text: "jeffzemerick.dev", options: { color: C.amber, hyperlink: { url: "https://jeffzemerick.dev" } } },
      { text: " is my resume.", options: { color: C.mutedDark } },
    ], { x: M, y: 6.2, w: W - 2 * M - 1.0, h: 0.45, fontFace: BODY, fontSize: 16, italic: true, margin: 0, isTextBox: true });
    s.addNotes("Without a big company name behind you, people need a way to trust you. For me that's been open source, conference talks, writing, and certifications. Anyone can look at my commits or watch a talk before we ever speak. Public work becomes your portfolio. At this point, my website is effectively my resume. I point people there instead of sending a CV.");
  }

  // 7b. Niche
  {
    const s = pres.addSlide();
    s.background = { color: C.light }; num(s, C.muted);
    eyebrow(s, "Finding your niche", C.amberDeep);
    title(s, "Know your angle", C.ink);
    s.addText("Mine: empower companies by getting them past a hard problem.", {
      x: M, y: 1.75, w: W - 2 * M, h: 0.6, fontFace: HEAD, fontSize: 24, italic: true, color: C.amberDeep, margin: 0, valign: "top", isTextBox: true,
    });
    const steps = [
      [I.work, "Do the work", "Solve the hard problem hands-on, alongside their team."],
      [I.doc, "Document it", "Leave behind clear write-ups so nothing lives only in my head."],
      [I.teach, "Teach their staff", "Make sure their people can own and extend it."],
      [I.leave, "Move on", "Leave them capable and independent, not dependent on me."],
    ];
    const arrow = 0.4, gap = 0.2, cw = (W - 2 * M - 3 * (arrow + 2 * gap)) / 4, cy = 2.55, d = 0.8;
    steps.forEach(([img, head, body], i) => {
      const x = M + i * (cw + arrow + 2 * gap);
      s.addShape(pres.shapes.OVAL, { x, y: cy, w: d, h: d, fill: { color: C.card }, line: { type: "none" } });
      s.addImage({ data: img, x: x + 0.2, y: cy + 0.2, w: 0.4, h: 0.4, altText: "" });
      s.addText(head, { x, y: cy + 1.05, w: cw, h: 0.5, fontFace: HEAD, fontSize: 22, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(body, { x, y: cy + 1.6, w: cw, h: 1.0, fontFace: BODY, fontSize: 16, color: C.muted, margin: 0, valign: "top", isTextBox: true });
      if (i < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + d + (cw + arrow + 2 * gap - d - arrow) / 2, y: cy + d / 2 - 0.14, w: arrow, h: 0.28, fill: { color: "C9C2B2" }, line: { type: "none" } });
    });
    s.addText("Find how you like to work and build that into your consulting, whether that means short engagements like mine, long-term partnerships, or something in between. It sounds counterintuitive, but short, successful engagements keep clients coming back for more.", {
      x: M, y: 5.6, w: W - 2 * M, h: 0.8, fontFace: BODY, fontSize: 14, italic: true, color: C.muted, margin: 0, valign: "top", isTextBox: true,
    });
    s.addNotes("Every independent consultant needs an angle: not just what you know, but how you like to work with clients. Mine is to empower companies by getting them past a hard problem. I do the work, document it, teach their staff, and move on, usually in three to six months. It sounds counterintuitive, but short, successful engagements keep clients coming back for more. Your angle might look different, like long-term partnerships, and that's fine. Figure out how you like to work and build your consulting around it. [Optional: share a short example engagement, without naming the client.]");
  }

  // 8b. Niche ideas
  {
    const s = pres.addSlide();
    s.background = { color: C.dark }; num(s, C.mutedDark);
    eyebrow(s, "Picking a niche", C.amber);
    title(s, "\u201CAI\u201D is too broad. Get specific.", C.white);
    const lw = 8.0, rx = M + lw + 0.3, rw = W - M - rx;
    const label = (t, x, w) => s.addText(t.toUpperCase(), { x, y: 1.95, w, h: 0.3, fontFace: BODY, fontSize: 13, bold: true, color: C.amber, charSpacing: 2, margin: 0, isTextBox: true });
    label("Specific problems", M, lw);
    label("Specific industries", rx, rw);
    const probs = [
      ["Search and retrieval (RAG)", "Help AI find and cite the right documents"],
      ["LLM evaluation", "Prove AI features are accurate before launch"],
      ["Privacy and PII", "Keep sensitive data out of prompts and logs"],
      ["Governance and compliance", "HIPAA, financial rules, the EU AI Act"],
      ["Document processing", "Extract data from contracts, claims, forms"],
      ["Agents and integration", "Connect AI safely to existing systems"],
      ["Cost and performance", "Smaller models, caching, self-hosting"],
      ["AI security", "Prompt injection, data leaks, misuse"],
    ];
    const g = 0.15, cw = (lw - g) / 2, ch = 0.72, y0 = 2.4;
    probs.forEach(([h, d], i) => {
      const x = M + (i % 2) * (cw + g), y = y0 + Math.floor(i / 2) * (ch + g);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: C.dark2 }, rectRadius: 0.08, line: { type: "none" } });
      s.addText(h, { x: x + 0.2, y: y + 0.08, w: cw - 0.4, h: 0.3, fontFace: HEAD, fontSize: 15, bold: true, color: C.white, margin: 0, isTextBox: true });
      s.addText(d, { x: x + 0.2, y: y + 0.39, w: cw - 0.4, h: 0.26, fontFace: BODY, fontSize: 13, color: C.mutedDark, margin: 0, isTextBox: true });
    });
    const gridH = 4 * ch + 3 * g;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rx, y: y0, w: rw, h: gridH, fill: { color: C.dark2 }, rectRadius: 0.08, line: { type: "none" } });
    const inds = ["Healthcare", "Legal", "Finance", "Insurance", "Government", "Manufacturing"];
    s.addText(inds.map((t, j) => ({ text: t, options: { bullet: { indent: 16 }, breakLine: j < inds.length - 1, paraSpaceAfter: 4 } })), {
      x: rx + 0.3, y: y0 + 0.22, w: rw - 0.6, h: 2.0, fontFace: BODY, fontSize: 16, color: C.white, margin: 0, valign: "top", isTextBox: true,
    });
    s.addText("Knowing an industry\u2019s vocabulary and regulations is often worth more than the technical skill.", {
      x: rx + 0.3, y: y0 + gridH - 0.95, w: rw - 0.6, h: 0.8, fontFace: BODY, fontSize: 13, italic: true, color: C.mutedDark, margin: 0, valign: "bottom", isTextBox: true,
    });
    s.addText([
      { text: "Pick two and put them together to find specific problems: ", options: { bold: true, color: C.white } },
      { text: "\u201Cdocument extraction for insurance claims\u201D", options: { color: C.mutedDark } },
    ], { x: M, y: 6.15, w: W - 2 * M - 1.0, h: 0.5, fontFace: BODY, fontSize: 16, margin: 0, valign: "middle", isTextBox: true });
    s.addNotes("Building on the idea of an angle: what does specific actually look like in AI? On the left are problem areas companies are paying for right now: search and retrieval for AI assistants, evaluating whether LLM features actually work, keeping private data out of AI pipelines, compliance, document extraction, agent integration, cost tuning, and AI security. On the right are industries. Pick one from each and put them together. \"AI consultant\" is hard to remember or refer. \"The person who does document extraction for insurance claims\" is easy. And knowing an industry's vocabulary and regulations is often worth more than the technical skill.");
  }

  // 8. Advice
  {
    const s = pres.addSlide();
    s.background = { color: C.light }; num(s, C.muted);
    eyebrow(s, "If you want to consult in AI", C.amberDeep);
    title(s, "Four things I'd tell my younger self", C.ink);
    const tips = [
      ["Learn the unglamorous parts.", "Data, cloud, security and compliance decide whether AI projects ship."],
      ["Work in the open.", "Contribute to open source, write, and give talks."],
      ["Find your focus,", "but keep your knowledge broad and be open."],
      ["Going back and forth is ok.", "Taking a full-time role between independent stretches is part of the path."],
    ];
    const y0 = 2.15, rh = 0.95;
    tips.forEach(([b, rest], i) => {
      const y = y0 + i * rh;
      s.addShape(pres.shapes.OVAL, { x: M, y: y + (rh - 0.56) / 2, w: 0.56, h: 0.56, fill: { color: C.amberDeep }, line: { type: "none" } });
      s.addText(String(i + 1), { x: M, y: y + (rh - 0.56) / 2, w: 0.56, h: 0.56, fontFace: HEAD, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText([
        { text: b + " ", options: { bold: true, color: C.ink } },
        { text: rest, options: { color: C.muted } },
      ], { x: M + 0.85, y, w: W - 2 * M - 0.85, h: rh, fontFace: BODY, fontSize: 20, valign: "middle", margin: 0, isTextBox: true });
    });
    s.addNotes("Four things I'd tell my younger self. First, learn the unglamorous parts. The model is rarely the hard part; the data, cloud, security, and compliance around it decide whether AI projects actually ship. Second, work in the open: contribute to open source, write, and give talks. Third, find your focus, but keep your knowledge broad and stay open to new things. The field changes fast. Fourth, going back and forth between independent work and full-time jobs is ok. It's not a failure. I did it myself, and each move taught me something.");
  }

  // 8c. Get started
  {
    const s = pres.addSlide();
    s.background = { color: C.light }; num(s, C.muted);
    eyebrow(s, "Getting started", C.amberDeep);
    title(s, "What you can do today", C.ink);
    const cols = [
      ["This week", [
        "Think about your interests and your focus",
        "Create a simple personal website that will serve as your resume",
      ]],
      ["This month", [
        "Make your first open-source contribution: docs, a test, or a small bug fix",
        "Write a short blog post about something you learned while doing it",
      ]],
      ["This semester", [
        "Build a small project in your niche and push it to GitHub",
        "Go to a meetup or conference and talk to people; make connections online and in-person",
      ]],
    ];
    const gap = 0.35, cw = (W - 2 * M - 2 * gap) / 3, cy = 1.95, ch = 3.7;
    cols.forEach(([when, items], i) => {
      const x = M + i * (cw + gap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, fill: { color: i === 0 ? "FFFFFF" : C.card }, rectRadius: 0.12, line: { color: "DDD8CC", width: 1 } });
      s.addText(when.toUpperCase(), { x: x + 0.35, y: cy + 0.35, w: cw - 0.7, h: 0.35, fontFace: BODY, fontSize: 15, bold: true, color: C.amberDeep, charSpacing: 2, margin: 0, isTextBox: true });
      items.forEach((t, j) => {
        const iy = cy + 0.95 + j * 1.35;
        s.addImage({ data: I.check, x: x + 0.35, y: iy + 0.04, w: 0.3, h: 0.3, altText: "" });
        s.addText(t, { x: x + 0.8, y: iy, w: cw - 1.15, h: 1.15, fontFace: BODY, fontSize: 17, color: C.ink, margin: 0, valign: "top", isTextBox: true });
      });
    });
    s.addText([
      { text: "You don\u2019t need permission or a job title to start doing the work in public.", options: { breakLine: true } },
      { text: "An open-source project or an industry conference isn\u2019t going to ask if you have a degree.", options: { breakLine: true } },
      { text: "Think of yourself as an expert. If you don\u2019t, others won\u2019t either." },
    ], {
      x: M, y: 5.95, w: W - 2 * M - 1.0, h: 0.95, valign: "top", fontFace: BODY, fontSize: 16, italic: true, color: C.muted, margin: 0, isTextBox: true,
    });
    s.addNotes("If consulting interests you, you can start building toward it now, while you're still a student. This week: think about your interests and where you want to focus, and create a simple personal website that will serve as your resume. This month: make your first open-source contribution. It doesn't have to be code; docs and tests are valuable and a great way in. Then write a short blog post about what you learned doing it. This semester: build a small project in your niche and push it to GitHub, and make connections at meetups, conferences, and online. You don't need permission or a job title to do any of this. An open-source project or a conference isn't going to ask if you have a degree. And think of yourself as an expert. If you don't, others won't either.");
  }

  // 9. Questions
  {
    const s = pres.addSlide();
    s.background = { color: C.dark }; num(s, C.mutedDark);
    s.addImage({ data: ridgeImg, x: 0, y: H - 2.92, w: W, h: 2.92, altText: "" });
    s.addText("Questions?", { x: M, y: 1.0, w: 11, h: 1.3, fontFace: HEAD, fontSize: 66, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText([
      { text: "Helping others matters to me. If you have questions later, feel free to email me or send me a message on " },
      { text: "LinkedIn", options: { color: C.amber, hyperlink: { url: "https://www.linkedin.com/in/jeffzemerick/" } } },
      { text: "." },
    ], { x: M, y: 2.4, w: 9.5, h: 0.85, fontFace: BODY, fontSize: 20, color: C.mutedDark, margin: 0, valign: "top", isTextBox: true });
    s.addText("Jeff Zemerick", { x: M, y: 3.5, w: 11, h: 0.5, fontFace: BODY, fontSize: 24, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText("jeffzemerick.dev", { x: M, y: 4.0, w: 11, h: 0.5, fontFace: BODY, fontSize: 22, color: C.amber, margin: 0, isTextBox: true, hyperlink: { url: "https://jeffzemerick.dev" } });
    s.addText("jeff.zemerick@mtnfog.com", { x: M, y: 4.5, w: 11, h: 0.5, fontFace: BODY, fontSize: 22, color: C.amber, margin: 0, isTextBox: true, hyperlink: { url: "mailto:jeff.zemerick@mtnfog.com" } });
    s.addNotes("Helping others matters to me, so please reach out after today too, by email or on LinkedIn. Happy to talk about anything: finding clients, setting rates, what AI projects look like in practice, open source, or the path from student to consultant.");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
