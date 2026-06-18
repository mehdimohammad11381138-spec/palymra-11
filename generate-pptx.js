const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// ── Color Palette (matching the example) ──
const C = {
  darkNavy: "3b4162",
  darkNavy2: "2d3250",
  warmBeige: "f5efe8",
  warmBeigeDarker: "ede4d8",
  teal: "5a9e8f",
  tealDark: "4a8a7b",
  headingNavy: "2d3250",
  bodyText: "4a4a5a",
  bodyMuted: "7a7a8a",
  accent: "c47a33",
  accentLight: "dda85a",
  salmon: "d4826a",
  white: "FFFFFF",
  offWhite: "f9f4ee",
  lightGray: "e8e0d6",
  danger: "c0392b",
};

// ── Font definitions ──
const FONT_SERIF = "Georgia";
const FONT_SANS = "Calibri";

// ── Image paths ──
const IMG = {
  colonnade: path.resolve("images/palmyra-colonnade-sunset.png"),
  theatre: path.resolve("images/palmyra-roman-theatre.png"),
  beforeAfter: path.resolve("images/palmyra-before-after-photo.png"),
  satellite: path.resolve("images/palmyra-satellite-comparison.png"),
  lion: path.resolve("images/palmyra-lion-sculpture.png"),
  stickers: path.resolve("images/stickers-strip.png"),
};

// Check images exist
for (const [key, imgPath] of Object.entries(IMG)) {
  if (!fs.existsSync(imgPath)) {
    console.warn(`WARNING: Image not found: ${imgPath}`);
  }
}

const pptx = new pptxgen();

// ── Presentation metadata ──
pptx.author = "Mehdi Mohammad";
pptx.company = "KE4H Project";
pptx.subject = "The Authenticity of Tourism in Palmyra";
pptx.title = "Palmyra Heritage - KE4H Project";
pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// ══════════════════════════════════════════════
// HELPER: Add section header (e.g., "02 / PROJECT CONTEXT")
// ══════════════════════════════════════════════
function addSectionHeader(slide, num, label) {
  // Teal accent bar (vertical)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.6,
    y: 0.4,
    w: 0.06,
    h: 0.45,
    fill: { color: C.teal },
  });
  // Section text
  slide.addText(`${num} / ${label}`, {
    x: 0.8,
    y: 0.4,
    w: 5,
    h: 0.45,
    fontSize: 11,
    fontFace: FONT_SANS,
    color: C.teal,
    bold: true,
    charSpacing: 3,
  });
}

// ══════════════════════════════════════════════
// HELPER: Add key insight box
// ══════════════════════════════════════════════
function addKeyInsight(slide, text, y, x = 0.6, w = 5.5) {
  // Divider line
  slide.addShape(pptx.shapes.LINE, {
    x: x,
    y: y,
    w: w,
    h: 0,
    line: { color: C.teal, width: 1.5 },
  });
  // "KEY INSIGHT" label
  slide.addText("KEY INSIGHT", {
    x: x,
    y: y + 0.12,
    w: 2,
    h: 0.3,
    fontSize: 9,
    fontFace: FONT_SANS,
    color: C.teal,
    bold: true,
    charSpacing: 2,
  });
  // Insight text
  slide.addText(text, {
    x: x,
    y: y + 0.4,
    w: w,
    h: 0.7,
    fontSize: 13,
    fontFace: FONT_SERIF,
    color: C.darkNavy2,
    bold: true,
    italic: true,
    lineSpacingMultiple: 1.3,
  });
}

// ══════════════════════════════════════════════
// HELPER: Content slide background
// ══════════════════════════════════════════════
function contentBg(slide) {
  slide.background = { fill: C.warmBeige };
}

// ══════════════════════════════════════════════
// SLIDE 1: COVER
// ══════════════════════════════════════════════
function slide01_cover() {
  const slide = pptx.addSlide();
  slide.background = { fill: C.darkNavy };

  // Left teal accent bar
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.6,
    y: 1.2,
    w: 0.06,
    h: 3.2,
    fill: { color: C.teal },
  });

  // Main title
  slide.addText("The Authenticity of\nTourism in Palmyra", {
    x: 0.85,
    y: 1.0,
    w: 7.5,
    h: 2.4,
    fontSize: 42,
    fontFace: FONT_SERIF,
    color: C.white,
    bold: true,
    lineSpacingMultiple: 1.15,
  });

  // Subtitle
  slide.addText(
    "Integrating Knowledge Graphs and Large Language\nModels for Cultural Heritage (KE4H Project)",
    {
      x: 0.85,
      y: 3.3,
      w: 7.5,
      h: 0.8,
      fontSize: 15,
      fontFace: FONT_SERIF,
      color: C.lightGray,
      lineSpacingMultiple: 1.4,
    }
  );

  // Course info
  const infoItems = [
    { label: "Course:", value: "Knowledge Engineering for Heritage", color: C.salmon },
    { label: "Student:", value: "Mehdi Mohammad", color: C.teal },
    { label: "Academic Year:", value: "2025-2026", color: C.accentLight },
  ];

  infoItems.forEach((item, i) => {
    slide.addText(
      [
        { text: `${item.label} `, options: { color: item.color, bold: true, fontSize: 12 } },
        { text: item.value, options: { color: C.lightGray, fontSize: 12 } },
      ],
      {
        x: 0.85,
        y: 4.4 + i * 0.38,
        w: 7,
        h: 0.35,
        fontFace: FONT_SANS,
      }
    );
  });

  // Website URL
  slide.addText(
    [
      {
        text: "https://mehdimohammad11381138-spec.github.io/palmyra-heritage-project/",
        options: {
          hyperlink: {
            url: "https://mehdimohammad11381138-spec.github.io/palmyra-heritage-project/",
          },
          color: C.teal,
          fontSize: 10,
          fontFace: FONT_SANS,
          underline: true,
        },
      },
    ],
    {
      x: 0.85,
      y: 5.7,
      w: 8,
      h: 0.3,
    }
  );

  // Colonnade image on right (subtle, with overlay feel)
  if (fs.existsSync(IMG.colonnade)) {
    slide.addImage({
      path: IMG.colonnade,
      x: 8.2,
      y: 0,
      w: 5.13,
      h: 7.5,
      sizing: { type: "cover", w: 5.13, h: 7.5 },
    });
    // Dark overlay on the image
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 8.2,
      y: 0,
      w: 5.13,
      h: 7.5,
      fill: { color: C.darkNavy, transparency: 40 },
    });
  }

  // Slide number
  slide.addText("KE4H Project · June 2026", {
    x: 0.85,
    y: 6.8,
    w: 5,
    h: 0.3,
    fontSize: 9,
    fontFace: FONT_SANS,
    color: C.bodyMuted,
  });
}

// ══════════════════════════════════════════════
// SLIDE 2: PROJECT CONTEXT
// ══════════════════════════════════════════════
function slide02_context() {
  const slide = pptx.addSlide();
  contentBg(slide);
  addSectionHeader(slide, "02", "PROJECT CONTEXT");

  // Heading
  slide.addText("The Goal of My Project", {
    x: 0.6,
    y: 1.0,
    w: 5,
    h: 0.8,
    fontSize: 32,
    fontFace: FONT_SERIF,
    color: C.headingNavy,
    bold: true,
  });

  // Content - right column
  slide.addText(
    "I studied Palmyra, a UNESCO World Heritage site, to identify digital information gaps in how historical data is stored. My goal was to enrich existing databases with accurate, AI-extracted data regarding cultural heritage.",
    {
      x: 6.2,
      y: 0.6,
      w: 6.3,
      h: 1.2,
      fontSize: 13,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.5,
    }
  );

  slide.addText(
    'I focused on "authenticity" and restoration efforts post-conflict, covering the critical period of 2015-2017. The project culminated in a comprehensive digital documentation website to record all my findings and methodologies.',
    {
      x: 6.2,
      y: 2.0,
      w: 6.3,
      h: 1.2,
      fontSize: 13,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.5,
    }
  );

  // My approach bullets
  const bullets = [
    "I identify missing information in existing knowledge graphs.",
    "I test SPARQL queries to inspect available data.",
    "I use LLMs to support enrichment and comparison.",
    "I document the full workflow on a website report.",
  ];

  slide.addText(
    bullets.map((b) => ({
      text: `•  ${b}`,
      options: { fontSize: 12, color: C.bodyText, bullet: false, paraSpaceAfter: 6 },
    })),
    {
      x: 0.6,
      y: 2.2,
      w: 5.2,
      h: 2.5,
      fontFace: FONT_SANS,
      lineSpacingMultiple: 1.5,
    }
  );

  addKeyInsight(
    slide,
    "This project bridges the gap between historical knowledge and digital data using AI.",
    5.0,
    6.2,
    6.3
  );
}

// ══════════════════════════════════════════════
// SLIDE 3: WHY PALMYRA
// ══════════════════════════════════════════════
function slide03_whyPalmyra() {
  const slide = pptx.addSlide();
  contentBg(slide);
  addSectionHeader(slide, "03", "SELECTED TOPIC");

  slide.addText("Why Palmyra?", {
    x: 0.6,
    y: 1.0,
    w: 6,
    h: 0.8,
    fontSize: 32,
    fontFace: FONT_SERIF,
    color: C.headingNavy,
    bold: true,
  });

  slide.addText(
    'Palmyra, "The Pearl of the Desert," is an ancient Syrian city and a major Silk Road trading hub. As a UNESCO site since 1980, it suffered severe damage during recent conflicts (2015-2017) and is currently undergoing international restoration.',
    {
      x: 0.6,
      y: 1.9,
      w: 6.8,
      h: 1.2,
      fontSize: 13,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.5,
    }
  );

  // Reasons
  slide.addText("My Reasons for Selection:", {
    x: 0.6,
    y: 3.2,
    w: 5,
    h: 0.35,
    fontSize: 13,
    fontFace: FONT_SANS,
    color: C.headingNavy,
    bold: true,
  });

  const reasons = [
    {
      label: "Critical Heritage:",
      text: "Palmyra represents a crucial case study in post-conflict heritage preservation.",
    },
    {
      label: "Data Gaps:",
      text: "I identified significant gaps in existing knowledge graphs regarding its restoration.",
    },
    {
      label: "Personal Interest:",
      text: "My drive for cultural heritage and AI-driven solutions led my selection.",
    },
  ];

  reasons.forEach((r, i) => {
    slide.addText(
      [
        { text: `•  ${r.label} `, options: { bold: true, color: C.headingNavy, fontSize: 12 } },
        { text: r.text, options: { color: C.bodyText, fontSize: 12 } },
      ],
      {
        x: 0.8,
        y: 3.65 + i * 0.55,
        w: 6.5,
        h: 0.5,
        fontFace: FONT_SANS,
        lineSpacingMultiple: 1.3,
      }
    );
  });

  addKeyInsight(
    slide,
    "Palmyra is a historically significant site facing critical modern preservation challenges.",
    5.5,
    0.6,
    6.5
  );

  // Sticker strip image on right
  if (fs.existsSync(IMG.stickers)) {
    slide.addImage({
      path: IMG.stickers,
      x: 8.2,
      y: 1.5,
      w: 4.6,
      h: 1.6,
    });
  }

  // Real photo below sticker
  if (fs.existsSync(IMG.colonnade)) {
    slide.addImage({
      path: IMG.colonnade,
      x: 8.2,
      y: 3.5,
      w: 4.6,
      h: 3.2,
      rounding: true,
      sizing: { type: "cover", w: 4.6, h: 3.2 },
    });
  }
}

// ══════════════════════════════════════════════
// SLIDE 4: METHODOLOGY
// ══════════════════════════════════════════════
function slide04_methodology() {
  const slide = pptx.addSlide();
  contentBg(slide);
  addSectionHeader(slide, "04", "METHODOLOGY");

  slide.addText("My Research Workflow", {
    x: 0.6,
    y: 1.0,
    w: 8,
    h: 0.8,
    fontSize: 32,
    fontFace: FONT_SERIF,
    color: C.headingNavy,
    bold: true,
  });

  const steps = [
    {
      num: "01",
      title: "Explore Knowledge Graphs",
      desc: "Inspect Wikidata and DBpedia for existing Palmyra data",
    },
    {
      num: "02",
      title: "Identify Gaps",
      desc: "Find missing or incomplete information about authenticity and restoration",
    },
    {
      num: "03",
      title: "LLM Extraction",
      desc: "Use ChatGPT, Claude, and Gemini to extract and synthesize relevant details",
    },
    {
      num: "04",
      title: "RDF Enrichment",
      desc: "Create new RDF triples to enrich the knowledge graph data",
    },
    {
      num: "05",
      title: "Document & Evaluate",
      desc: "Record the full process and results on the project website",
    },
  ];

  steps.forEach((step, i) => {
    const xBase = 0.6;
    const yBase = 2.1 + i * 1.0;

    // Step number circle
    slide.addShape(pptx.shapes.OVAL, {
      x: xBase,
      y: yBase,
      w: 0.55,
      h: 0.55,
      fill: { color: i < 3 ? C.teal : C.accent },
    });

    slide.addText(step.num, {
      x: xBase,
      y: yBase,
      w: 0.55,
      h: 0.55,
      fontSize: 13,
      fontFace: FONT_SANS,
      color: C.white,
      bold: true,
      align: "center",
      valign: "middle",
    });

    // Connector line
    if (i < steps.length - 1) {
      slide.addShape(pptx.shapes.LINE, {
        x: xBase + 0.275,
        y: yBase + 0.55,
        w: 0,
        h: 0.45,
        line: { color: C.lightGray, width: 2 },
      });
    }

    // Title
    slide.addText(step.title, {
      x: 1.4,
      y: yBase,
      w: 3.5,
      h: 0.3,
      fontSize: 14,
      fontFace: FONT_SANS,
      color: C.headingNavy,
      bold: true,
    });

    // Description
    slide.addText(step.desc, {
      x: 1.4,
      y: yBase + 0.28,
      w: 4.5,
      h: 0.35,
      fontSize: 11,
      fontFace: FONT_SANS,
      color: C.bodyMuted,
    });
  });

  // Theatre image on right
  if (fs.existsSync(IMG.theatre)) {
    slide.addImage({
      path: IMG.theatre,
      x: 7.5,
      y: 1.5,
      w: 5.3,
      h: 5.3,
      sizing: { type: "cover", w: 5.3, h: 5.3 },
    });
    // Soft overlay
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 7.5,
      y: 1.5,
      w: 5.3,
      h: 5.3,
      fill: { color: C.warmBeige, transparency: 80 },
    });
  }

  addKeyInsight(
    slide,
    "Each step builds upon the previous one, creating a reproducible research pipeline.",
    6.6,
    7.5,
    5.3
  );
}

// ══════════════════════════════════════════════
// SLIDE 5: KNOWLEDGE GRAPHS & SPARQL
// ══════════════════════════════════════════════
function slide05_kgSparql() {
  const slide = pptx.addSlide();
  contentBg(slide);
  addSectionHeader(slide, "05", "KNOWLEDGE GRAPHS & SPARQL");

  slide.addText("Exploring Structured Data", {
    x: 0.6,
    y: 1.0,
    w: 8,
    h: 0.8,
    fontSize: 32,
    fontFace: FONT_SERIF,
    color: C.headingNavy,
    bold: true,
  });

  // Left column - Knowledge Graphs
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6,
    y: 2.0,
    w: 5.8,
    h: 2.3,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 8, offset: 2, color: "000000", opacity: 0.08 },
    rectRadius: 0.15,
  });

  slide.addText("Knowledge Graphs", {
    x: 0.9,
    y: 2.15,
    w: 5,
    h: 0.4,
    fontSize: 16,
    fontFace: FONT_SANS,
    color: C.teal,
    bold: true,
  });

  slide.addText(
    "Knowledge graphs help me represent heritage information as connected facts instead of isolated notes. They make it possible to link places, monuments, restoration methods, dates, and external identifiers in a machine-readable way.",
    {
      x: 0.9,
      y: 2.6,
      w: 5.2,
      h: 1.4,
      fontSize: 12,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.5,
    }
  );

  // Right column - SPARQL
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 6.8,
    y: 2.0,
    w: 5.8,
    h: 2.3,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 8, offset: 2, color: "000000", opacity: 0.08 },
    rectRadius: 0.15,
  });

  slide.addText("SPARQL Queries", {
    x: 7.1,
    y: 2.15,
    w: 5,
    h: 0.4,
    fontSize: 16,
    fontFace: FONT_SANS,
    color: C.accent,
    bold: true,
  });

  slide.addText(
    "I wrote several SPARQL queries to inspect monuments, restoration data, and UNESCO-related information. This helped me understand what was already present in the graph and where the gaps were.",
    {
      x: 7.1,
      y: 2.6,
      w: 5.2,
      h: 1.4,
      fontSize: 12,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.5,
    }
  );

  // Query examples box
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6,
    y: 4.6,
    w: 12,
    h: 2.3,
    fill: { color: C.darkNavy },
    rectRadius: 0.15,
  });

  slide.addText("Query Examples", {
    x: 0.9,
    y: 4.75,
    w: 5,
    h: 0.4,
    fontSize: 14,
    fontFace: FONT_SANS,
    color: C.teal,
    bold: true,
  });

  const queries = [
    "Monument discovery — retrieving all heritage sites in Palmyra from Wikidata",
    "Damage & restoration — checking for structured conflict and restoration data",
    "UNESCO status — verifying heritage references and recognition metadata",
  ];

  queries.forEach((q, i) => {
    slide.addText(`→  ${q}`, {
      x: 0.9,
      y: 5.2 + i * 0.45,
      w: 11,
      h: 0.4,
      fontSize: 12,
      fontFace: FONT_SANS,
      color: C.offWhite,
      lineSpacingMultiple: 1.3,
    });
  });
}

// ══════════════════════════════════════════════
// SLIDE 6: GAP IDENTIFIED
// ══════════════════════════════════════════════
function slide06_gap() {
  const slide = pptx.addSlide();
  contentBg(slide);
  addSectionHeader(slide, "06", "GAP IDENTIFIED");

  slide.addText("What's Missing?", {
    x: 0.6,
    y: 1.0,
    w: 6,
    h: 0.8,
    fontSize: 32,
    fontFace: FONT_SERIF,
    color: C.headingNavy,
    bold: true,
  });

  slide.addText(
    "The main gap I found is that current knowledge graphs do not fully capture the complexity of authenticity, restoration, and post-conflict heritage documentation for Palmyra.",
    {
      x: 0.6,
      y: 1.9,
      w: 6.5,
      h: 1.0,
      fontSize: 13,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.5,
    }
  );

  // Missing items as cards
  const gaps = [
    { icon: "📋", text: "Detailed restoration history" },
    { icon: "🔍", text: "Explicit authenticity-related metadata" },
    { icon: "📅", text: "Temporal information about change over time" },
    { icon: "👥", text: "Structured stakeholder and community perspectives" },
  ];

  gaps.forEach((gap, i) => {
    const x = 0.6 + (i % 2) * 3.5;
    const y = 3.2 + Math.floor(i / 2) * 1.2;

    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: x,
      y: y,
      w: 3.2,
      h: 0.9,
      fill: { color: C.white },
      line: { color: C.lightGray, width: 1 },
      rectRadius: 0.1,
    });

    slide.addText(`${gap.icon}  ${gap.text}`, {
      x: x + 0.15,
      y: y + 0.1,
      w: 2.9,
      h: 0.7,
      fontSize: 12,
      fontFace: FONT_SANS,
      color: C.headingNavy,
      bold: false,
      valign: "middle",
    });
  });

  // Satellite comparison image on right
  if (fs.existsSync(IMG.satellite)) {
    slide.addImage({
      path: IMG.satellite,
      x: 7.8,
      y: 1.2,
      w: 4.8,
      h: 5.5,
      sizing: { type: "cover", w: 4.8, h: 5.5 },
    });
  }

  // Caption under satellite
  slide.addText("Satellite comparison: Sept 2010 vs Aug 2015", {
    x: 7.8,
    y: 6.8,
    w: 4.8,
    h: 0.3,
    fontSize: 9,
    fontFace: FONT_SANS,
    color: C.bodyMuted,
    italic: true,
    align: "center",
  });

  addKeyInsight(
    slide,
    "These gaps show the need for structured, verifiable heritage enrichment.",
    5.8,
    0.6,
    6.5
  );
}

// ══════════════════════════════════════════════
// SLIDE 7: LLM ENRICHMENT & PROMPTING
// ══════════════════════════════════════════════
function slide07_llm() {
  const slide = pptx.addSlide();
  contentBg(slide);
  addSectionHeader(slide, "07", "LLM ENRICHMENT");

  slide.addText("AI-Powered Knowledge Extraction", {
    x: 0.6,
    y: 1.0,
    w: 10,
    h: 0.8,
    fontSize: 32,
    fontFace: FONT_SERIF,
    color: C.headingNavy,
    bold: true,
  });

  slide.addText(
    "I used large language models to help extract missing information from texts and reports. I did not rely on the models alone — I checked the outputs against reliable sources before turning them into structured data.",
    {
      x: 0.6,
      y: 1.9,
      w: 12,
      h: 0.8,
      fontSize: 13,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.5,
    }
  );

  // Models used cards
  const models = [
    { name: "ChatGPT", desc: "Broad coverage, good for initial extraction", color: C.teal },
    { name: "Claude", desc: "Strong reasoning, careful about uncertainty", color: C.accent },
    { name: "Gemini", desc: "Cross-referencing, multimodal capabilities", color: C.salmon },
  ];

  models.forEach((m, i) => {
    const x = 0.6 + i * 4.1;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: x,
      y: 3.0,
      w: 3.8,
      h: 1.1,
      fill: { color: C.white },
      line: { color: m.color, width: 2 },
      rectRadius: 0.1,
    });

    slide.addText(m.name, {
      x: x + 0.2,
      y: 3.1,
      w: 3.4,
      h: 0.35,
      fontSize: 15,
      fontFace: FONT_SANS,
      color: m.color,
      bold: true,
    });

    slide.addText(m.desc, {
      x: x + 0.2,
      y: 3.5,
      w: 3.4,
      h: 0.4,
      fontSize: 11,
      fontFace: FONT_SANS,
      color: C.bodyMuted,
    });
  });

  // Prompting techniques section
  slide.addText("Prompting Techniques", {
    x: 0.6,
    y: 4.5,
    w: 5,
    h: 0.4,
    fontSize: 16,
    fontFace: FONT_SANS,
    color: C.headingNavy,
    bold: true,
  });

  const techniques = [
    {
      name: "Zero-shot",
      desc: "Direct answers quickly without examples",
    },
    {
      name: "Few-shot",
      desc: "Improved consistency by giving examples",
    },
    {
      name: "Chain-of-thought",
      desc: "More transparent reasoning when needed",
    },
  ];

  techniques.forEach((t, i) => {
    const x = 0.6 + i * 4.1;
    slide.addText(
      [
        { text: `${t.name}\n`, options: { fontSize: 13, bold: true, color: C.teal } },
        { text: t.desc, options: { fontSize: 11, color: C.bodyText } },
      ],
      {
        x: x,
        y: 5.0,
        w: 3.8,
        h: 0.8,
        fontFace: FONT_SANS,
        lineSpacingMultiple: 1.3,
      }
    );
  });

  addKeyInsight(
    slide,
    "LLMs were used as assistants — not as the final authority. Every claim was verified manually.",
    6.2,
    0.6,
    12
  );
}

// ══════════════════════════════════════════════
// SLIDE 8: RDF ENRICHMENT & VOCABULARY
// ══════════════════════════════════════════════
function slide08_rdf() {
  const slide = pptx.addSlide();
  contentBg(slide);
  addSectionHeader(slide, "08", "RDF ENRICHMENT");

  slide.addText("Structured Data Creation", {
    x: 0.6,
    y: 1.0,
    w: 8,
    h: 0.8,
    fontSize: 32,
    fontFace: FONT_SERIF,
    color: C.headingNavy,
    bold: true,
  });

  slide.addText(
    "I created new RDF triples to enrich the dataset in a standardized way, making the information reusable, structured, and easier to query.",
    {
      x: 0.6,
      y: 1.9,
      w: 7,
      h: 0.7,
      fontSize: 13,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.5,
    }
  );

  // RDF triples box
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6,
    y: 2.8,
    w: 7,
    h: 2.2,
    fill: { color: C.darkNavy },
    rectRadius: 0.15,
  });

  slide.addText("Example RDF Triples", {
    x: 0.9,
    y: 2.9,
    w: 5,
    h: 0.35,
    fontSize: 12,
    fontFace: FONT_SANS,
    color: C.teal,
    bold: true,
  });

  const triples = [
    "Temple_of_Bel  →  hasRestorationMethod  →  Anastylosis",
    "Temple_of_Bel  →  authenticityStatus  →  Composite_Authenticity",
    "Temple_of_Bel  →  lastUpdated  →  2026-05-27",
  ];

  triples.forEach((t, i) => {
    slide.addText(t, {
      x: 0.9,
      y: 3.35 + i * 0.5,
      w: 6.5,
      h: 0.4,
      fontSize: 12,
      fontFace: "Consolas",
      color: C.accentLight,
    });
  });

  // Vocabulary Extension - right side
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 8.0,
    y: 2.8,
    w: 4.6,
    h: 2.2,
    fill: { color: C.white },
    line: { color: C.teal, width: 2 },
    rectRadius: 0.15,
  });

  slide.addText("Vocabulary Extension", {
    x: 8.3,
    y: 2.9,
    w: 4,
    h: 0.35,
    fontSize: 14,
    fontFace: FONT_SANS,
    color: C.teal,
    bold: true,
  });

  slide.addText(
    "The existing vocabulary was not always enough. I introduced a small extension to represent heritage-recognition information more clearly.",
    {
      x: 8.3,
      y: 3.3,
      w: 4,
      h: 0.8,
      fontSize: 11,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.4,
    }
  );

  slide.addText(
    [
      { text: "Added: ", options: { bold: true, color: C.headingNavy } },
      { text: "ex:hasHeritageRecognition", options: { color: C.accent, bold: true } },
    ],
    {
      x: 8.3,
      y: 4.2,
      w: 4,
      h: 0.35,
      fontSize: 12,
      fontFace: FONT_SANS,
    }
  );

  slide.addText("Links a site to its heritage recognition with provenance visibility.", {
    x: 8.3,
    y: 4.5,
    w: 4,
    h: 0.4,
    fontSize: 10,
    fontFace: FONT_SANS,
    color: C.bodyMuted,
    italic: true,
  });

  // Lion sculpture image
  if (fs.existsSync(IMG.lion)) {
    slide.addImage({
      path: IMG.lion,
      x: 8.5,
      y: 5.3,
      w: 2.0,
      h: 2.0,
      sizing: { type: "cover", w: 2.0, h: 2.0 },
    });
  }

  addKeyInsight(
    slide,
    "Structured RDF triples make the data queryable, reusable, and machine-readable.",
    5.3,
    0.6,
    6.5
  );
}

// ══════════════════════════════════════════════
// SLIDE 9: SOURCE & LLM COMPARISON
// ══════════════════════════════════════════════
function slide09_comparison() {
  const slide = pptx.addSlide();
  contentBg(slide);
  addSectionHeader(slide, "09", "COMPARISON");

  slide.addText("Source & Model Analysis", {
    x: 0.6,
    y: 1.0,
    w: 8,
    h: 0.8,
    fontSize: 32,
    fontFace: FONT_SERIF,
    color: C.headingNavy,
    bold: true,
  });

  // Source comparison table
  slide.addText("Source Comparison", {
    x: 0.6,
    y: 2.0,
    w: 5,
    h: 0.35,
    fontSize: 15,
    fontFace: FONT_SANS,
    color: C.teal,
    bold: true,
  });

  const sources = [
    { source: "Wikipedia", framing: "Historical & encyclopedic" },
    { source: "Official Site", framing: "Institutional & visitor-oriented" },
    { source: "UNESCO", framing: "Heritage & preservation focus" },
    { source: "Knowledge Graphs", framing: "Structured & factual data" },
  ];

  // Table rows
  const tableRows = [
    [
      { text: "Source", options: { bold: true, color: C.white, fill: { color: C.teal }, fontSize: 11, fontFace: FONT_SANS } },
      { text: "Framing", options: { bold: true, color: C.white, fill: { color: C.teal }, fontSize: 11, fontFace: FONT_SANS } },
    ],
    ...sources.map((s) => [
      { text: s.source, options: { bold: true, color: C.headingNavy, fontSize: 11, fontFace: FONT_SANS } },
      { text: s.framing, options: { color: C.bodyText, fontSize: 11, fontFace: FONT_SANS } },
    ]),
  ];

  slide.addTable(tableRows, {
    x: 0.6,
    y: 2.5,
    w: 5.5,
    border: { type: "solid", pt: 0.5, color: C.lightGray },
    rowH: [0.4, 0.4, 0.4, 0.4, 0.4],
    colW: [2.0, 3.5],
  });

  // LLM Comparison - right side
  slide.addText("LLM Comparison", {
    x: 7.0,
    y: 2.0,
    w: 5,
    h: 0.35,
    fontSize: 15,
    fontFace: FONT_SANS,
    color: C.accent,
    bold: true,
  });

  slide.addText(
    "I compared how different models handled the same task and found that they vary in clarity, completeness, and factual reliability.",
    {
      x: 7.0,
      y: 2.5,
      w: 5.5,
      h: 0.8,
      fontSize: 12,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.4,
    }
  );

  const llmFindings = [
    "Models differ in factual grounding and confidence levels",
    "Some outputs sounded plausible but lacked reliable sourcing",
    "A hybrid method combining RDF + source excerpts gives the best results",
  ];

  llmFindings.forEach((f, i) => {
    slide.addText(`→  ${f}`, {
      x: 7.0,
      y: 3.5 + i * 0.5,
      w: 5.5,
      h: 0.4,
      fontSize: 12,
      fontFace: FONT_SANS,
      color: C.bodyText,
    });
  });

  addKeyInsight(
    slide,
    "Combining multiple sources and models produces a more complete and reliable picture.",
    5.5,
    0.6,
    12
  );
}

// ══════════════════════════════════════════════
// SLIDE 10: VISUAL DOCUMENTATION (Photo Gallery)
// ══════════════════════════════════════════════
function slide10_gallery() {
  const slide = pptx.addSlide();
  slide.background = { fill: C.darkNavy };

  slide.addText("Palmyra — Visual Documentation", {
    x: 0.6,
    y: 0.3,
    w: 12,
    h: 0.6,
    fontSize: 26,
    fontFace: FONT_SERIF,
    color: C.white,
    bold: true,
  });

  // Image grid — 3 photos top, 2 bottom
  const topImages = [
    { path: IMG.colonnade, caption: "The Great Colonnade at Sunset" },
    { path: IMG.theatre, caption: "The Roman Theatre" },
    { path: IMG.beforeAfter, caption: "Temple of Bel — Then & Now" },
  ];

  topImages.forEach((img, i) => {
    if (fs.existsSync(img.path)) {
      const x = 0.4 + i * 4.2;
      slide.addImage({
        path: img.path,
        x: x,
        y: 1.1,
        w: 3.9,
        h: 2.8,
        sizing: { type: "cover", w: 3.9, h: 2.8 },
      });
      slide.addText(img.caption, {
        x: x,
        y: 3.95,
        w: 3.9,
        h: 0.3,
        fontSize: 9,
        fontFace: FONT_SANS,
        color: C.accentLight,
        align: "center",
        italic: true,
      });
    }
  });

  const bottomImages = [
    { path: IMG.satellite, caption: "Satellite Evidence: Sept 2010 vs Aug 2015" },
    { path: IMG.lion, caption: "The Lion of Al-lāt — Restored" },
  ];

  bottomImages.forEach((img, i) => {
    if (fs.existsSync(img.path)) {
      const x = 1.6 + i * 5.6;
      slide.addImage({
        path: img.path,
        x: x,
        y: 4.5,
        w: 4.8,
        h: 2.5,
        sizing: { type: "cover", w: 4.8, h: 2.5 },
      });
      slide.addText(img.caption, {
        x: x,
        y: 7.05,
        w: 4.8,
        h: 0.3,
        fontSize: 9,
        fontFace: FONT_SANS,
        color: C.accentLight,
        align: "center",
        italic: true,
      });
    }
  });
}

// ══════════════════════════════════════════════
// SLIDE 11: CHALLENGES
// ══════════════════════════════════════════════
function slide11_challenges() {
  const slide = pptx.addSlide();
  contentBg(slide);
  addSectionHeader(slide, "11", "CHALLENGES");

  slide.addText("Obstacles & Solutions", {
    x: 0.6,
    y: 1.0,
    w: 8,
    h: 0.8,
    fontSize: 32,
    fontFace: FONT_SERIF,
    color: C.headingNavy,
    bold: true,
  });

  slide.addText(
    "During the project, I faced several challenges that shaped my approach and improved the quality of the final output.",
    {
      x: 0.6,
      y: 1.85,
      w: 12,
      h: 0.5,
      fontSize: 13,
      fontFace: FONT_SANS,
      color: C.bodyText,
      lineSpacingMultiple: 1.4,
    }
  );

  const challenges = [
    {
      num: "1",
      title: "Finding Reliable Data",
      problem: "Different sources sometimes gave conflicting dates or descriptions of Palmyra's history and restoration.",
      solution: "I cross-checked information across Wikidata, academic sources, UNESCO material, and the official site to ensure accuracy.",
      color: C.teal,
    },
    {
      num: "2",
      title: "Avoiding LLM Hallucinations",
      problem: "Some LLM outputs sounded plausible but were not fully reliable. Models sometimes generated facts that could not be verified.",
      solution: "I verified every important claim manually before converting it into RDF triples. No unverified data entered the knowledge graph.",
      color: C.accent,
    },
    {
      num: "3",
      title: "Balancing Depth and Clarity",
      problem: "I had to keep the presentation concise while still explaining the full project workflow and methodology clearly.",
      solution: "I focused on the methodology, the gap, the enrichment process, and evaluation — rather than listing every technical detail.",
      color: C.salmon,
    },
  ];

  challenges.forEach((c, i) => {
    const y = 2.6 + i * 1.55;

    // Challenge card
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6,
      y: y,
      w: 12,
      h: 1.35,
      fill: { color: C.white },
      line: { color: c.color, width: 0, type: "none" },
      shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.06 },
      rectRadius: 0.1,
    });

    // Left accent bar
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0.6,
      y: y,
      w: 0.08,
      h: 1.35,
      fill: { color: c.color },
    });

    // Number badge
    slide.addShape(pptx.shapes.OVAL, {
      x: 0.85,
      y: y + 0.15,
      w: 0.45,
      h: 0.45,
      fill: { color: c.color },
    });

    slide.addText(c.num, {
      x: 0.85,
      y: y + 0.15,
      w: 0.45,
      h: 0.45,
      fontSize: 14,
      fontFace: FONT_SANS,
      color: C.white,
      bold: true,
      align: "center",
      valign: "middle",
    });

    // Title
    slide.addText(c.title, {
      x: 1.5,
      y: y + 0.1,
      w: 4,
      h: 0.4,
      fontSize: 14,
      fontFace: FONT_SANS,
      color: C.headingNavy,
      bold: true,
    });

    // Problem
    slide.addText(
      [
        { text: "Challenge: ", options: { bold: true, color: C.danger, fontSize: 10.5 } },
        { text: c.problem, options: { color: C.bodyText, fontSize: 10.5 } },
      ],
      {
        x: 1.5,
        y: y + 0.5,
        w: 5,
        h: 0.7,
        fontFace: FONT_SANS,
        lineSpacingMultiple: 1.3,
      }
    );

    // Solution
    slide.addText(
      [
        { text: "Solution: ", options: { bold: true, color: C.teal, fontSize: 10.5 } },
        { text: c.solution, options: { color: C.bodyText, fontSize: 10.5 } },
      ],
      {
        x: 6.8,
        y: y + 0.5,
        w: 5.5,
        h: 0.7,
        fontFace: FONT_SANS,
        lineSpacingMultiple: 1.3,
      }
    );
  });

  addKeyInsight(
    slide,
    "This project was not only about generating data — it was about evaluating quality, reliability, and presentation clarity.",
    6.9,
    0.6,
    12
  );
}

// ══════════════════════════════════════════════
// SLIDE 12: CONCLUSION
// ══════════════════════════════════════════════
function slide12_conclusion() {
  const slide = pptx.addSlide();
  slide.background = { fill: C.darkNavy };

  // Left accent bar
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.6,
    y: 0.5,
    w: 0.06,
    h: 1.0,
    fill: { color: C.teal },
  });

  slide.addText("Conclusion", {
    x: 0.85,
    y: 0.5,
    w: 8,
    h: 1.0,
    fontSize: 38,
    fontFace: FONT_SERIF,
    color: C.white,
    bold: true,
  });

  slide.addText(
    "My project shows that knowledge graphs and LLMs can work together to improve cultural heritage documentation, as long as the information is checked carefully and kept grounded in reliable sources.",
    {
      x: 0.85,
      y: 1.6,
      w: 7,
      h: 1.0,
      fontSize: 14,
      fontFace: FONT_SANS,
      color: C.offWhite,
      lineSpacingMultiple: 1.6,
    }
  );

  // Key takeaways
  const takeaways = [
    "I identified important gaps in Palmyra-related heritage data.",
    "I enriched the dataset with verified RDF triples.",
    "I used LLMs as support, not as the final authority.",
    "I documented the full process in a reproducible website report.",
  ];

  takeaways.forEach((t, i) => {
    slide.addShape(pptx.shapes.OVAL, {
      x: 0.85,
      y: 3.0 + i * 0.65,
      w: 0.2,
      h: 0.2,
      fill: { color: C.teal },
    });

    slide.addText(t, {
      x: 1.2,
      y: 2.9 + i * 0.65,
      w: 6.5,
      h: 0.4,
      fontSize: 13,
      fontFace: FONT_SANS,
      color: C.offWhite,
      lineSpacingMultiple: 1.3,
    });
  });

  // Divider
  slide.addShape(pptx.shapes.LINE, {
    x: 0.85,
    y: 5.8,
    w: 6.5,
    h: 0,
    line: { color: C.teal, width: 1.5 },
  });

  // Thank you
  slide.addText("Thank You", {
    x: 0.85,
    y: 6.0,
    w: 5,
    h: 0.6,
    fontSize: 22,
    fontFace: FONT_SERIF,
    color: C.accentLight,
    bold: true,
    italic: true,
  });

  // Student info
  slide.addText("Mehdi Mohammad · KE4H Project · June 2026", {
    x: 0.85,
    y: 6.6,
    w: 6,
    h: 0.3,
    fontSize: 10,
    fontFace: FONT_SANS,
    color: C.bodyMuted,
  });

  // Website
  slide.addText(
    [
      {
        text: "mehdimohammad11381138-spec.github.io/palmyra-heritage-project/",
        options: {
          hyperlink: {
            url: "https://mehdimohammad11381138-spec.github.io/palmyra-heritage-project/",
          },
          color: C.teal,
          fontSize: 10,
          fontFace: FONT_SANS,
          underline: true,
        },
      },
    ],
    {
      x: 0.85,
      y: 6.9,
      w: 8,
      h: 0.3,
    }
  );

  // Before/After photo on right
  if (fs.existsSync(IMG.beforeAfter)) {
    slide.addImage({
      path: IMG.beforeAfter,
      x: 8.5,
      y: 0.5,
      w: 4.3,
      h: 6.5,
      sizing: { type: "cover", w: 4.3, h: 6.5 },
    });
    // Dark overlay
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 8.5,
      y: 0.5,
      w: 4.3,
      h: 6.5,
      fill: { color: C.darkNavy, transparency: 35 },
    });
  }
}

// ══════════════════════════════════════════════
// BUILD ALL SLIDES
// ══════════════════════════════════════════════
console.log("🏛️  Generating Palmyra Heritage presentation...");

slide01_cover();
slide02_context();
slide03_whyPalmyra();
slide04_methodology();
slide05_kgSparql();
slide06_gap();
slide07_llm();
slide08_rdf();
slide09_comparison();
slide10_gallery();
slide11_challenges();
slide12_conclusion();

// Save
const outputPath = path.resolve("Palmyra_Heritage_KE4H_Mehdi_Mohammad.pptx");
pptx
  .writeFile({ fileName: outputPath })
  .then(() => {
    console.log(`✅ Presentation saved to: ${outputPath}`);
    console.log(`📊 Total slides: 12`);
  })
  .catch((err) => {
    console.error("Error creating PPTX:", err);
  });
