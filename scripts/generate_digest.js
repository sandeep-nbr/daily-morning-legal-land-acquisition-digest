/**
 * NBR Group — Daily Morning Legal & Land Acquisition Digest
 * Generates a formatted .docx report for the Legal Department & Land Acquisition team.
 *
 * Usage: node scripts/generate_digest.js
 * Output: output/NBR_Legal_Land_Digest_<YYYY-MM-DD>.docx
 *
 * To produce a new day's digest: update the CONTENT object below with that
 * morning's research (Layer 1 Bengaluru/Karnataka, Layer 2 South India,
 * competitor tables, recommendations), then re-run this script.
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  Header, Footer, PageNumber, VerticalAlign, convertInchesToTwip,
  TabStopType, LevelFormat,
} = require("docx");

// ---------------------------------------------------------------------------
// BRAND PALETTE
// ---------------------------------------------------------------------------
const NAVY = "0B2545";
const NAVY_DARK = "081A33";
const GOLD = "C9972C";
const GOLD_LIGHT = "E8D6A8";
const GREY_BAND = "F2F2F2";
const GREY_LINE = "D9D9D9";
const WHITE = "FFFFFF";
const TEXT_DARK = "1A1A1A";
const RED_RISK = "A3121B";
const GREEN_POS = "1E7B34";
const AMBER_WATCH = "9C6F00";

// ---------------------------------------------------------------------------
// CONTENT — edit this block each morning with fresh research
// ---------------------------------------------------------------------------
const CONTENT = {
  dateLabel: "Wednesday, 22 July 2026",
  editionNo: "Vol. I — Edition 001",

  execSummary: [
    { tag: "HIGH", color: RED_RISK, text: "Karnataka HC (June 2026) set aside a KIAD Act land acquisition as a \"colourable exercise of power\" for benefiting a private entity — raises the litigation bar for any acquisition NBR Group pursues via industrial-area routes." },
    { tag: "HIGH", color: AMBER_WATCH, text: "BDA's Bengaluru Business Corridor (rebranded Peripheral Ring Road) land acquisition is live — ₹21,000 Cr earmarked, revised compensation approved for 948 acres. Corridor land values in North/East Bengaluru will re-rate." },
    { tag: "WATCH", color: AMBER_WATCH, text: "B-Khata to A-Khata conversion charge cut from 5% to 2% of guidance value for a 100-day window (15 May – 23 Aug 2026) — a narrowing opportunity for NBR's land-bank regularisation." },
    { tag: "WATCH", color: GREEN_POS, text: "Guidance values raised 6–15% across Bengaluru (Feb 2026); registration now ~7.5–7.6% of property value — factor into upcoming JDA and outright purchase cost models." },
    { tag: "INFO", color: NAVY, text: "Listed developers dominate FY26 land deals — Bengaluru led all Indian cities with 17 deals / 293 acres. Godrej Properties (17 deals, 443.5 acres) and Brigade (8 deals, ~81 acres) are the most acquisitive Grade A peers this year." },
  ],

  layer1: {
    regulatory: [
      ["K-RERA Registrations", "Karnataka RERA has now registered 8,357+ projects and 5,417+ agents; Karnataka HC has ruled the K-RERA registry cannot itself decide complaint maintainability — only the adjudicating authority can.", "Ensure every NBR project intake is registered before marketing; route buyer disputes only through the adjudicating authority, not registry correspondence."],
      ["K-RERA Delay Penalties", "K-RERA/REAT enforced up to 5% of project cost per day penalty on a developer (Mantri Developers) for non-compliance with a refund order; REAT (26 Mar 2026) clarified promoter defences and pandemic-extension compensation math.", "Audit all NBR possession-date commitments against realistic construction timelines; build RERA-compliant escrow and delay-compensation clauses into buyer agreements now."],
      ["e-Khata / e-Aasthi", "e-Khata mandatory across Bengaluru via BBMP e-Aasthi; Kaveri 2.0 now auto-triggers khata transfer on registration; 10 lakh e-Khata certificates targeted for April 2026 distribution.", "Pre-clear e-Khata status on all NBR land parcels and unsold inventory ahead of buyer registration to avoid transfer delays."],
      ["B-Khata → A-Khata Window", "Conversion charge cut from 5% to 2% of guidance value for a 100-day special window (15 May – 23 Aug 2026).", "Fast-track conversion for any NBR-held or under-acquisition B-Khata parcels before the 23 Aug 2026 window closes — this is a hard deadline."],
      ["Guidance Value Revision", "Feb 2026 statewide revision: +6–15% in Bengaluru urban limits, +4–8% in rural pockets; stamp duty 5% + 2% registration + 0.5% BBMP cess (~7.5–7.6% all-in above ₹45L).", "Re-run acquisition cost models for pending LOIs; guidance-value floor now sits closer to market price in several NBR target micro-markets (Sarjapur, Budigere, Devanahalli)."],
      ["Land Reforms Amendment 2025", "Karnataka Land Reforms & Certain Other Law (Amendment) Bill, 2025 expands Deputy Commissioners' Section 109 exemption power (0.5 ha → 4 ha, outside Bengaluru Urban/Rural); resale-after-7-years cases now need High-Power Committee (Chief Secretary) approval instead of State Government.", "For any agricultural-to-non-agricultural conversion outside core Bengaluru districts, exemption approvals may now be faster at DC level — legal team to reassess conversion timelines on outstation parcels."],
      ["Section 95 Land Conversion", "Government has committed to allow farmland-to-non-agricultural conversion within 3 days via self-declaration under Karnataka Land Revenue Act, 1964 Section 95.", "Monitor notification date closely; this could materially cut NBR's land-to-launch timeline once operationalised — legal to track rollout circulars."],
    ],
    landAcquisition: [
      ["BDA Bulk e-Auctions", "Third bulk-land e-auction round; first two rounds raised ~₹2,097 Cr. Premium Sites e-Auction 2026 (175 residential/corner/commercial sites) opened 13 Jul 2026, registration closes 27 Jul 2026.", "Legal/land team to evaluate BDA e-auction sites in NBR's active micro-markets (Sarjapur, Yelahanka, Kanakapura) before the 27 Jul deadline — competitive bidding expected from Grade A players."],
      ["Bengaluru Business Corridor (ex-PRR)", "Cabinet approved revised compensation for 948 acres; BDA to draw ₹500 Cr from HUDCO by Mar 2026 against a ₹27,000 Cr project (₹21,000 Cr for land alone).", "Corridor-adjacent land (North/East Bengaluru) is a medium-term price-appreciation zone — recommend NBR land-banking survey along BBC alignment before compensation finalisation drives up asking prices."],
      ["BDA Approval for 6,217-Acre Acquisition", "State government approved BDA's plan to acquire 6,217 acres across Bengaluru for planned development.", "Track layout notifications under this acquisition — potential TDR/compensation-linked land opportunities for NBR to evaluate on a case-by-case basis."],
      ["NICE Road Litigation Withdrawn", "NICE (Bangalore-Mysore Infrastructure Corridor) withdrew its land acquisition case in Karnataka HC after a Chief-Secretary-led mediation meeting.", "Illustrates state preference for negotiated settlement over prolonged litigation on infrastructure-linked land disputes — a useful precedent for NBR's own JDA negotiations with government-linked landholders."],
    ],
    judicial: [
      ["IMTMA v. State of Karnataka (KIAD Act)", "Karnataka HC Division Bench (Justices D K Singh & T M Nadaf) set aside acquisition of land for a private entity's expansion, calling it a \"colourable exercise of power\" and \"fraud on the statute\"; 12 landowner appeals allowed.", "Any NBR acquisition routed through KIADB/industrial-area mechanisms must have unambiguous public-purpose documentation — this ruling raises the evidentiary bar and litigation risk for such routes."],
      ["Banashankari VI Stage (BDA)", "Division Bench (Justices D K Singh & S Rachaiah) upheld BDA's acquisition for the Banashankari VI Stage township, overturning a Jan 2025 single-bench order favouring landowners; ₹50,000 costs imposed on petitioners.", "Confirms courts are increasingly upholding BDA's public-purpose township acquisitions on appeal — relevant precedent if NBR is affected by adjacent BDA layout notifications."],
    ],
    gradeA: [
      ["1", "Prestige Group", "Acquired 21-acre Whitefield parcel (₹450 Cr, ₹4,500 Cr GDV, 3rd East Bengaluru deal in 12 months); Q1 FY26 total 102 acres across Bengaluru/Hyderabad/Chennai/Mumbai (₹20,400 Cr GDV) incl. Poojanahalli-Devanahalli, Kothanuru-KR Puram, Ittangur-Sarjapura.", "₹20,400 Cr GDV (Q1 FY26, multi-city)", "Most aggressive land-bank pace among Grade A peers — directly competes with NBR in Sarjapura corridor."],
      ["2", "Brigade Group", "JDA for 8.63-acre Gunjur parcel unlocking a 39-acre township (₹7,200 Cr GDV) on Whitefield–Sarjapur corridor; 4.4-acre Whitefield buy (₹950 Cr GDV); 11-acre ITPL-facing commercial parcel (₹2,000 Cr+ GDV).", "8 deals / ~81 acres FY26", "JDA-led (capital-light) model on the same corridor NBR operates in — legal to track JDA structuring precedents being used."],
      ["3", "Sobha Limited", "₹1,150–1,160 Cr FY26 land spend; targeting ~10 million sq ft/year; new launches Sobha Altair (East Bengaluru), World City township (Hoskote) expansion.", "~10 msf/yr acquisition target", "Bengaluru-concentrated strategy (unlike diversified DLF/Godrej) — closest strategic comparable to NBR's own city focus."],
      ["4", "Puravankara Limited", "Multiple 2026 buys: 53.5-acre Anekal (₹4,800 Cr GDV), 14.57-acre Mandur-Budigere (₹2,300 Cr GDV), 11.23-acre Doddagubbi JDA (₹1,100 Cr GDV), 9.73-acre Sanna Ammanikere/airport corridor (₹800 Cr GDV).", "~89 acres across 4 deals, ~₹9,000 Cr combined GDV", "Heavy North Bengaluru airport-corridor and Budigere focus — overlaps with NBR's Sarjapur-Mullur landholding periphery."],
      ["5", "Godrej Properties", "National land-deal leader — 17 deals / 443.5 acres FY26; 20-acre East Bengaluru (NH-648/Whitefield corridor) parcel, ₹1,350 Cr revenue potential (Mar 2026); also 26-acre Sarjapur Road parcel.", "17 deals / 443.5 acres (FY26, national)", "Now directly present on Sarjapur Road — a core NBR micro-market; benchmark their pricing and launch cadence closely."],
      ["6", "Embassy Group / Embassy Developments", "6 new North Bengaluru residential launches worth ₹10,300 Cr planned for FY26 (pre-sales target ₹5,000 Cr); Embassy Sky Terraces (₹2,600 Cr, Q4 FY26) and Embassy Yelahanka additions.", "₹10,300 Cr planned launches FY26", "North Bengaluru concentration; less direct overlap with NBR's Sarjapur/South-East focus but relevant for city-wide pricing benchmarks."],
      ["7", "Salarpuria Sattva Group", "Active in Malleshwaram (Sattva Luxuria) and broader southern-corridor premium apartment segment alongside Prestige/Brigade/Embassy/Godrej/Assetz/Shriram per market trackers.", "Multiple ongoing premium launches", "Competing on redevelopment plays in mature localities — a model NBR could evaluate for infill parcels."],
      ["8", "Shriram Properties", "Named among active southern-corridor developers for plotted development and premium apartments; historical South-India plotted-project acquisitions (~₹125 Cr scale deals).", "Plotted + apartment mix", "Plotted development remains Shriram's core strength — relevant benchmark if NBR expands its own plotted-development pipeline."],
    ],
    midTier: [
      ["1", "Century Real Estate", "50+ years in Bangalore; large land-bank / plotted-development legacy portfolio, incl. Devanahalli and Kanakapura Road projects.", "Plotted development, land-bank monetisation"],
      ["2", "Assetz Property Group", "Named among active southern-corridor developers (with Prestige/Brigade/Embassy/Godrej/Sattva/Shriram) in premium apartment and redevelopment plays.", "Premium apartments, redevelopment"],
      ["3", "Vaishnavi Group", "Established Bengaluru residential/commercial/retail developer; steady project pipeline across the city.", "Residential, commercial, retail"],
      ["4", "Adarsh Developers", "Founded 1988; 33 completed + 21 ongoing projects; Adarsh Savana Phase 3 (Devanahalli plotted development, 1,200–4,000 sq ft plots) in progress.", "Apartments, villas, plotted development"],
      ["5", "Concorde Group", "Long-standing mid-tier Bengaluru developer with residential and commercial portfolio across multiple micro-markets.", "Residential, commercial"],
      ["6", "Nitesh Estates", "Boutique/premium developer with a history of high-end residential and hospitality-linked mixed-use projects in Bengaluru.", "Premium residential, mixed-use"],
      ["7", "Mahaveer Group", "Mid-tier Bengaluru developer focused on affordable-to-mid-segment apartments in peripheral growth corridors.", "Mid-segment apartments"],
      ["8", "SJR Group", "Established Bengaluru developer active in residential and commercial segments, incl. IT-corridor adjacent projects.", "Residential, commercial/IT-corridor"],
      ["9", "Sumadhura Group", "Six-project FY26 launch across East/North Bengaluru (Panathur, Rachenahalli, Devanahalli) — ₹5,500–6,000 Cr investment, ~90 acres, 4,000+ units incl. 400+ plots; also active in Hyderabad.", "Residential, plotted, multi-city (Bengaluru + Hyderabad)"],
      ["10", "DivyaSree Developers", "Bengaluru-origin developer with residential and commercial/IT-park legacy portfolio.", "Residential, commercial/IT parks"],
      ["11", "Total Environment", "Design-led boutique developer known for premium, low-density residential projects across Bengaluru.", "Premium boutique residential"],
      ["12", "Casagrand", "Chennai-headquartered major expanding into Bengaluru; among the most acquisitive South India developers by deal count nationally.", "Residential, multi-city (South India)"],
      ["13", "Birla Estates", "Recent entrant into Bengaluru via select premium project launches as part of national expansion.", "Premium residential, new entrant"],
      ["14", "Confident Group", "Bengaluru-focused developer with a broad mid-segment residential and villa portfolio.", "Mid-segment residential, villas"],
      ["15", "Bren Corporation", "Bengaluru mid-tier developer active in apartment and villa developments across peripheral corridors.", "Apartments, villas"],
      ["16", "Vaswani Group", "Established Bengaluru developer with residential and commercial project portfolio across multiple localities.", "Residential, commercial"],
    ],
    recommendations: [
      { head: "Land Acquisition Strategy", color: NAVY, items: [
        "Fast-track any B-Khata-to-A-Khata conversions on NBR-held parcels before the 23 Aug 2026 window closes (charge is 2% vs. standard 5% of guidance value).",
        "Evaluate BDA's Premium Sites e-Auction 2026 (175 sites, registration closes 27 Jul 2026) for infill opportunities in NBR's active corridors before Grade A players crowd the bidding.",
        "Commission a land-banking survey along the Bengaluru Business Corridor (ex-PRR) alignment now, ahead of final compensation notification — Godrej, Prestige and Puravankara are already active on adjacent corridors (Sarjapur, Whitefield, Devanahalli, Budigere) that ring NBR's Mullur landholding.",
        "Re-run acquisition underwriting models to reflect the Feb 2026 guidance-value hike (+6–15% Bengaluru urban) and the ~7.5–7.6% all-in stamp duty/registration/cess load.",
      ]},
      { head: "Legal & Compliance", color: RED_RISK, items: [
        "Where any NBR acquisition uses an industrial-area (KIADB) or similar special-purpose route, document public-purpose justification exhaustively in light of the IMTMA ruling that struck down a comparable acquisition as a \"colourable exercise of power.\"",
        "Complete e-Khata verification (BBMP e-Aasthi) on all unsold NBR inventory and land-bank parcels proactively — Kaveri 2.0's auto-khata-transfer on registration will otherwise expose gaps at the point of buyer registration.",
        "Audit all live and upcoming project possession-date commitments against realistic timelines; K-RERA/REAT is actively enforcing delay compensation and has imposed up to 5%-of-project-cost daily penalties for non-compliance on peer developers.",
      ]},
      { head: "Competitive Positioning", color: GREEN_POS, items: [
        "Benchmark launch pricing against Godrej Properties' new Sarjapur Road and East Bengaluru parcels and Puravankara's Sanna Ammanikere/airport-corridor project — both now bracket NBR's core operating micro-markets.",
        "Consider a capital-light JDA structure (per Brigade's Gunjur model) for any land NBR is negotiating with fragmented family/agricultural landholders, rather than outright purchase, to preserve balance-sheet flexibility.",
        "Track Sobha's Bengaluru-concentrated (vs. Godrej/DLF's multi-city) strategy as the closest scale comparable for a single-city-focused developer — useful reference points for NBR's own land-bank pacing (~10 msf/year benchmark).",
      ]},
    ],
  },

  layer2: {
    regional: [
      ["Tamil Nadu / Chennai", "TNRERA continues to drive buyer trust (ready-to-move demand +35% post-RERA); Jan 2026 guideline-value revision raised rates along OMR, ECR and GST Road growth corridors; stamp duty 7% + registration 4% (11% all-in) charged on the higher of guideline value or sale price.", "If NBR evaluates a Chennai entry or JV, price underwriting must use the higher of guideline/market value; OMR/ECR corridors carry the steepest revised guideline values."],
      ["Telangana / Hyderabad", "HYDRAA continues aggressive government-land reclamation (₹1,000 Cr near HITEX; ₹30,000 Cr / ~200 acres at Puppalaguda-Gandipet); 650 acres near Shamshabad Airport flagged as targeted by developers amid farmer protests; strict building-legality and lake-buffer enforcement ongoing.", "Any NBR land diligence in Hyderabad must independently verify government/lake-buffer land status beyond seller representations — HYDRAA enforcement risk is materially elevated versus Karnataka."],
      ["Andhra Pradesh / Amaravati", "Amaravati land-pooling Phase II (from 3 Jan 2026) targets 16,666.78 acres across 7 villages/9 units; ₹40,000/acre annual lease + ₹1.5L crop loans offered to ~31,000 farmers; earmarked for an International Sports City, greenfield airport, new rail station and Inner Ring Road.", "A long-horizon watch item rather than a near-term NBR opportunity; monitor IRR-linked land re-rating for possible future entry once Phase II parcels are notified for private development."],
      ["Kerala", "State government to constitute an expert committee to overhaul land laws — described as a potentially major policy shift; 2026 Government Land Assignment Rules now govern alternate-use permissions on assigned land; June 2026 market shows steady NRI-driven demand and improved digitised land records.", "Land-law reform is still at the committee stage — not yet actionable, but NBR's legal team should track committee recommendations given potential impact on assignment-land conversions."],
    ],
    deals: [
      ["Godrej Properties (national)", "Multi-city", "National FY26 land-deal leader with 17 deals / 443.5 acres, including South India parcels.", "17 deals / 443.5 acres"],
      ["Listed developers (South India, FY26)", "Chennai & Hyderabad", "5 land deals each in Chennai (74+ acres) and Hyderabad (~38 acres) among listed players in FY26; Coimbatore also saw listed-player activity.", "~112 acres combined (Chennai+Hyderabad)"],
      ["Casagrand", "Chennai (HQ) + Bengaluru expansion", "Among the most active South India-headquartered developers by deal count, expanding into Bengaluru.", "Multi-project, multi-city"],
      ["Sumadhura Group", "Bengaluru + Hyderabad", "Concurrent FY26 pipeline across both cities (₹5,500–6,000 Cr, ~90 acres) signals continued South India cross-city land banking by mid-tier players.", "~90 acres, ₹5,500–6,000 Cr"],
    ],
    recommendations: [
      { head: "Regional Positioning", color: NAVY, items: [
        "Treat Chennai (OMR/ECR/GST Road) and Hyderabad as medium-term watch markets rather than immediate expansion targets — both carry materially higher regulatory/compliance overhead (TN's 11% all-in stamp duty; Telangana's HYDRAA enforcement risk) than Karnataka.",
        "If a South India footprint beyond Karnataka is on NBR's roadmap, Hyderabad diligence must independently verify government/lake-buffer land status given HYDRAA's active reclamation campaign — do not rely solely on seller title representations.",
        "Monitor Amaravati Phase II land-pooling notifications as a long-horizon (3–5 year) opportunity once parcels are released for private development, given the scale of committed infrastructure investment (airport, rail, Inner Ring Road).",
      ]},
      { head: "Legal Watch", color: RED_RISK, items: [
        "Track Kerala's land-law expert committee recommendations — any reform to assignment-land conversion rules could open or close future acquisition routes in that state.",
        "Where NBR engages South India-based contractors, JV partners or landholders (e.g., in a Casagrand-style multi-city expansion posture), confirm TNRERA/TS-RERA/Kerala RERA registration status independently before any MoU is signed.",
      ]},
    ],
  },

  sources: [
    "Deccan Herald — Karnataka/Bengaluru state & city desk",
    "Business Standard — Companies & Capital Market News",
    "Knight Frank India — India Real Estate Report H1/H2 2026",
    "JLL India — Residential Dynamics Report Q1 2026 / Bengaluru Mass Residential Market Dynamics",
    "LiveLaw.in — Karnataka High Court coverage",
    "Mondaq / Khurana & Khurana — RERA judgment analysis",
    "Cyril Amarchand Mangaldas (India Corporate Law blog) — Karnataka Land Revenue Rules Amendment 2025",
    "Lexology / Fox Mandal — Karnataka land conversion & registration reform coverage",
    "Karnataka RERA (rera.karnataka.gov.in) — official registrations & orders",
    "TNREGINET — Tamil Nadu guideline value portal",
    "Kaveri Online (Karnataka Dept. of Stamps & Registration)",
    "Sakshi Post / Deccan Chronicle / Gulte — HYDRAA (Telangana) coverage",
    "PropNewz — Tamil Nadu stamp duty & guideline value analysis",
    "Kerala Realty / helloaddress.com — Kerala real estate market trend coverage",
    "Company disclosures/press releases — Prestige, Brigade, Sobha, Puravankara, Godrej Properties, Embassy, Sumadhura, Adarsh Developers",
  ],
};

// ---------------------------------------------------------------------------
// LAYOUT HELPERS
// ---------------------------------------------------------------------------
const FULL_WIDTH = 10466; // A4 portrait, 0.5in margins

function sectionBanner(text, opts = {}) {
  const { fill = NAVY, size = 26 } = opts;
  return new Paragraph({
    shading: { type: ShadingType.CLEAR, color: "auto", fill },
    spacing: { before: 280, after: 160 },
    children: [
      new TextRun({ text: `  ${text}`, bold: true, color: WHITE, size, font: "Calibri" }),
    ],
  });
}

function subHeading(text, color = NAVY) {
  return new Paragraph({
    spacing: { before: 220, after: 100 },
    border: { bottom: { color, space: 2, style: BorderStyle.SINGLE, size: 10 } },
    children: [new TextRun({ text, bold: true, color, size: 22, font: "Calibri" })],
  });
}

function bodyText(text, opts = {}) {
  const { italics = false, size = 19, color = TEXT_DARK, spacingAfter = 120 } = opts;
  return new Paragraph({
    spacing: { after: spacingAfter },
    children: [new TextRun({ text, italics, size, color, font: "Calibri" })],
  });
}

function bullet(text, opts = {}) {
  const { size = 19, color = TEXT_DARK } = opts;
  return new Paragraph({
    numbering: { reference: "digest-bullets", level: 0 },
    spacing: { after: 90 },
    children: [new TextRun({ text, size, color, font: "Calibri" })],
  });
}

function headerCell(text, width, opts = {}) {
  const { align = AlignmentType.LEFT } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold: true, color: WHITE, size: 18, font: "Calibri" })],
    })],
  });
}

function bodyCell(text, width, opts = {}) {
  const { fill = WHITE, bold = false, color = TEXT_DARK, align = AlignmentType.LEFT } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, color, size: 17, font: "Calibri" })],
    })],
  });
}

function buildTable(headers, widths, rows) {
  const headerRow = new TableRow({
    children: headers.map((h, i) => headerCell(h, widths[i])),
    tableHeader: true,
  });
  const bodyRows = rows.map((row, ridx) => new TableRow({
    children: row.map((cell, cidx) => bodyCell(cell, widths[cidx], {
      fill: ridx % 2 === 0 ? WHITE : GREY_BAND,
    })),
  }));
  return new Table({
    width: { size: FULL_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: [headerRow, ...bodyRows],
  });
}

function spacer(h = 120) {
  return new Paragraph({ spacing: { after: h }, children: [] });
}

function goldRule() {
  return new Paragraph({
    spacing: { after: 200 },
    border: { bottom: { color: GOLD, space: 1, style: BorderStyle.SINGLE, size: 24 } },
    children: [],
  });
}

// ---------------------------------------------------------------------------
// DOCUMENT ASSEMBLY
// ---------------------------------------------------------------------------
const children = [];

// --- Cover Page -------------------------------------------------------------
children.push(
  new Paragraph({
    shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY_DARK },
    alignment: AlignmentType.CENTER,
    spacing: { before: 600, after: 100 },
    children: [new TextRun({ text: "NBR GROUP", bold: true, color: GOLD, size: 40, font: "Calibri" })],
  }),
  new Paragraph({
    shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY_DARK },
    alignment: AlignmentType.CENTER,
    spacing: { after: 600 },
    children: [new TextRun({ text: "LEGAL DEPARTMENT  |  LAND ACQUISITION TEAM", bold: true, color: WHITE, size: 20, font: "Calibri" })],
  }),
  new Paragraph({ spacing: { before: 900 }, children: [] }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "DAILY MORNING", bold: true, color: NAVY, size: 48, font: "Calibri" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({ text: "LEGAL & LAND ACQUISITION DIGEST", bold: true, color: NAVY, size: 48, font: "Calibri" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    border: { top: { color: GOLD, space: 8, style: BorderStyle.SINGLE, size: 24 }, bottom: { color: GOLD, space: 8, style: BorderStyle.SINGLE, size: 24 } },
    spacing: { before: 200, after: 200 },
    children: [new TextRun({ text: `${CONTENT.dateLabel}   •   ${CONTENT.editionNo}`, bold: true, color: GOLD, size: 24, font: "Calibri" })],
  }),
  new Paragraph({ spacing: { before: 500 }, alignment: AlignmentType.CENTER, children: [
    new TextRun({ text: "LAYER 1", bold: true, color: WHITE, size: 22, font: "Calibri" }),
  ]}),
  new Paragraph({
    shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 160 },
    children: [new TextRun({ text: "Bengaluru & Karnataka — Intensive Focus", bold: true, color: WHITE, size: 22, font: "Calibri" })],
  }),
  new Paragraph({
    shading: { type: ShadingType.CLEAR, color: "auto", fill: GOLD },
    alignment: AlignmentType.CENTER,
    spacing: { after: 160 },
    children: [new TextRun({ text: "LAYER 2  —  South India Regional Watch", bold: true, color: NAVY_DARK, size: 22, font: "Calibri" })],
  }),
  new Paragraph({ spacing: { before: 900 }, alignment: AlignmentType.CENTER, children: [
    new TextRun({ text: "Prepared for Senior Management  |  Strictly Internal & Confidential", italics: true, color: TEXT_DARK, size: 18, font: "Calibri" }),
  ]}),
  new Paragraph({ children: [new (require("docx").PageBreak)()] }),
);

// --- Executive Summary -------------------------------------------------------
children.push(sectionBanner("EXECUTIVE SUMMARY — TOP SIGNALS TODAY"));
CONTENT.execSummary.forEach((item) => {
  children.push(new Paragraph({
    spacing: { after: 130 },
    children: [
      new TextRun({ text: `[${item.tag}]  `, bold: true, color: item.color, size: 19, font: "Calibri" }),
      new TextRun({ text: item.text, size: 19, color: TEXT_DARK, font: "Calibri" }),
    ],
  }));
});
children.push(spacer(80));

// --- LAYER 1 ------------------------------------------------------------------
children.push(sectionBanner("LAYER 1  —  BENGALURU & KARNATAKA INTENSIVE UPDATE", { fill: NAVY, size: 28 }));

children.push(subHeading("1.1  Regulatory & Legal Developments"));
children.push(buildTable(
  ["Topic", "Development", "Impact for NBR Group"],
  [2200, 5266, 3000],
  CONTENT.layer1.regulatory,
));
children.push(spacer());

children.push(subHeading("1.2  Land Acquisition & Government Land News"));
children.push(buildTable(
  ["Project / Area", "Development", "Financial / Scale"],
  [2200, 5266, 3000],
  CONTENT.layer1.landAcquisition,
));
children.push(spacer());

children.push(subHeading("1.3  Judicial Developments (Karnataka High Court)"));
children.push(buildTable(
  ["Case / Project", "Court Ruling & Holding", "Implication for NBR"],
  [2200, 5266, 3000],
  CONTENT.layer1.judicial,
));
children.push(spacer());

children.push(subHeading("1.4  Competitor Analysis — Top 8 Grade A Developers (Bengaluru)"));
children.push(buildTable(
  ["#", "Developer", "Recent Land / Legal Development", "Deal Size / GDV", "NBR Relevance"],
  [500, 1500, 5066, 1900, 1500],
  CONTENT.layer1.gradeA,
));
children.push(spacer());

children.push(subHeading("1.5  Competitor Watch — 16 Mid-Tier & Boutique Developers (Bengaluru)"));
children.push(buildTable(
  ["#", "Developer", "Snapshot / Recent Activity", "Segment Focus"],
  [500, 1800, 6166, 2000],
  CONTENT.layer1.midTier,
));
children.push(spacer());

children.push(subHeading("1.6  Recommendations for NBR Group — Bengaluru / Karnataka"));
CONTENT.layer1.recommendations.forEach((block) => {
  children.push(new Paragraph({
    spacing: { before: 140, after: 80 },
    children: [new TextRun({ text: block.head, bold: true, color: block.color, size: 20, font: "Calibri" })],
  }));
  block.items.forEach((it) => children.push(bullet(it)));
});
children.push(spacer());

// --- LAYER 2 ------------------------------------------------------------------
children.push(new Paragraph({ children: [new (require("docx").PageBreak)()] }));
children.push(sectionBanner("LAYER 2  —  SOUTH INDIA REGIONAL WATCH", { fill: GOLD, size: 28 }));
children.push(new Paragraph({
  spacing: { after: 160 },
  children: [new TextRun({ text: "Tamil Nadu  •  Telangana  •  Andhra Pradesh  •  Kerala", italics: true, color: NAVY, size: 19, font: "Calibri" })],
}));

children.push(subHeading("2.1  Regional Legal & Regulatory Developments"));
children.push(buildTable(
  ["State / City", "Development", "Impact for NBR Group"],
  [2200, 5266, 3000],
  CONTENT.layer2.regional,
));
children.push(spacer());

children.push(subHeading("2.2  Notable South India Land Deals"));
children.push(buildTable(
  ["Developer / Project", "Location", "Details", "Scale"],
  [2400, 2000, 4566, 1500],
  CONTENT.layer2.deals,
));
children.push(spacer());

children.push(subHeading("2.3  Recommendations for NBR Group — South India"));
CONTENT.layer2.recommendations.forEach((block) => {
  children.push(new Paragraph({
    spacing: { before: 140, after: 80 },
    children: [new TextRun({ text: block.head, bold: true, color: block.color, size: 20, font: "Calibri" })],
  }));
  block.items.forEach((it) => children.push(bullet(it)));
});
children.push(spacer());

// --- Sources ------------------------------------------------------------------
children.push(goldRule());
children.push(subHeading("Sources & Publications Referenced", NAVY));
CONTENT.sources.forEach((s) => children.push(bullet(s, { size: 17, color: "444444" })));

children.push(spacer(200));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({
    text: "This digest is compiled from publicly available industry publications, regulatory portals and news sources for internal briefing purposes only. It does not constitute legal advice.",
    italics: true, size: 15, color: "666666", font: "Calibri",
  })],
}));

// ---------------------------------------------------------------------------
// DOCUMENT DEFINITION
// ---------------------------------------------------------------------------
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "digest-bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: convertInchesToTwip(0.28), hanging: convertInchesToTwip(0.18) } } } },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          margin: { top: 720, bottom: 720, left: 720, right: 720 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { color: GOLD, space: 4, style: BorderStyle.SINGLE, size: 8 } },
            tabStops: [{ type: TabStopType.RIGHT, position: FULL_WIDTH }],
            children: [
              new TextRun({ text: "NBR GROUP — Legal & Land Acquisition Digest", bold: true, color: NAVY, size: 15, font: "Calibri" }),
              new TextRun({ text: `\t${CONTENT.dateLabel}`, color: "666666", size: 15, font: "Calibri" }),
            ],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "Strictly Private & Confidential — For Internal Circulation Only  |  Page ", size: 14, color: "666666", font: "Calibri" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 14, color: "666666", font: "Calibri" }),
              new TextRun({ text: " of ", size: 14, color: "666666", font: "Calibri" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 14, color: "666666", font: "Calibri" }),
            ],
          })],
        }),
      },
      children,
    },
  ],
});

// ---------------------------------------------------------------------------
// WRITE FILE
// ---------------------------------------------------------------------------
const dateForFile = process.env.DIGEST_DATE || new Date().toISOString().slice(0, 10);
const outDir = path.join(__dirname, "..", "output");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outPath = path.join(outDir, `NBR_Legal_Land_Digest_${dateForFile}.docx`);

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("Written:", outPath);
});
