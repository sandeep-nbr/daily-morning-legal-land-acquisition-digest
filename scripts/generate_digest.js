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
  dateLabel: "Thursday, 23 July 2026",
  editionNo: "Vol. I — Edition 002",

  execSummary: [
    { tag: "HIGH", color: RED_RISK, text: "Karnataka HC quashed a 10-acre Hosahalli land acquisition dating to 1989, citing fraud, suppression of facts and record manipulation — a fresh reminder (alongside the June IMTMA ruling) that even decades-old acquisitions remain vulnerable to title/procedural challenge." },
    { tag: "HIGH", color: AMBER_WATCH, text: "BBMP was formally dissolved (Sept 2025); the Greater Bengaluru Authority (GBA) plus five new city corporations now run civic administration — NBR's legal/compliance team should confirm continuity of e-Khata, property-tax and approval records under the new structure." },
    { tag: "WATCH", color: AMBER_WATCH, text: "BDA's Premium Sites e-Auction 2026 (175 sites) closes for registration in 4 days — 27 Jul 2026, 5 PM. Last call for NBR to evaluate sites in active corridors before the window shuts." },
    { tag: "WATCH", color: GREEN_POS, text: "Bengaluru Business Corridor (ex-PRR): 83% of Phase 1 land now acquired (as of Mar 2026); farmers pushing for 2014 Land Acquisition Act rates versus the state's offer of twice guidance value — a live compensation dispute worth tracking for corridor re-rating timing." },
    { tag: "INFO", color: NAVY, text: "K-RERA enforcement this week hit two developers NBR tracks directly: Adarsh Developers (₹79.84L delay-interest order + 60-day completion deadline on Adarsh Palm Acres Phase-2B) and Casagrand (₹52.74L GST-overcharge refund order) — a cautionary signal on possession timelines and GST billing practices." },
  ],

  layer1: {
    regulatory: [
      ["K-RERA Registrations & Scope", "K-RERA has now registered 28,000+ projects and 35,000+ agents since inception (up sharply from ~8,357/5,417 reported in July); K-RERA has also clarified that residential projects which applied for an Occupancy Certificate before 1 May 2017 are exempt from mandatory RERA registration.", "Confirm registration status (or exemption basis) for every NBR project before marketing; where an older project relies on the pre-2017 OC exemption, keep the OC-application-date evidence on file in case of a buyer challenge."],
      ["K-RERA Enforcement — Adarsh & Casagrand Orders", "K-RERA ordered Adarsh Nivas Pvt. Ltd. to pay ₹79.84 lakh delay-interest to a homebuyer and complete Adarsh Palm Acres Phase-2, Part-B within 60 days; separately, K-RERA ordered Casagrand Garden City to refund ₹52.74 lakh to a Bengaluru homebuyer over unfair GST charges.", "Two of NBR's tracked mid-tier peers were sanctioned in the same week — proactively audit NBR's own possession-date buffers and confirm GST is being passed through to buyers strictly per statute, not on an inflated or non-transparent basis."],
      ["BBMP → GBA Transition", "BBMP was officially dissolved in September 2025; the Greater Bengaluru Authority (GBA) and five new city corporations began functioning in its place for civic administration across Bengaluru.", "Verify with counsel that e-Khata records, property-tax PIDs and pending approval applications filed under the old BBMP regime carry over cleanly to the relevant GBA-era corporation — do not assume automatic continuity for parcels mid-process."],
      ["e-Khata / e-Aasthi / Kaveri 2.0", "Since July 2025 the BBMP/GBA e-Aasthi portal issues a single e-Khata PDF with QR-code verification (combining Certificate- and Extract-style data); Kaveri 2.0 continues to auto-trigger khata-transfer requests on registration.", "Pre-clear e-Khata status (QR-verifiable) on all NBR land parcels and unsold inventory ahead of buyer registration to avoid transfer delays under the new single-document format."],
      ["B-Khata → A-Khata Window", "Conversion charge remains cut from 5% to 2% of guidance value for the special window (15 May – 23 Aug 2026) — now roughly one month from closing.", "Fast-track conversion for any NBR-held or under-acquisition B-Khata parcels; with ~4 weeks left, this should now be treated as a near-term legal-team deadline, not a background item."],
      ["Guidance Value Revision", "The Feb 2026 revision (+6–15% Bengaluru urban, +4–8% rural) is in force; as of June 2026 a further proposed 10–15% increase remains un-gazetted, so current rates continue to apply pending formal notification.", "Consider sequencing any pending LOIs/registrations ahead of the next gazette notification — locking in acquisitions at today's guidance value could avoid a further step-up in stamp duty/registration cost."],
      ["Land Reforms Amendment 2025", "Karnataka Land Reforms & Certain Other Law (Amendment) Bill, 2025 expands Deputy Commissioners' Section 109 exemption power (0.5 ha → 4 ha, outside Bengaluru Urban/Rural); resale-after-7-years cases now need High-Power Committee (Chief Secretary) approval instead of State Government.", "For any agricultural-to-non-agricultural conversion outside core Bengaluru districts, exemption approvals may now be faster at DC level — legal team to reassess conversion timelines on outstation parcels."],
    ],
    landAcquisition: [
      ["BDA Premium Sites e-Auction 2026", "175 residential/corner/commercial sites; online auction opened 13 Jul 2026, registration and interest closes 5 PM, 27 Jul 2026 — four days from today.", "Time-critical: legal/land team to finalise any go/no-go on target sites in NBR's active micro-markets (Sarjapur, Yelahanka, Kanakapura) within the next 3 business days."],
      ["Bengaluru Business Corridor (ex-PRR)", "83% of Phase 1 land acquired as of March 2026; BDA needs ~2,560 acres in total but lacks liquid cash to pay upfront at current market rates; farmers want 2014 Land Acquisition Act-based compensation while government has offered twice the guidance value; Package 1 (20 km, ₹3,348 Cr) tenders are at final-award stage with two firms bidding; overall project cost ~₹26,786 Cr (~₹20,511 Cr for land); completion targeted around 2029.", "Corridor-adjacent land (North/East Bengaluru) remains a medium-term price-appreciation zone, but the farmer-compensation dispute could slow the remaining 17% of land acquisition — factor possible timeline slippage into any NBR land-banking survey along the alignment."],
      ["BDA Approval for 6,217-Acre Acquisition", "State government approved BDA's plan to acquire 6,217 acres across Bengaluru for planned development.", "Track layout notifications under this acquisition — potential TDR/compensation-linked land opportunities for NBR to evaluate on a case-by-case basis."],
      ["NICE Road Litigation Withdrawn", "NICE (Bangalore-Mysore Infrastructure Corridor) withdrew its land acquisition case in Karnataka HC after a Chief-Secretary-led mediation meeting.", "Illustrates state preference for negotiated settlement over prolonged litigation on infrastructure-linked land disputes — a useful precedent for NBR's own JDA negotiations with government-linked landholders."],
    ],
    judicial: [
      ["Hosahalli Village Acquisition Quashed", "Karnataka HC quashed a 10+ acre land acquisition in Hosahalli village, Bengaluru South Taluk — proceedings dating to 1989 — citing fraud, suppression of material facts and manipulation of official records; the court relied on the G.V.K. Rao Committee's findings (which had recommended the acquisition be dropped) and set aside a single-judge order that had treated the acquisition as final.", "Reinforces that title defects and procedural fraud can unwind even a nearly 40-year-old \"final\" acquisition — NBR's diligence on any parcel with a house-building cooperative society or old acquisition history in its chain of title should go beyond face-value finality and check for underlying committee findings or record irregularities."],
      ["IMTMA v. State of Karnataka (KIAD Act)", "Karnataka HC Division Bench set aside acquisition of land for a private entity's expansion, calling it a \"colourable exercise of power\" and \"fraud on the statute\"; 12 landowner appeals allowed.", "Any NBR acquisition routed through KIADB/industrial-area mechanisms must have unambiguous public-purpose documentation — this ruling raises the evidentiary bar and litigation risk for such routes."],
      ["Banashankari VI Stage (BDA)", "Division Bench upheld BDA's acquisition for the Banashankari VI Stage township, overturning a Jan 2025 single-bench order favouring landowners; ₹50,000 costs imposed on petitioners.", "Confirms courts are increasingly upholding BDA's public-purpose township acquisitions on appeal — relevant precedent if NBR is affected by adjacent BDA layout notifications."],
    ],
    gradeA: [
      ["1", "Prestige Group", "Added a 30+ acre parcel at KIADB Aerospace Park Phase 2 (North Bengaluru), sited near Brigade's upcoming second World Trade Center; also jointly took a 7.5-acre, 99-year lease parcel on Old Madras Road with Brigade Group for a commercial project. This follows the 21-acre Whitefield (₹450 Cr/₹4,500 Cr GDV) and Q1 FY26 102-acre multi-city (₹20,400 Cr GDV) buys already on file.", "₹20,400 Cr GDV (Q1 FY26, multi-city) + new Aerospace Park/OMR parcels", "Prestige is now both expanding its own aerospace-corridor land bank and co-investing with Brigade on capital-light commercial plays — watch for more Grade A joint-venture structures in North/East Bengaluru."],
      ["2", "Brigade Group", "Partnering with Prestige on the 7.5-acre Old Madras Road commercial lease (99-year) and building a second World Trade Center near Prestige's new Aerospace Park Phase 2 site; retains its Gunjur JDA (39-acre township, ₹7,200 Cr GDV) and Whitefield/ITPL parcels already on file.", "8 deals / ~81 acres FY26 (plus new JV parcel)", "Brigade's willingness to co-invest with a Grade A peer confirms a capital-light, alliance-based playbook NBR could adapt for parcels beyond its own balance-sheet capacity."],
      ["3", "Sobha Limited", "₹1,150–1,160 Cr FY26 land spend continues at ~10 million sq ft/year pace; East Bengaluru (Sobha Altair) and Hoskote World City township remain the active fronts — no new July land news beyond the established run-rate.", "~10 msf/yr acquisition target", "Bengaluru-concentrated strategy remains the closest scale comparable to NBR's own city-focused model."],
      ["4", "Puravankara Limited", "Godrej Fund Management has bought a land parcel in Bengaluru from Puravankara — a monetisation/divestment move rather than a fresh acquisition — alongside the previously reported Anekal, Mandur-Budigere, Doddagubbi and Sanna Ammanikere buys.", "~89 acres across 4 prior deals, ~₹9,000 Cr combined GDV, plus a fund-backed divestment", "Puravankara selectively monetising parcels to an institutional fund buyer signals a maturing secondary land market in Bengaluru — a potential capital-recycling model for NBR's own land bank."],
      ["5", "Godrej Properties", "Godrej Fund Management (Godrej's institutional-capital arm) acquired a Bengaluru land parcel from Puravankara, layering onto Godrej Properties' own 17-deal/443.5-acre FY26 run and its Sarjapur Road and East Bengaluru (NH-648/Whitefield) parcels already on file.", "17 deals / 443.5 acres (FY26, national) + fund-side acquisition from Puravankara", "Godrej is now active on Sarjapur Road both as a direct developer and, via its fund arm, as an institutional land buyer from peers — a two-track playbook worth benchmarking."],
      ["6", "Embassy Group / Embassy Developments", "An 18.6-acre residential project off Whitefield on SH-648 (reported ongoing from March 2026) is planned with six high-rise towers and 1,200+ apartments, alongside the previously reported ₹10,300 Cr FY26 North Bengaluru launch pipeline.", "₹10,300 Cr planned launches FY26 + 18.6-acre SH-648 project", "Embassy's SH-648/Whitefield project adds another large-scale East Bengaluru entrant alongside Prestige and Brigade — East Bengaluru corridor density is rising fast."],
      ["7", "Salarpuria Sattva Group", "Ongoing premium listings (Sattva Luxuria in Malleshwaram, Sattva Magnificia on Old Madras Road) continue to be refreshed on marketplace listings; no new land acquisition announced this week.", "Multiple ongoing premium launches", "Steady-state redevelopment/infill activity in mature localities — a model NBR could evaluate for infill parcels."],
      ["8", "Shriram Properties", "No fresh July land or legal news; remains active in plotted development and premium apartments per prior reporting.", "Plotted + apartment mix", "Plotted development remains Shriram's core strength — relevant benchmark if NBR expands its own plotted-development pipeline."],
    ],
    midTier: [
      ["1", "Century Real Estate", "Planning a ₹3,000 Cr mixed-use project on a 14-acre Outer Ring Road parcel (1.7 million sq ft residential across ~750 apartments plus 0.5 million sq ft commercial/retail), backed in part by a recent ₹1,850 Cr funding round tied to a ₹14,000 Cr overall GDV pipeline; also launching Century Attur Yelahanka (458 units).", "Mixed-use ORR development, land-bank monetisation"],
      ["2", "Assetz Property Group", "New pre-launch project \"Assetz Melodies of Life\" off Hosa Road and Sarjapur Road in Choodasandra — directly adjacent to NBR's own Sarjapur-Mullur footprint.", "Premium apartments, redevelopment — now with a direct Sarjapur-adjacent entrant"],
      ["3", "Vaishnavi Group", "No new July news; established Bengaluru residential/commercial/retail developer with a steady project pipeline.", "Residential, commercial, retail"],
      ["4", "Adarsh Developers", "K-RERA ordered Adarsh Nivas Pvt. Ltd. to pay ₹79.84 lakh delay-interest and complete Adarsh Palm Acres Phase-2, Part-B within 60 days — a live enforcement action alongside its Devanahalli plotted pipeline (Adarsh Savana Phase 3).", "Apartments, villas, plotted development — under active K-RERA enforcement"],
      ["5", "Concorde Group", "No new July news; long-standing mid-tier developer with residential and commercial portfolio across multiple micro-markets.", "Residential, commercial"],
      ["6", "Nitesh Estates", "No new July news; boutique/premium developer with high-end residential and hospitality-linked mixed-use history.", "Premium residential, mixed-use"],
      ["7", "Mahaveer Group", "No new July news; mid-tier developer focused on affordable-to-mid-segment apartments in peripheral growth corridors.", "Mid-segment apartments"],
      ["8", "SJR Group", "No new July news; active in residential and commercial segments, incl. IT-corridor adjacent projects.", "Residential, commercial/IT-corridor"],
      ["9", "Sumadhura Group", "Forayed formally into plotted development, targeting ~₹1,500 Cr topline from the new vertical; Sumadhura Elysium (Curve) at Panathur — 14.78 acres, 319 units from ₹2.5 Cr — has a Pre-RERA EOI with launch scheduled this month (July 2026).", "Residential, plotted (new vertical), multi-city (Bengaluru + Hyderabad)"],
      ["10", "DivyaSree Developers", "No new July news; Bengaluru-origin developer with residential and commercial/IT-park legacy portfolio.", "Residential, commercial/IT parks"],
      ["11", "Total Environment", "No new July news; design-led boutique developer known for premium, low-density residential projects.", "Premium boutique residential"],
      ["12", "Casagrand", "K-RERA ordered Casagrand Garden City to refund ₹52.74 lakh to a Bengaluru homebuyer over unfair GST charges; continues sales at Casagrand Moondance (Kumbalgodu) across budget-to-luxury unit mixes.", "Residential, multi-city (South India) — under active K-RERA enforcement"],
      ["13", "Birla Estates", "No new July news; recent entrant into Bengaluru via select premium project launches.", "Premium residential, new entrant"],
      ["14", "Confident Group", "No new July news; Bengaluru-focused mid-segment residential and villa portfolio.", "Mid-segment residential, villas"],
      ["15", "Bren Corporation", "No new July news; mid-tier developer active in apartment and villa developments across peripheral corridors.", "Apartments, villas"],
      ["16", "Vaswani Group", "No new July news; established Bengaluru developer with residential and commercial project portfolio.", "Residential, commercial"],
    ],
    recommendations: [
      { head: "Land Acquisition Strategy", color: NAVY, items: [
        "Time-critical: decide within the next 3 business days on any BDA Premium Sites e-Auction 2026 targets — registration/interest closes 5 PM, 27 Jul 2026.",
        "Assetz's new \"Melodies of Life\" pre-launch off Hosa Road/Sarjapur Road sits directly adjacent to NBR's own Sarjapur-Mullur footprint — treat this as a signal to accelerate, not defer, any parcel negotiations NBR has pending in that immediate radius.",
        "Sequence pending land LOIs and registrations ahead of the next (un-gazetted) guidance-value notification — locking in at today's rate avoids a further step-up once the proposed 10–15% hike is formalised.",
        "Continue the Bengaluru Business Corridor land-banking survey, but model in possible timeline slippage given the farmer-compensation dispute (2014 LA Act vs. twice-guidance-value) holding up the remaining ~17% of Phase 1 land.",
      ]},
      { head: "Legal & Compliance", color: RED_RISK, items: [
        "Confirm with counsel that e-Khata, property-tax and pending-approval records for any NBR parcel carry over cleanly from the dissolved BBMP to the relevant GBA-era city corporation — do not assume automatic continuity.",
        "Treat this week's Adarsh Developers and Casagrand K-RERA enforcement orders as a live cautionary signal: proactively audit NBR's own possession-date buffers against realistic construction timelines, and confirm GST is being passed through to buyers strictly per statute.",
        "The Hosahalli ruling shows even 1989-vintage \"final\" acquisitions can be unwound decades later on fraud/procedural grounds — extend title-diligence checks on any NBR parcel with an old acquisition or cooperative-society history beyond face-value finality.",
      ]},
      { head: "Competitive Positioning", color: GREEN_POS, items: [
        "Prestige and Brigade's joint 99-year-lease commercial parcel on Old Madras Road shows two Grade A peers now co-investing on capital-light structures — evaluate whether a similar JV/alliance model could unlock a larger NBR parcel beyond its own balance-sheet capacity.",
        "Godrej Fund Management's purchase of a Bengaluru parcel from Puravankara points to a maturing institutional-capital secondary market for land — consider fund-backed monetisation or JV capital partners as an option for recycling capital out of NBR's own land bank.",
        "Benchmark launch pricing against Godrej's Sarjapur Road presence (now reinforced by its fund arm's activity) and Assetz's new Sarjapur-adjacent pre-launch — both directly bracket NBR's core operating micro-market.",
      ]},
    ],
  },

  layer2: {
    regional: [
      ["Tamil Nadu / Chennai", "TNRERA continues to drive buyer trust (ready-to-move demand +35% post-RERA); Jan 2026 guideline-value revision raised rates along OMR, ECR and GST Road growth corridors; stamp duty 7% + registration 4% (11% all-in) charged on the higher of guideline value or sale price.", "If NBR evaluates a Chennai entry or JV, price underwriting must use the higher of guideline/market value; OMR/ECR corridors carry the steepest revised guideline values."],
      ["Telangana / Hyderabad", "Telangana HC (20 Jul 2026) restrained HYDRAA from taking coercive action on 26.03 acres in Bahadurguda village near Shamshabad — part of ~650 acres proposed for the bullet-train project — after HYDRAA entered the land with bulldozers/JCBs on 18 Jul and fenced it without court sanction; separately, HYDRAA continues lake-restoration work (Tammidikunta, Sunnam Cheruvu, due August) and general government-land reclamation.", "The Telangana HC's intervention is an early check on HYDRAA's enforcement reach, but does not eliminate the underlying risk — any NBR land diligence in Hyderabad must still independently verify government/lake-buffer land status beyond seller representations."],
      ["Andhra Pradesh / Amaravati", "President Droupadi Murmu gave assent (Apr 2026) to the Andhra Pradesh Reorganisation (Amendment) Bill, 2026, making Amaravati the sole and permanent state capital — removing a long-standing legal-finality risk; however, land allotments to spiritual/religious organisations under the Amaravati Land Allotment Regulations, 2017 (Regulation 6.6.1.2, \"social amenities\") have sparked controversy (16 Jul 2026) over CRDA's priorities for farmer-pooled land, alongside the ongoing Phase II pooling (16,666.78 acres).", "Amaravati's confirmed permanent-capital status is a modest long-horizon positive for future land bets there, but the allotment controversy is a reminder that CRDA-pooled land can be redirected to non-commercial \"social amenity\" uses — a relevant risk factor if NBR ever engages with CRDA-administered land."],
      ["Kerala", "No new July development beyond the previously reported land-law expert committee proposal and 2026 Government Land Assignment Rules; June 2026 market data continues to show steady NRI-driven demand and improved digitised land records.", "Land-law reform is still at the committee stage — not yet actionable, but NBR's legal team should keep tracking committee recommendations given potential impact on assignment-land conversions."],
    ],
    deals: [
      ["Godrej Properties / Godrej Fund Management", "Bengaluru (South India-wide franchise)", "Godrej Fund Management's purchase of a Bengaluru parcel from Puravankara illustrates the institutional-capital secondary market for land now active across Godrej's South India operations.", "Fund-backed acquisition from a peer developer"],
      ["Listed developers (South India, FY26)", "Chennai & Hyderabad", "5 land deals each in Chennai (74+ acres) and Hyderabad (~38 acres) among listed players in FY26; Coimbatore also saw listed-player activity.", "~112 acres combined (Chennai+Hyderabad)"],
      ["Casagrand", "Chennai (HQ) + Bengaluru expansion", "Continues multi-city project delivery (incl. Casagrand Moondance, Bengaluru) alongside a K-RERA GST-refund enforcement order in Bengaluru this week.", "Multi-project, multi-city — under K-RERA enforcement"],
      ["Sumadhura Group", "Bengaluru + Hyderabad", "New plotted-development vertical (~₹1,500 Cr topline target) adds to its concurrent Bengaluru/Hyderabad pipeline, reinforcing cross-city land banking by mid-tier South India players.", "New plotted vertical, ~₹1,500 Cr target"],
    ],
    recommendations: [
      { head: "Regional Positioning", color: NAVY, items: [
        "Treat Chennai (OMR/ECR/GST Road) and Hyderabad as medium-term watch markets rather than immediate expansion targets — Tamil Nadu's 11% all-in stamp duty and Telangana's HYDRAA enforcement risk both remain materially higher-friction than Karnataka, even after the Telangana HC's partial check on HYDRAA.",
        "Amaravati's now-confirmed permanent-capital status modestly de-risks the long-horizon opportunity there, but the fresh controversy over CRDA land allotments to non-commercial \"social amenity\" uses is a reminder to structure any future CRDA-linked engagement with clear end-use safeguards from the outset.",
        "If a South India footprint beyond Karnataka is on NBR's roadmap, Hyderabad diligence must independently verify government/lake-buffer land status given HYDRAA's active reclamation campaign — do not rely solely on seller title representations, even where HC intervention has occurred on a specific parcel elsewhere.",
      ]},
      { head: "Legal Watch", color: RED_RISK, items: [
        "Track Kerala's land-law expert committee recommendations — any reform to assignment-land conversion rules could open or close future acquisition routes in that state.",
        "Where NBR engages South India-based contractors, JV partners or landholders (e.g., in a Casagrand-style multi-city expansion posture), confirm TNRERA/TS-RERA/Kerala RERA registration status independently before any MoU is signed — and note that even established South India players (Casagrand) are currently under active K-RERA enforcement.",
      ]},
    ],
  },

  sources: [
    "Deccan Herald — Karnataka/Bengaluru state & city desk",
    "Business Standard — Companies & Capital Market News",
    "LiveLawBiz — RERA Cases Weekly Digest (29 Jun–5 Jul 2026)",
    "The420.in — Karnataka High Court / Hosahalli land-fraud coverage",
    "LiveLaw.in — Karnataka High Court coverage",
    "Mondaq / Khurana & Khurana — RERA judgment analysis",
    "Karnataka RERA (rera.karnataka.gov.in) — official registrations & orders",
    "Kaveri Online (Karnataka Dept. of Stamps & Registration)",
    "TNREGINET — Tamil Nadu guideline value portal",
    "The News Minute / Siasat / NewsMeter — Telangana HC–HYDRAA (Bahadurguda) coverage",
    "The Federal — Amaravati land-allotment controversy coverage",
    "NewsOnAir (All India Radio) — Andhra Pradesh Reorganisation (Amendment) Bill, 2026",
    "Republic World / Outlook India — NBR Group company profile coverage",
    "Constrofacilitator — Godrej Fund Management / Puravankara land-deal coverage",
    "Company disclosures/press releases — Prestige, Brigade, Sobha, Puravankara, Godrej Properties, Embassy, Sumadhura, Century Real Estate, Assetz, Casagrand, Adarsh Developers",
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
