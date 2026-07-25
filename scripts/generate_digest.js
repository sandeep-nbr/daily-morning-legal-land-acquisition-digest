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
  dateLabel: "Saturday, 25 July 2026",
  editionNo: "Vol. I — Edition 004",

  execSummary: [
    { tag: "HIGH", color: RED_RISK, text: "The Bengaluru Business Corridor's original 2008-09 compensation award (₹20–30 lakh/acre) is colliding with today's Varthur–Gunjur land demands of ₹15–40 Cr/acre — BDA is now offering farmers five compensation options (cash, residential plot, commercial plot, TDR, FAR) and aims to secure 80% of the land within 4 months." },
    { tag: "HIGH", color: AMBER_WATCH, text: "BDA's Premium Sites e-Auction 2026 (its 13th auction this year — 175 sites) still closes for registration 27 Jul 5 PM, live bidding to 31 Jul — the pace of BDA auctions this year (13 rounds) signals sustained public land supply into the market." },
    { tag: "WATCH", color: GREEN_POS, text: "Godrej Properties posted back-to-back launch-week blowouts in Bengaluru — Godrej Vanantara sold ₹2,000 Cr of homes in its launch week, and Godrej Woodscapes (Whitefield-Budigere Cross) has sold over ₹3,150 Cr — a strong demand-side signal for NBR's own upcoming launches in adjacent corridors." },
    { tag: "WATCH", color: AMBER_WATCH, text: "N. Jayaram, a retired 2004-batch IAS officer and former Secretary to the Chief Minister, has been appointed a K-RERA Member for a 5-year term — a leadership refresh worth noting for any pending or upcoming NBR interactions with the authority." },
    { tag: "INFO", color: NAVY, text: "Karnataka HC has ruled that courts cannot use Section 34 of the Arbitration Act to modify land acquisition awards or fix compensation under the National Highways Act (P. Nagaraju v. Special LAO) — relevant precedent for any NBR parcel affected by NH-linked acquisition or compensation disputes." },
  ],

  layer1: {
    regulatory: [
      ["K-RERA Registrations & Scope", "K-RERA has now registered 28,000+ projects and 35,000+ agents since inception; residential projects which applied for an Occupancy Certificate before 1 May 2017 remain exempt from mandatory RERA registration; N. Jayaram (retired 2004-batch IAS officer, former Secretary to CM Siddaramaiah) has been appointed a K-RERA Member for a 5-year term.", "Confirm registration status (or exemption basis) for every NBR project before marketing; note the K-RERA Member bench refresh for any pending or upcoming NBR matters before the authority."],
      ["K-RERA Enforcement — Adarsh & Casagrand Orders", "K-RERA ordered Adarsh Nivas Pvt. Ltd. to pay ₹79.84 lakh delay-interest to a homebuyer and complete Adarsh Palm Acres Phase-2, Part-B within 60 days; separately, K-RERA ordered Casagrand Garden City to refund ₹52.74 lakh to a Bengaluru homebuyer over unfair GST charges.", "Two of NBR's tracked mid-tier peers were sanctioned in the same week — proactively audit NBR's own possession-date buffers and confirm GST is being passed through to buyers strictly per statute, not on an inflated or non-transparent basis."],
      ["BBMP → GBA Transition", "BBMP was officially dissolved in September 2025; GBA plus five new city corporations are in a \"Phased Implementation\" stage (2025–2026) — merging the corporations' databases, clearing a ₹437 Cr backlog of unpaid civic dues, and finalising the 2026 Master Plan. Election timing carries some ambiguity in current reporting: the Supreme Court directed completion by 30 June 2026, while other coverage still cites an August 2026 target and the state has said there is \"no delay\" despite a pending GBA Act amendment.", "Verify with counsel that e-Khata records, property-tax PIDs and pending approval applications filed under the old BBMP regime carry over cleanly during the database-merger window — treat the uncertain election timeline as a reason for extra diligence on record continuity, not a formality."],
      ["e-Khata / e-Aasthi / Kaveri 2.0", "Since July 2025 the BBMP/GBA e-Aasthi portal issues a single e-Khata PDF with QR-code verification (combining Certificate- and Extract-style data); Kaveri 2.0 continues to auto-trigger khata-transfer requests on registration.", "Pre-clear e-Khata status (QR-verifiable) on all NBR land parcels and unsold inventory ahead of buyer registration to avoid transfer delays under the new single-document format."],
      ["B-Khata → A-Khata Window", "Conversion charge remains cut from 5% to 2% of guidance value for the special window (15 May – 23 Aug 2026) — now roughly one month from closing.", "Fast-track conversion for any NBR-held or under-acquisition B-Khata parcels; with ~4 weeks left, this should now be treated as a near-term legal-team deadline, not a background item."],
      ["Guidance Value Revision", "The Feb 2026 revision (+6–15% Bengaluru urban, +4–8% rural) is in force; as of June 2026 a further proposed 10–15% increase remains un-gazetted, so current rates continue to apply pending formal notification.", "Consider sequencing any pending LOIs/registrations ahead of the next gazette notification — locking in acquisitions at today's guidance value could avoid a further step-up in stamp duty/registration cost."],
      ["Land Reforms Amendment 2025", "Karnataka Land Reforms & Certain Other Law (Amendment) Bill, 2025 expands Deputy Commissioners' Section 109 exemption power (0.5 ha → 4 ha, outside Bengaluru Urban/Rural); resale-after-7-years cases now need High-Power Committee (Chief Secretary) approval instead of State Government.", "For any agricultural-to-non-agricultural conversion outside core Bengaluru districts, exemption approvals may now be faster at DC level — legal team to reassess conversion timelines on outstation parcels."],
    ],
    landAcquisition: [
      ["BDA Premium Sites e-Auction 2026", "175 residential/corner/commercial sites; registration/interest closes 5 PM, 27 Jul 2026 (EMD ₹4 lakh/site payable on the government portal before live bidding); live bidding then runs to 5 PM, 31 Jul 2026, with a minimum increment of ₹500/sq. mtr.", "Time-critical: legal/land team has roughly one week end-to-end — finalise go/no-go on target sites in NBR's active micro-markets (Sarjapur, Yelahanka, Kanakapura) and arrange EMD funding now to keep the 27 Jul registration deadline in reach."],
      ["Bengaluru Business Corridor (ex-PRR)", "The original 2008-09 award of ₹20–30 lakh/acre (for 948 of the 1,810 phase-1 acres) is now badly out of step with today's Varthur–Gunjur land demands of ₹15–40 Cr/acre. BDA is offering farmers five compensation routes — cash, residential plot, commercial plot, TDR, or FAR — and is targeting 80% land acquisition within 4 months; BDA needs ~2,560 acres in total (~₹20,511 Cr for land); Package 1 (20 km, ₹3,348 Cr) tenders are at final-award stage; completion targeted around 2029.", "The scale of the valuation gap (roughly 50-100x the original award) explains why compensation negotiations remain the critical path — NBR's land-banking survey along the BBC alignment should model a range of final settlement outcomes rather than a single compensation assumption."],
      ["BDA Approval for 6,217-Acre Acquisition", "State government approved BDA's plan to acquire 6,217 acres across Bengaluru for planned development.", "Track layout notifications under this acquisition — potential TDR/compensation-linked land opportunities for NBR to evaluate on a case-by-case basis."],
      ["NICE Road Litigation Withdrawn", "NICE (Bangalore-Mysore Infrastructure Corridor) withdrew its land acquisition case in Karnataka HC after a Chief-Secretary-led mediation meeting.", "Illustrates state preference for negotiated settlement over prolonged litigation on infrastructure-linked land disputes — a useful precedent for NBR's own JDA negotiations with government-linked landholders."],
    ],
    judicial: [
      ["Hosahalli Village Acquisition Quashed", "Karnataka HC quashed a 10+ acre land acquisition in Hosahalli village, Bengaluru South Taluk — proceedings dating to 1989 — citing fraud, suppression of material facts and manipulation of official records; the court relied on the G.V.K. Rao Committee's findings (which had recommended the acquisition be dropped) and set aside a single-judge order that had treated the acquisition as final.", "Reinforces that title defects and procedural fraud can unwind even a nearly 40-year-old \"final\" acquisition — NBR's diligence on any parcel with a house-building cooperative society or old acquisition history in its chain of title should go beyond face-value finality and check for underlying committee findings or record irregularities."],
      ["IMTMA v. State of Karnataka (KIAD Act)", "Karnataka HC Division Bench set aside acquisition of land for a private entity's expansion, calling it a \"colourable exercise of power\" and \"fraud on the statute\"; 12 landowner appeals allowed.", "Any NBR acquisition routed through KIADB/industrial-area mechanisms must have unambiguous public-purpose documentation — this ruling raises the evidentiary bar and litigation risk for such routes."],
      ["Banashankari VI Stage (BDA)", "Division Bench upheld BDA's acquisition for the Banashankari VI Stage township, overturning a Jan 2025 single-bench order favouring landowners; ₹50,000 costs imposed on petitioners.", "Confirms courts are increasingly upholding BDA's public-purpose township acquisitions on appeal — relevant precedent if NBR is affected by adjacent BDA layout notifications."],
      ["Bidadi Integrated Township — PIL Dismissed", "Karnataka HC dismissed a PIL challenging acquisition of 516 acres for the Bidadi Integrated Township Project, finding no legal violations and confirming the process followed the Right to Fair Compensation Act (RFCTLARR), 2013.", "A useful counterpoint to Hosahalli: acquisitions run strictly per RFCTLARR procedure and documentation are being upheld even at large scale (516 acres) — reinforces that clean process, not just age or size, is what determines litigation survivability."],
      ["P. Nagaraju v. Special LAO (NH Act Compensation)", "Karnataka HC ruled that courts cannot invoke Section 34 of the Arbitration Act to modify a land acquisition award or fix compensation under the National Highways Act — compensation disputes under that Act must follow their own statutory route, not general arbitration-law challenge.", "Relevant if any NBR parcel is ever affected by a National Highways Act acquisition (e.g., near a highway-widening project) — confirms the correct legal remedy is the Act's own mechanism, not an arbitration-style challenge, which affects how legal budgets and timelines for such disputes should be planned."],
    ],
    gradeA: [
      ["1", "Prestige Group", "Plans a large integrated destination within Bengaluru Airport City in partnership with Bengaluru Airport City itself — a ₹1,800 Cr investment with construction expected to start early 2027 — alongside its Aerospace Park Phase 2 (30+ acres) and Old Madras Road JV parcel already on file.", "₹20,400 Cr GDV (Q1 FY26, multi-city) + ₹1,800 Cr Airport City destination", "Prestige's Airport City tie-up is a new large-scale, partnership-anchored format (vs. outright land purchase) — worth tracking as a template for institution-anchored development deals."],
      ["2", "Brigade Group", "Acquired a further 20.19-acre parcel in the Whitefield–Hoskote corridor for ₹588.33 Cr through subsidiary Ananthay Properties, in addition to partnering with Prestige on the 7.5-acre Old Madras Road commercial lease (99-year) and its Gunjur JDA (39-acre township, ₹7,200 Cr GDV).", "8+ deals / ~101 acres FY26 (incl. new Whitefield-Hoskote parcel)", "Brigade is now adding scale on its own (Whitefield-Hoskote) even as it co-invests with Prestige elsewhere — a dual-track land strategy worth benchmarking alongside its JV playbook."],
      ["3", "Sobha Limited", "₹1,150–1,160 Cr FY26 land spend continues at ~10 million sq ft/year pace (East Bengaluru's Sobha Altair, Hoskote World City); separately, Godrej Industries Group's family office (Anamudi Real Estates LLP) divested a 4.4% stake in Sobha for ₹858 Cr via open-market transactions.", "~10 msf/yr acquisition target + ₹858 Cr Godrej-family-office stake sale", "The stake divestment is a shareholding/ownership event, not an operational change, but it's a reminder that Sobha's cap table includes cross-holdings with Godrej — a dynamic worth monitoring given Godrej Properties' own expansion in Sobha's core Bengaluru turf."],
      ["4", "Puravankara Limited", "Signed a fresh JDA for a 6.4-acre parcel in Sarjapur, Bengaluru — 0.8 msf saleable area, ~₹1,000 Cr revenue potential — directly inside NBR's own home corridor; also secured (via subsidiary) a ₹133.34 Cr LOI for civil/structural works on Embassy's Eden project. Godrej Fund Management separately bought a Bengaluru parcel from Puravankara.", "~95 acres across 5 deals, ~₹10,000 Cr combined GDV, plus a fund-backed divestment", "Puravankara is now developing inside the Sarjapur micro-market NBR itself operates in, not just its periphery — this is the most direct competitive overlap on this list and warrants immediate pricing/positioning review."],
      ["5", "Godrej Properties", "Strong Bengaluru sales momentum: Godrej Vanantara sold ₹2,000 Cr of homes (1,000+ units) in its launch week (~3.53 msf developable, ~₹3,700 Cr revenue potential), and Godrej Woodscapes (Whitefield-Budigere Cross) has sold 2,000+ homes worth over ₹3,150 Cr — on top of its 17-deal/443.5-acre FY26 land run and Sarjapur Road/East Bengaluru parcels already on file.", "17 deals / 443.5 acres (FY26, national) + ₹5,150 Cr+ in recent Bengaluru launch-week sales", "Godrej's launch-week sales velocity in Whitefield-Budigere and other East Bengaluru corridors is a strong demand-side benchmark — worth comparing directly against NBR's own absorption rates in similar micro-markets."],
      ["6", "Embassy Group / Embassy Developments", "An 18.6-acre residential project off Whitefield on SH-648 (reported ongoing from March 2026) is planned with six high-rise towers and 1,200+ apartments; its Embassy Eden project has now engaged Puravankara's construction arm (₹133.34 Cr LOI, 34-month execution) alongside the previously reported ₹10,300 Cr FY26 North Bengaluru launch pipeline.", "₹10,300 Cr planned launches FY26 + 18.6-acre SH-648 project", "Embassy's SH-648/Whitefield project adds another large-scale East Bengaluru entrant alongside Prestige and Brigade — East Bengaluru corridor density is rising fast."],
      ["7", "Salarpuria Sattva Group", "Pre-launched Sattva Vasanta Cove, a 1-acre residential development within its flagship Sattva Urbana township in North Bengaluru, alongside ongoing premium listings (Sattva Luxuria, Sattva Magnificia).", "Multiple ongoing premium launches, incl. new Sattva Urbana infill", "Steady-state redevelopment/infill activity (including infill within its own existing townships) in mature localities — a model NBR could evaluate for infill parcels."],
      ["8", "Shriram Properties", "No fresh July land or legal news; remains active in plotted development and premium apartments per prior reporting.", "Plotted + apartment mix", "Plotted development remains Shriram's core strength — relevant benchmark if NBR expands its own plotted-development pipeline."],
    ],
    midTier: [
      ["1", "Century Real Estate", "Planning a ₹3,000 Cr mixed-use project on a 14-acre Outer Ring Road parcel (1.7 million sq ft residential across ~750 apartments plus 0.5 million sq ft commercial/retail), backed in part by a recent ₹1,850 Cr funding round tied to a ₹14,000 Cr overall GDV pipeline; also launching Century Attur Yelahanka (458 units).", "Mixed-use ORR development, land-bank monetisation"],
      ["2", "Assetz Property Group", "New pre-launch project \"Assetz Melodies of Life\" off Hosa Road and Sarjapur Road in Choodasandra — directly adjacent to NBR's own Sarjapur-Mullur footprint.", "Premium apartments, redevelopment — now with a direct Sarjapur-adjacent entrant"],
      ["3", "Vaishnavi Group", "No new July news; established Bengaluru residential/commercial/retail developer with a steady project pipeline.", "Residential, commercial, retail"],
      ["4", "Adarsh Developers", "Proposed a new 12-acre residential project on CDP Road, Varthur (11 towers, G+25, 1,248 units planned), even as K-RERA's order to pay ₹79.84 lakh delay-interest and complete Adarsh Palm Acres Phase-2, Part-B within 60 days remains a live enforcement action.", "Apartments, villas, plotted development — expanding into East Bengaluru (Varthur) while under active K-RERA enforcement"],
      ["5", "Concorde Group", "No new July news; long-standing mid-tier developer with residential and commercial portfolio across multiple micro-markets.", "Residential, commercial"],
      ["6", "Nitesh Estates", "No new July news; boutique/premium developer with high-end residential and hospitality-linked mixed-use history.", "Premium residential, mixed-use"],
      ["7", "Mahaveer Group", "No new July news; mid-tier developer focused on affordable-to-mid-segment apartments in peripheral growth corridors.", "Mid-segment apartments"],
      ["8", "SJR Group", "No new July news; active in residential and commercial segments, incl. IT-corridor adjacent projects.", "Residential, commercial/IT-corridor"],
      ["9", "Sumadhura Group", "Acquired 40 acres, setting up a residential pipeline with revenue potential of up to ₹6,000 Cr; new launch on Soukya Road, East Bengaluru — 3 towers, 523 units (2/3 BHK) from ₹90 lakh — on top of its plotted-development foray (Panathur) already on file.", "Residential, plotted, multi-city (Bengaluru + Hyderabad) — now with a 40-acre pipeline addition"],
      ["10", "DivyaSree Developers", "No new July news; Bengaluru-origin developer with residential and commercial/IT-park legacy portfolio.", "Residential, commercial/IT parks"],
      ["11", "Total Environment", "No new July news; design-led boutique developer known for premium, low-density residential projects.", "Premium boutique residential"],
      ["12", "Casagrand", "Won Best Landscape Project (17th Realty+ Excellence Awards, Pune 2026) and Best Residential/Best Affordable Project (ET Real Estate Conclave & Awards 2026), even as K-RERA's ₹52.74 lakh GST-overcharge refund order against Casagrand Garden City remains a live enforcement action.", "Residential, multi-city (South India) — award-winning portfolio, but under active K-RERA enforcement"],
      ["13", "Birla Estates", "Aditya Birla Real Estate's wholly-owned subsidiary Birla Estates has entered a joint-venture agreement with Mitsubishi Estate Co. (MEC) to develop a residential project in Southeast Bengaluru.", "Premium residential, new entrant — now with an international JV partner (Mitsubishi Estate)"],
      ["14", "Confident Group", "No new July news; Bengaluru-focused mid-segment residential and villa portfolio.", "Mid-segment residential, villas"],
      ["15", "Bren Corporation", "No new July news; mid-tier developer active in apartment and villa developments across peripheral corridors.", "Apartments, villas"],
      ["16", "Vaswani Group", "No new July news; established Bengaluru developer with residential and commercial project portfolio.", "Residential, commercial"],
    ],
    recommendations: [
      { head: "Land Acquisition Strategy", color: NAVY, items: [
        "Time-critical: arrange EMD funding (₹4L/site) and finalise go/no-go on BDA Premium Sites e-Auction 2026 targets now — registration closes 27 Jul, live bidding runs to 31 Jul.",
        "Model a range of Bengaluru Business Corridor settlement outcomes, not a single figure — the gap between the original 2008-09 award (₹20-30L/acre) and current Varthur-Gunjur asks (₹15-40 Cr/acre) is wide enough that final compensation terms could land almost anywhere within it, materially affecting corridor-adjacent land economics.",
        "Puravankara's Sarjapur JDA and Assetz's Sarjapur-adjacent pre-launch remain an urgent signal to accelerate NBR's own pending Sarjapur-Mullur parcel negotiations rather than defer them.",
        "Sequence pending land LOIs and registrations ahead of the next (un-gazetted) guidance-value notification — locking in at today's rate avoids a further step-up once the proposed 10–15% hike is formalised.",
      ]},
      { head: "Legal & Compliance", color: RED_RISK, items: [
        "Given the ambiguity around the GBA election timeline (Supreme Court's 30 Jun deadline vs. an August target still being cited), treat e-Khata/property-tax record continuity as an open diligence item on every NBR parcel until the database-merger and election process is confirmed complete — do not assume continuity.",
        "Continue treating the Adarsh Developers and Casagrand K-RERA enforcement orders as a live cautionary signal: audit NBR's own possession-date buffers and GST pass-through practices proactively.",
        "Note the new P. Nagaraju precedent (Section 34 Arbitration Act cannot modify NH Act compensation) alongside the Hosahalli-vs-Bidadi contrast — for any NBR parcel touched by highway or public-purpose acquisition, use the correct statutory remedy route from the outset rather than a general arbitration-style challenge.",
      ]},
      { head: "Competitive Positioning", color: GREEN_POS, items: [
        "Puravankara developing inside Sarjapur itself remains the most direct competitive overlap identified to date — commission an immediate pricing and positioning comparison against this specific JDA project.",
        "Benchmark NBR's own absorption rates against Godrej Properties' recent Bengaluru launch-week sales velocity (₹2,000 Cr+ and ₹3,150 Cr+ in two separate launches) — both a demand-side reality check and a pricing reference point for East Bengaluru corridors.",
        "Prestige's Bengaluru Airport City tie-up and Birla Estates' JV with Mitsubishi Estate both show Grade A/emerging players increasingly using institutional or international partnership structures rather than outright land purchase — worth evaluating as an alternative to NBR's own balance-sheet-funded acquisitions.",
      ]},
    ],
  },

  layer2: {
    regional: [
      ["Tamil Nadu / Chennai", "TNRERA continues to drive buyer trust (ready-to-move demand +35% post-RERA); Jan 2026 guideline-value revision raised rates along OMR, ECR and GST Road growth corridors; stamp duty 7% + registration 4% (11% all-in) charged on the higher of guideline value or sale price.", "If NBR evaluates a Chennai entry or JV, price underwriting must use the higher of guideline/market value; OMR/ECR corridors carry the steepest revised guideline values."],
      ["Telangana / Hyderabad", "Telangana Cabinet approved merging 27 municipalities/corporations abutting the Outer Ring Road into the Greater Hyderabad Municipal Corporation — a GHMC footprint expansion echoing Bengaluru's own BBMP→GBA restructuring; this follows the Telangana HC's 20 Jul order restraining HYDRAA from coercive action on 26.03 acres in Bahadurguda pending the bullet-train land dispute, while HYDRAA continues lake-restoration and general reclamation work elsewhere.", "Hyderabad is undergoing a governance boundary change of its own (GHMC expansion) at the same time Bengaluru is mid-transition (GBA) — any NBR diligence on a Hyderabad parcel near the ORR should specifically check which civic body will hold jurisdiction post-merger before relying on current approvals."],
      ["Andhra Pradesh / Amaravati", "The State Revenue Minister reaffirmed there will be no changes to the Amaravati capital region; a fresh government order targets acquisition of 20,494 acres in total (including 16,562.56 acres of patta/private land) for key infrastructure — an International Sports City, greenfield international airport and related projects — layering onto the confirmed permanent-capital status (President's April 2026 assent) and the ongoing Phase II pooling (16,666.78 acres).", "The scale of the new patta-land acquisition order (16,562+ acres of private land) is a reminder that Amaravati's build-out will keep generating fresh land-acquisition and compensation activity for years — a long-horizon watch item, not yet an actionable NBR opportunity."],
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
    "Business Standard — Puravankara Sarjapur JDA and Embassy Eden LOI coverage",
    "BookNewProperty / RealtyPromoo — Brigade Whitefield-Hoskote and Adarsh Varthur project coverage",
    "Deccan Herald — Bidadi Township PIL and Bengaluru Business Corridor award coverage",
    "TheRealtyToday / BookNewProperty — Greater Bengaluru Authority governance-model coverage",
    "eauctionsindia.com / eauctioninfo.com — BDA Premium Sites e-Auction 2026 timeline",
    "Telangana Today / NewsOnAir — GHMC-ORR municipality merger coverage",
    "Star of Mysore — K-RERA Member appointment coverage",
    "Adredge — Karnataka HC Section 34/National Highways Act compensation ruling analysis",
    "NewsFirst / Deccan Herald — Bengaluru Business Corridor compensation-gap coverage",
    "Whalesbook / Business Standard — Godrej Vanantara and Godrej Woodscapes launch-week sales coverage",
    "Business Standard — Sobha stake divestment (Anamudi Real Estates/Godrej Industries Group) coverage",
    "Business Standard — Aditya Birla Real Estate / Birla Estates–Mitsubishi Estate JV coverage",
    "Company disclosures/press releases — Prestige, Brigade, Sobha, Puravankara, Godrej Properties, Embassy, Sattva, Sumadhura, Century Real Estate, Assetz, Casagrand, Adarsh Developers, Birla Estates",
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
