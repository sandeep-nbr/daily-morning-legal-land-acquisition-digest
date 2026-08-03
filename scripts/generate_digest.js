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
  dateLabel: "Monday, 3 August 2026",
  editionNo: "Vol. I — Edition 013",

  execSummary: [
    { tag: "HIGH", color: RED_RISK, text: "The optimistic BBC momentum tracked over the past several editions has a real counter-current: multiple outlets report the project \"stalled\"/\"slowed\" as farmers continue to \"stick to legal compensation\" demands rather than accept BDA's negotiated offers — the state has now formed a NEW high-level committee (chaired by the Additional Chief Secretary, Urban Development, with the BDA Commissioner and Law/Finance/Revenue principal secretaries) specifically to break the deadlock." },
    { tag: "WATCH", color: GREEN_POS, text: "Karnataka HC (Mar 2026) has clarified that RERA and the Karnataka Apartment Ownership Act, 1972 are not repugnant — the two statutes govern different stages of a project's life (RERA pre-completion, KAOA post-completion) and both apply, not one or the other." },
    { tag: "INFO", color: NAVY, text: "The K-RERA retroactive-delay-fee ruling flagged earlier is now confirmed to have struck down an entire K-RERA CIRCULAR (not just one case's fee) for lacking legislative backing — a broader precedent than previously understood, applicable to any developer, not just the one case." },
    { tag: "INFO", color: NAVY, text: "\"SNC,\" the lowest bidder (L1) for BBC Package 1 flagged in earlier editions, is now confirmed as Shankara Narayana Constructions (bidding alongside Vishwa Samudra Engineering) — resolving an abbreviation this digest had left unexpanded." },
    { tag: "INFO", color: NAVY, text: "No new NBR-specific news or fresh Grade A/mid-tier land deals since yesterday; the B-Khata→A-Khata window remains open through 23 Aug and the guidance-value hike remains un-gazetted." },
  ],

  layer1: {
    regulatory: [
      ["K-RERA Registrations & Scope", "K-RERA has now registered 28,000+ projects and 35,000+ agents since inception; residential projects which applied for an Occupancy Certificate before 1 May 2017 remain exempt from mandatory RERA registration; N. Jayaram has been appointed a K-RERA Member for a 5-year term. The underlying dispute over BDA's RERA status now has two pulling-apart rulings: the Karnataka RERA Appellate Tribunal (KREAT, Mar 2026) classified BDA as a \"promoter\" under RERA and ordered it to register the Nadaprabhu Kempegowda Layout project, while Karnataka HC (3 Jun 2026) stayed RERA applicability to BDA layouts with pre-2016 land acquisition notifications — interim relief ran only until 6 Jul 2026, a date now passed without a confirmed follow-up order.", "Confirm registration status (or exemption basis) for every NBR project before marketing. Treat BDA's RERA-promoter status as genuinely contested rather than settled either way — the Tribunal and the HC stay point in opposite directions, and NBR should watch for the HC's final ruling before relying on either outcome for any project involving BDA-acquired land."],
      ["K-RERA Enforcement — Orders, Recovery, Routes & Limits", "Beyond the Adarsh Nivas (₹79.84 lakh) and Casagrand Garden City (₹52.74 lakh) orders, the ₹92 Cr/₹758 Cr recovery gap, and the 8-week recovery-certificate enforcement mandate already on file, Karnataka HC has separately ruled that K-RERA orders CANNOT be executed through civil courts — reinforcing that the revenue-recovery-certificate route is the only proper enforcement channel; the HC has also struck down an entire K-RERA circular that imposed retroactive delay fees for late quarterly/annual filings, for lacking legislative backing (see Judicial Developments for detail).", "The civil-court ruling closes off an alternative enforcement route homebuyers or K-RERA might otherwise try — confirms that any K-RERA order against NBR (or any peer) will be enforced only via the recovery-certificate/revenue-recovery mechanism, now itself subject to the 8-week mandate. Separately, confirm NBR has no outstanding demand traceable to the now-struck-down retroactive-fee circular."],
      ["BBMP → GBA Transition", "GBA's revised draft voter list is now out — 88.92 lakh voters identified across 5 corporations, 369 wards, with 8,024 polling stations planned; the state government has dismissed as \"baseless\" claims that a proposed 650+-member nomination process would delay elections, though the Supreme Court's 30 June 2026 completion deadline has already passed without a confirmed poll date.", "The voter-roll milestone is concrete progress, but the missed Supreme Court deadline means election timing remains genuinely uncertain — continue treating e-Khata/property-tax record continuity as an open diligence item until elections and the database merger are both confirmed complete."],
      ["e-Khata / e-Aasthi / Kaveri 2.0", "GBA has committed to auto-approving e-Khata applications left unprocessed beyond 5 working days; a verified e-Khata reference number is now mandatory for any sale-deed registration within GBA limits; e-Aasthi began accepting SAS Property Tax ID-based e-Khata downloads on 25 Apr 2026 (~13 lakh records covered). Processing has largely stabilised, though applicants near inter-corporation boundaries should expect extra time.", "Pre-clear e-Khata reference numbers on all NBR land parcels and unsold inventory now — registration will simply be blocked without one; flag any parcel near a corporation boundary for early application given the noted boundary-area delays."],
      ["B-Khata → A-Khata Window", "Conversion charge remains cut from 5% to 2% of guidance value for the special window (15 May – 23 Aug 2026) — now roughly one month from closing.", "Fast-track conversion for any NBR-held or under-acquisition B-Khata parcels; with ~4 weeks left, this should now be treated as a near-term legal-team deadline, not a background item."],
      ["Guidance Value Revision", "The Feb 2026 revision (+6–15% Bengaluru urban, +4–8% rural) is in force; as of June 2026 a further proposed 10–15% increase remains un-gazetted, so current rates continue to apply pending formal notification.", "Consider sequencing any pending LOIs/registrations ahead of the next gazette notification — locking in acquisitions at today's guidance value could avoid a further step-up in stamp duty/registration cost."],
      ["Land Reforms Amendment 2025", "Karnataka Land Reforms & Certain Other Law (Amendment) Bill, 2025 expands Deputy Commissioners' Section 109 exemption power (0.5 ha → 4 ha, outside Bengaluru Urban/Rural); resale-after-7-years cases now need High-Power Committee (Chief Secretary) approval instead of State Government.", "For any agricultural-to-non-agricultural conversion outside core Bengaluru districts, exemption approvals may now be faster at DC level — legal team to reassess conversion timelines on outstation parcels."],
    ],
    landAcquisition: [
      ["BDA Premium Sites e-Auction 2026", "13th BDA e-auction this year (175 sites), following two prior bulk-land rounds that raised a record ~₹2,097 Cr; Deccan Herald reports the rapid pace of consecutive auctions is now \"raising policy concerns\" in some quarters, alongside the purely commercial framing (site availability, bidding mechanics) already on file.", "Continue evaluating BDA e-auction sites in NBR's active micro-markets on their commercial merits, but be aware of the emerging critical commentary around the state's land-sale pace — this could eventually translate into policy scrutiny or slower future auction cadence, which would affect NBR's assumptions about ongoing public-land supply."],
      ["Bengaluru Business Corridor (ex-PRR)", "A more contested picture than recent editions suggested: despite the 8-committee/18-location consultation structure, published village-wise compensation (₹2.50–15.60 Cr/acre) and the 40%-residential-plot offer, multiple outlets describe the project as \"stalled\"/\"slowed\" because many farmers continue to demand statutory legal-route compensation rather than accept BDA's negotiated offer. The state has now formed a NEW high-level committee — chaired by the Additional Chief Secretary (Urban Development), including the BDA Commissioner and the principal secretaries of Law, Finance and Revenue (Disaster Management) — specifically to resolve the deadlock. Package 1 (19.80 km, ₹3,348 Cr) drew bids from Shankara Narayana Constructions (\"SNC\") and Vishwa Samudra Engineering.", "Treat BBC land-acquisition completion as a genuinely contested, two-sided process — administrative momentum (committees, DCs, consent letters) is real, but so is organised farmer resistance to the negotiated-compensation framework; NBR's land-banking survey should model both an optimistic (near-term completion) and a pessimistic (continued multi-committee deadlock) timeline rather than picking one."],
      ["BDA Approval for 6,217-Acre Acquisition", "State government approved BDA's plan to acquire 6,217 acres across Bengaluru for planned development.", "Track layout notifications under this acquisition — potential TDR/compensation-linked land opportunities for NBR to evaluate on a case-by-case basis."],
      ["NICE / BMICP Land Acquisitions — Pattern of Quashings", "Beyond the 29 Jul appellate ruling (see Judicial Developments), separate Thalaghattapura-area BMICP acquisitions have also been struck down: a 1-acre parcel notified ~17 years ago was set aside (5 Dec 2025, Justice K.S. Hemalekha) for 17+ years of acquisition dormancy with no award passed, and a 43-acre-25-gunta tract was ordered denotified after the Authority itself had recommended deletion back in 2008 and the State never acted on it.", "The NICE/BMICP project now has a documented pattern of courts striking down its old, dormant land acquisitions — a much stronger cautionary signal than a single ruling; NBR should treat any parcel with a decades-old, never-executed acquisition notation in its history as high-risk for a similar dormancy-based challenge (in NBR's favour if NBR is the landowner, against NBR if it is relying on such an acquisition)."],
      ["Devanahalli KIADB Land Resistance", "A land-acquisition resistance committee is actively opposing KIADB's move to acquire farmers' land in Channarayapatna hobli, Devanahalli taluk, Bengaluru Rural — a North Bengaluru growth corridor near the airport.", "Devanahalli is an active growth corridor for several Grade A/mid-tier peers (Century, Adarsh, Sumadhura all have projects there) — ongoing resistance to KIADB acquisition in the area is a reason for extra title/consent diligence on any NBR-adjacent parcel in that taluk."],
    ],
    judicial: [
      ["NICE Ltd. v. State of Karnataka (KIADB/BMICP) — Appeal Dismissed", "Karnataka HC Division Bench (29 Jul 2026, headed by Justice D.K. Singh) dismissed NICE's appeal and upheld a single-judge order quashing the KIADB land acquisition notification for the Bangalore-Mysuru Infrastructure Corridor Project; the Bench cited serious allegations of encroachment on lakes and private land during project execution, and noted the operator had collected substantial toll revenue while affected farmers had not received adequate compensation.", "A major legal victory for landowners against a large, long-running infrastructure-linked acquisition — corrects our earlier note that NICE had merely withdrawn its case (it litigated an appeal and lost). Any NBR land near NICE/BMICP-linked corridors should be checked for a similar unresolved compensation or encroachment history."],
      ["Hosahalli Village Acquisition Quashed", "Karnataka HC quashed a 10+ acre land acquisition in Hosahalli village, Bengaluru South Taluk — proceedings dating to 1989 — citing fraud, suppression of material facts and manipulation of official records; the court relied on the G.V.K. Rao Committee's findings (which had recommended the acquisition be dropped) and set aside a single-judge order that had treated the acquisition as final.", "Reinforces that title defects and procedural fraud can unwind even a nearly 40-year-old \"final\" acquisition — NBR's diligence on any parcel with a house-building cooperative society or old acquisition history in its chain of title should go beyond face-value finality and check for underlying committee findings or record irregularities."],
      ["IMTMA v. State of Karnataka (KIAD Act)", "Karnataka HC Division Bench set aside acquisition of land for a private entity's expansion, calling it a \"colourable exercise of power\" and \"fraud on the statute\"; 12 landowner appeals allowed.", "Any NBR acquisition routed through KIADB/industrial-area mechanisms must have unambiguous public-purpose documentation — this ruling raises the evidentiary bar and litigation risk for such routes."],
      ["Banashankari VI Stage (BDA)", "Division Bench upheld BDA's acquisition for the Banashankari VI Stage township, overturning a Jan 2025 single-bench order favouring landowners; ₹50,000 costs imposed on petitioners.", "Confirms courts are increasingly upholding BDA's public-purpose township acquisitions on appeal — relevant precedent if NBR is affected by adjacent BDA layout notifications."],
      ["Bidadi Integrated Township — PIL Dismissed", "Karnataka HC dismissed a PIL challenging acquisition of 516 acres for the Bidadi Integrated Township Project, finding no legal violations and confirming the process followed the Right to Fair Compensation Act (RFCTLARR), 2013.", "A useful counterpoint to Hosahalli: acquisitions run strictly per RFCTLARR procedure and documentation are being upheld even at large scale (516 acres) — reinforces that clean process, not just age or size, is what determines litigation survivability."],
      ["P. Nagaraju v. Special LAO (NH Act Compensation)", "Karnataka HC ruled that courts cannot invoke Section 34 of the Arbitration Act to modify a land acquisition award or fix compensation under the National Highways Act — compensation disputes under that Act must follow their own statutory route, not general arbitration-law challenge.", "Relevant if any NBR parcel is ever affected by a National Highways Act acquisition (e.g., near a highway-widening project) — confirms the correct legal remedy is the Act's own mechanism, not an arbitration-style challenge, which affects how legal budgets and timelines for such disputes should be planned."],
      ["K-RERA Retroactive Delay-Fee Circular Struck Down", "Karnataka HC struck down an entire K-RERA CIRCULAR (not a single case-specific order) that sought to impose retroactive delay fees on developers/promoters for late submission of quarterly updates and annual audit statements, holding the circular lacked legislative backing and such fees cannot be levied by executive directive alone.", "This is broader than a single-developer precedent — it invalidates the underlying circular for ALL developers, including NBR; confirm no NBR project has an outstanding demand traceable to this now-struck-down circular."],
      ["RERA & Karnataka Apartment Ownership Act — Not Repugnant", "Karnataka HC held that RERA, 2016 and the Karnataka Apartment Ownership Act, 1972 are not repugnant to each other — the two statutes govern different stages of a project's life (RERA: pre-completion regulatory compliance; KAOA: post-completion apartment ownership/association/maintenance framework) and both apply.", "Confirms NBR must comply with both frameworks at their respective stages — RERA obligations don't lapse once KAOA-governed owners' associations take over, and vice versa; useful clarity for structuring post-handover documentation and association formation on NBR projects."],
    ],
    gradeA: [
      ["1", "Prestige Group", "Secured BBMP/GBA approval to build a 1.5 km private flyover connecting Bellandur and Kariyammana Agrahara, on top of the Bengaluru Airport City integrated destination (₹1,800 Cr, construction from early 2027), Aerospace Park Phase 2 (30+ acres) and Old Madras Road JV parcel already on file.", "₹20,400 Cr GDV (Q1 FY26, multi-city) + ₹1,800 Cr Airport City destination + self-funded 1.5 km flyover", "A developer building its own private flyover to solve connectivity is a notable escalation in how far Grade A players will go to unlock value on a constrained site — worth watching whether regulators extend similar approvals to other developers, which could reshape access economics for NBR's own connectivity-constrained parcels."],
      ["2", "Brigade Group", "The 20.19-acre Whitefield-Hoskote parcel (₹588.33 Cr, via Ananthay Properties) is now confirmed as a ₹5,200 Cr GDV mixed-use development spanning 4.2 msf of residential, commercial and retail space — on top of the Malur plotted development, Prestige JV on Old Madras Road, and Gunjur JDA already on file.", "8+ deals / ~121 acres FY26; Whitefield-Hoskote alone now valued at ₹5,200 Cr GDV", "The scale-up from a land-cost figure (₹588 Cr) to a full project GDV (₹5,200 Cr) shows Brigade's Whitefield-Hoskote play is a major mixed-use bet, not just a land banking move — a useful reminder to track GDV, not just acquisition cost, when sizing up peer moves."],
      ["3", "Sobha Limited", "Confirmed it plans to maintain its ~₹1,150–1,160 Cr FY26 land-spend pace into FY27 as well (~10 million sq ft/year target continuing) — a signal of sustained rather than one-off acquisition intensity — alongside the East Bengaluru pipeline (Sobha Altair, Hoskote World City) and Godrej Industries Group family office's ₹858 Cr stake divestment already on file.", "~10 msf/yr acquisition target sustained into FY27 + ₹858 Cr Godrej-family-office stake sale", "A confirmed FY27 continuation (not just FY26) makes Sobha's Bengaluru-concentrated pace a multi-year benchmark, not a single-year spike — useful for NBR's own multi-year land-bank pacing comparisons."],
      ["4", "Puravankara Limited", "Construction arm Starworth Infrastructure & Construction Ltd (SICL) has secured a second major third-party contract — ₹311.18 Cr for the Sattva AANGANE project (3.04 msf, 37-month execution) — alongside the Embassy Eden LOI (₹133.34 Cr), its 25.61 msf Bengaluru land bank, and the KVN Property Holdings JV (24.59 acres) already on file.", "25.61 msf Bengaluru land bank; ~120 acres across 6 recent land deals; SICL now holds ₹444.5 Cr+ in third-party construction contracts", "Puravankara's construction arm is now embedded in both Embassy's and Sattva's active projects — a notable diversification into contracted construction services alongside its own land-bank growth, and a reminder of how interconnected Grade A peers' execution capacity has become."],
      ["5", "Godrej Properties", "Strong Bengaluru sales momentum continues: Godrej Vanantara sold ₹2,000 Cr of homes in its launch week, Godrej Woodscapes (Whitefield-Budigere Cross) has sold 2,000+ homes worth over ₹3,150 Cr, and a 14-acre Hoskote parcel (₹1,500 Cr revenue potential) adds to its East Bengaluru pipeline — all on top of FY24-25 record sales bookings of ~₹29,444 Cr nationally and its 17-deal/443.5-acre FY26 Bengaluru-area land run.", "17 deals / 443.5 acres (FY26, national) + ₹5,150 Cr+ recent Bengaluru launch-week sales + ₹29,444 Cr FY24-25 national bookings", "Godrej's scale (national bookings) combined with hyper-local East Bengaluru execution (Whitefield-Budigere, Hoskote) makes it the benchmark to watch most closely for both pricing and absorption-rate comparisons."],
      ["6", "Embassy Group / Embassy Developments", "An 18.6-acre residential project off Whitefield on SH-648 (reported ongoing from March 2026) is planned with six high-rise towers and 1,200+ apartments; its Embassy Eden project has now engaged Puravankara's construction arm (₹133.34 Cr LOI, 34-month execution) alongside the previously reported ₹10,300 Cr FY26 North Bengaluru launch pipeline.", "₹10,300 Cr planned launches FY26 + 18.6-acre SH-648 project", "Embassy's SH-648/Whitefield project adds another large-scale East Bengaluru entrant alongside Prestige and Brigade — East Bengaluru corridor density is rising fast."],
      ["7", "Salarpuria Sattva Group", "Pre-launched Sattva Vasanta Cove, a 1-acre residential development within its flagship Sattva Urbana township in North Bengaluru, alongside ongoing premium listings (Sattva Luxuria, Sattva Magnificia).", "Multiple ongoing premium launches, incl. new Sattva Urbana infill", "Steady-state redevelopment/infill activity (including infill within its own existing townships) in mature localities — a model NBR could evaluate for infill parcels."],
      ["8", "Shriram Properties", "Targeting up to ₹4,000 Cr in pre-sales by FY27, with Pune emerging as a fast-growing market alongside Bengaluru, Chennai and Kolkata — context for its fresh ₹600 Cr Yelahanka row-housing JDA (7 acres) flagged yesterday.", "₹4,000 Cr FY27 pre-sales target (multi-city); ₹600 Cr JDA, 7 acres (Yelahanka)", "Shriram's multi-city diversification (adding Pune) alongside its Bengaluru row-housing move suggests it is pursuing growth on two fronts simultaneously — geographic expansion and product-format diversification."],
    ],
    midTier: [
      ["1", "Century Real Estate", "Planning a ₹3,000 Cr mixed-use project on a 14-acre Outer Ring Road parcel (1.7 million sq ft residential across ~750 apartments plus 0.5 million sq ft commercial/retail), backed in part by a recent ₹1,850 Cr funding round tied to a ₹14,000 Cr overall GDV pipeline; also launching Century Attur Yelahanka (458 units).", "Mixed-use ORR development, land-bank monetisation"],
      ["2", "Assetz Property Group", "New pre-launch project \"Assetz Melodies of Life\" off Hosa Road and Sarjapur Road in Choodasandra (directly adjacent to NBR's own Sarjapur-Mullur footprint); also launching Assetz Codename Sublime in Hoskote (3 BHK from ₹1.84 Cr).", "Premium apartments, redevelopment — Sarjapur-adjacent and Hoskote entrants both active"],
      ["3", "Vaishnavi Group", "No new July news; established Bengaluru residential/commercial/retail developer with a steady project pipeline.", "Residential, commercial, retail"],
      ["4", "Adarsh Developers", "Raised ₹1,600 Cr+ in funding to drive a 10 msf residential development portfolio, on top of the Sarjapur Road (Tropica Phase II), Bellandur (Rosewood) and CDP Road/Varthur launches already on file — even as K-RERA's ₹79.84 lakh delay-interest order and 60-day completion deadline on Adarsh Palm Acres Phase-2B remains a live enforcement action.", "₹1,600 Cr+ raised for a 10 msf portfolio; active Sarjapur Road launch inside NBR's core corridor, while under active K-RERA enforcement"],
      ["5", "Concorde Group", "No new July news; long-standing mid-tier developer with residential and commercial portfolio across multiple micro-markets.", "Residential, commercial"],
      ["6", "Nitesh Estates", "No new July news; boutique/premium developer with high-end residential and hospitality-linked mixed-use history.", "Premium residential, mixed-use"],
      ["7", "Mahaveer Group", "No new July news; mid-tier developer focused on affordable-to-mid-segment apartments in peripheral growth corridors.", "Mid-segment apartments"],
      ["8", "SJR Group", "No new July news; active in residential and commercial segments, incl. IT-corridor adjacent projects.", "Residential, commercial/IT-corridor"],
      ["9", "Sumadhura Group", "Going vertical with a ₹250 Cr project comprising 297 apartments across two slender 28-floor towers, on top of the 40-acre pipeline (~₹6,000 Cr potential), Sumadhura Aikya (Soukya Road) and Panathur plotted foray already on file.", "₹250 Cr twin-tower (28-floor) project; Residential, plotted, multi-city (Bengaluru + Hyderabad)"],
      ["10", "DivyaSree Developers", "No new July news; Bengaluru-origin developer with residential and commercial/IT-park legacy portfolio.", "Residential, commercial/IT parks"],
      ["11", "Total Environment", "No new July news; design-led boutique developer known for premium, low-density residential projects.", "Premium boutique residential"],
      ["12", "Casagrand", "Won Best Landscape Project (17th Realty+ Excellence Awards, Pune 2026) and Best Residential/Best Affordable Project (ET Real Estate Conclave & Awards 2026), even as K-RERA's ₹52.74 lakh GST-overcharge refund order against Casagrand Garden City remains a live enforcement action.", "Residential, multi-city (South India) — award-winning portfolio, but under active K-RERA enforcement"],
      ["13", "Birla Estates", "JV with M S Ramaiah Realty LLP for a 52-acre integrated township in Devanahalli (Birla Trimaya, ₹3,000 Cr target) has booked ~₹2,459 Cr across phases (Phase 4 alone ~₹650 Cr, 85%+ sold); its separate JV with Mitsubishi Estate Co. for a Southeast Bengaluru residential project is now confirmed at a ₹560 Cr investment.", "Premium residential, new entrant — two large JVs (Ramaiah Realty in Devanahalli; ₹560 Cr Mitsubishi Estate JV in Southeast Bengaluru)"],
      ["14", "Confident Group", "No new July news; Bengaluru-focused mid-segment residential and villa portfolio.", "Mid-segment residential, villas"],
      ["15", "Bren Corporation", "No new July news; mid-tier developer active in apartment and villa developments across peripheral corridors.", "Apartments, villas"],
      ["16", "Vaswani Group", "No new July news; established Bengaluru developer with residential and commercial project portfolio.", "Residential, commercial"],
    ],
    recommendations: [
      { head: "Land Acquisition Strategy", color: NAVY, items: [
        "Model BBC Phase 1 completion on TWO scenarios, not one — an optimistic near-term-completion case (8-committee push, consent letters, new deadlock-resolution committee) and a pessimistic continued-standoff case (farmers holding out for statutory legal-route compensation) — rather than treating recent momentum signals as a settled trajectory.",
        "Note the emerging critical commentary on BDA's rapid e-auction cadence (\"raising policy concerns\") alongside the commercial opportunity already on file — factor a chance of slower future auction cadence or added scrutiny into any multi-round BDA site-acquisition strategy.",
        "Sequence pending land LOIs and registrations ahead of the next (un-gazetted) guidance-value notification — locking in at today's rate avoids a further step-up once the proposed 10–15% hike is formalised.",
      ]},
      { head: "Legal & Compliance", color: RED_RISK, items: [
        "Treat BDA's RERA-promoter status as an open legal question, not a settled one — KREAT ruled BDA IS a promoter (must register layouts like NPKL), while a separate HC stay favours BDA on pre-2016 acquisitions; watch for the HC's final ruling before relying on either outcome for any NBR project involving BDA-acquired land.",
        "Confirm NBR has no outstanding demand traceable to the now-struck-down K-RERA retroactive-delay-fee circular — the ruling invalidated the circular itself, not just one developer's case, so it applies across the board.",
        "The NICE/BMICP ruling (29 Jul) and the Thalaghattapura BMICP quashings together show a documented pattern of Karnataka courts striking down old, dormant land acquisitions — extend legal-team review to any NBR parcel history involving a decades-old, never-executed acquisition notation.",
      ]},
      { head: "Competitive Positioning", color: GREEN_POS, items: [
        "Prestige's self-funded flyover approval and Puravankara's cross-developer construction-arm role (SICL working for both Embassy and Sattva) both show Grade A peers finding creative structures to unlock value or capacity beyond simple land acquisition — worth a standing watch item for alternative deal structures NBR could adapt.",
        "Adarsh Developers' ₹1,600 Cr+ funding raise and Sumadhura's move into 28-floor towers both show mid-tier peers scaling up ambition — reassess whether NBR's own mid-tier competitive set is still an accurate description of these peers' trajectories.",
        "Benchmark NBR's own absorption rates against Godrej Properties' national-scale momentum (₹29,444 Cr FY24-25 bookings) combined with hyper-local East Bengaluru execution — still the strongest dual-scale benchmark identified to date.",
      ]},
    ],
  },

  layer2: {
    regional: [
      ["Tamil Nadu / Chennai", "TNRERA continues to drive buyer trust (ready-to-move demand +35% post-RERA); Jan 2026 guideline-value revision raised rates along OMR, ECR and GST Road growth corridors; stamp duty 7% + registration 4% (11% all-in) charged on the higher of guideline value or sale price.", "If NBR evaluates a Chennai entry or JV, price underwriting must use the higher of guideline/market value; OMR/ECR corridors carry the steepest revised guideline values."],
      ["Telangana / Hyderabad", "The Telangana HC has moved from asking the Chief Secretary to consider replacing HYDRAA Commissioner Ranganath to directly ORDERING his removal (1 Aug 2026) over repeat contemptuous violations of court orders — the clearest resolution yet of the week-long escalation, against a backdrop of Hyderabad's 14% YoY residential registration decline (Jan 2026) and the HC's ruling that GHMC cannot acquire private land without owner consent.", "A confirmed removal order is a more concrete governance outcome than the ambiguous \"replacement request\" reported earlier this week — NBR should watch for the successor's enforcement posture, as a leadership change could shift HYDRAA's approach (for better or worse) on any Hyderabad parcel under diligence."],
      ["Andhra Pradesh / Amaravati", "A fresh government order details acquisition of 20,494 acres in total — 16,562.56 acres of patta (private) land, 104.01 acres of assigned land, and 3,828.30 acres of government land — for an International Sports City, greenfield international airport and related infrastructure, layering onto the confirmed permanent-capital status (President's April 2026 assent) and the ongoing Phase II pooling (16,666.78 acres).", "The precise private/assigned/government land breakdown shows the bulk (16,562+ acres) is coming from private patta holders, not government land — a reminder that Amaravati's build-out will keep generating fresh land-acquisition and compensation activity for years; a long-horizon watch item, not yet an actionable NBR opportunity."],
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
        "HYDRAA's commissioner has now been formally ordered removed — a genuine governance reset for Hyderabad's enforcement posture, though direction (stricter or laxer under new leadership) is not yet known; hold off recalibrating Hyderabad risk assumptions until the successor's approach becomes clear.",
        "Amaravati's precise 20,494-acre acquisition breakdown (16,562+ acres from private patta holders) is a reminder the capital's build-out will keep generating years of fresh land-acquisition and compensation activity — a long-horizon watch item, not a near-term opportunity.",
        "If a South India footprint beyond Karnataka is on NBR's roadmap, Hyderabad diligence must independently and repeatedly verify government/lake-buffer land status regardless of HYDRAA's leadership — institutional risk outlasts any single commissioner.",
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
    "Whalesbook — Birla Estates–M S Ramaiah Realty (Devanahalli township) JV coverage",
    "Swarajya Mag / Asianet Newsable — Bengaluru Business Corridor compensation and Section 27 lapse-challenge coverage",
    "The News Minute — Telangana HC Army-deployment (Malkajgiri) coverage",
    "Deccan Herald — Devanahalli KIADB land-resistance and GBA election-process coverage",
    "Vault PropTech / PropNewz — e-Khata/e-Aasthi GBA auto-approval coverage",
    "BiddingPulse — BDA e-auction bidding-process mechanics",
    "Varthabharati / NewsFirst Prime — Bengaluru Business Corridor 1,000+ farmer objections and HUDCO funding-pressure coverage",
    "Proplisto / BigInfo — Puravankara land-bank and Embassy Eden (Starworth/SICL) contract coverage",
    "Company websites (adarshdevelopers.com) — Adarsh Tropica Phase II and Adarsh Rosewood launch coverage",
    "Namma Ward — GBA revised draft voter list coverage",
    "ThePrint / TheNewsMinute — Telangana HC HYDRAA-chief custody warning and GHMC consent ruling coverage",
    "PropNewz — Karnataka guidance value gazette-notification status (14 Apr 2026 revision) coverage",
    "The Hans India / Deccan Herald — Bengaluru Business Corridor SNC L1-bid and BBCL land-acquisition-progress coverage",
    "NewsMeter — HYDRAA Commissioner Ranganath contempt-case counter-narrative and Chief Secretary replacement-request coverage",
    "BookNewProperty — Prestige BBMP/GBA private-flyover approval coverage",
    "NewsFirst Prime / Deccan Herald — Thalaghattapura BMICP land-acquisition quashing coverage (17-year dormancy, 2008 denotification-committee finding)",
    "Casemine — Karnataka land-acquisition-denotification case-law reference",
    "Business Standard / Tradebrains — Adarsh Developers funding, Sumadhura twin-tower project, Birla Estates-Mitsubishi JV figure coverage",
    "India Law / Daksha Legal — K-RERA recovery-certificate 8-week enforcement ruling (Rajesh Rao C.V. v. State of Karnataka) coverage",
    "BookNewProperty — Karnataka HC retroactive K-RERA delay-fee ruling coverage",
    "Asianet Newsable — Bengaluru Business Corridor village-wise compensation and consent-letter progress coverage",
    "Deccan Herald — BDA 40%-residential-plot farmer offer coverage",
    "The Hans India — Hyderabad residential registration decline (Jan 2026) coverage",
    "Bar and Bench — Karnataka HC RERA-applicability stay on pre-2016 BDA land-acquisition layouts coverage",
    "ILF Law / IBC Laws — Karnataka HC rulings on RERA-order civil-court execution and Planning/Regulatory Authority coordination",
    "Propvale / The Daily Jagran — Bengaluru Business Corridor DC (Land Acquisition) appointment and ₹7,000 Cr tender coverage",
    "Republic World — NBR Group \"Tier 1 developer league\" trade-press coverage",
    "Whalesbook — Sobha FY27 land-spend continuity coverage",
    "NewsMeter — Telangana HC HYDRAA-commissioner removal order (1 Aug 2026) coverage",
    "TheRealtyToday — KREAT \"BDA as promoter\" (Nadaprabhu Kempegowda Layout) ruling coverage",
    "Deccan Herald — BDA landowner-consultation committee structure (8 committees, 18 locations) and e-auction policy-concern coverage",
    "ConstructionWorld / NewsFirst Prime — Bengaluru Business Corridor \"stalled over land dispute\" coverage",
    "Deccan Herald — farmers' legal-compensation demand and new BBC deadlock-resolution committee (Addl. Chief Secretary-chaired) coverage",
    "The Hans India / ILF Law — Karnataka HC civil-court-execution ruling (Oct 2025) and retroactive-delay-fee circular strike-down (Sep 2025) coverage",
    "LiveLawBiz — RERA / Karnataka Apartment Ownership Act non-repugnancy ruling coverage",
    "PropNewz — Bengaluru Business Corridor Package 1 bidder (Shankara Narayana Constructions / Vishwa Samudra Engineering) coverage",
    "Company disclosures/press releases — Prestige, Brigade, Sobha, Puravankara, Godrej Properties, Embassy, Sattva, Shriram Properties, Sumadhura, Century Real Estate, Assetz, Casagrand, Adarsh Developers, Birla Estates",
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
