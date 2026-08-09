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
  dateLabel: "Sunday, 9 August 2026",
  editionNo: "Vol. I — Edition 019",

  execSummary: [
    { tag: "WATCH", color: AMBER_WATCH, text: "Godrej Properties has posted its strongest-ever Q1 result — ₹8,651 Cr in sales bookings for Q1 FY27 (Apr-Jun 2026), up 22% YoY — with Bengaluru contributing 44% of the national total (~₹3,800 Cr), led by the Godrej Vanantara launch alone (₹3,237 Cr). Net profit fell 42% to ₹349 Cr even as bookings and collections (+18% YoY) both grew." },
    { tag: "WATCH", color: AMBER_WATCH, text: "Telangana HC has escalated the HYDRAA fight sharply: contempt notices have now been issued to five senior IAS/IPS officers — including Special Chief Secretary Jayesh Ranjan, not just Commissioner Ranganath — while HYDRAA separately reclaimed 861 acres of government land (₹15,000 Cr value) in a fresh Sangareddy district operation." },
    { tag: "INFO", color: NAVY, text: "TNRERA's leadership vacuum is now attributable to a named resignation (ex-chairperson Shiv Das Meena); a Dec-2025 \"Three-Account System\" mandate requiring developers to segregate project funds across three bank accounts with automated transfers remains in force regardless of the interim panel." },
    { tag: "INFO", color: NAVY, text: "Kerala's \"Land Reforms 2.0\" review is drawing early controversy — critics allege draft amendments could effectively benefit illegal land occupiers/resort operators rather than farmers. No new Bengaluru-specific Grade A/mid-tier land deal, BDA/BBC development, or Karnataka HC ruling today." },
  ],

  layer1: {
    regulatory: [
      ["K-RERA Registrations & Scope", "K-RERA has now registered 28,000+ projects and 35,000+ agents since inception; residential projects which applied for an Occupancy Certificate before 1 May 2017 remain exempt from mandatory RERA registration; N. Jayaram has been appointed a K-RERA Member for a 5-year term. The underlying dispute over BDA's RERA status now has two pulling-apart rulings: the Karnataka RERA Appellate Tribunal (KREAT, Mar 2026) classified BDA as a \"promoter\" under RERA and ordered it to register the Nadaprabhu Kempegowda Layout project, while Karnataka HC (3 Jun 2026) stayed RERA applicability to BDA layouts with pre-2016 land acquisition notifications — interim relief ran only until 6 Jul 2026, a date now passed without a confirmed follow-up order.", "Confirm registration status (or exemption basis) for every NBR project before marketing. Treat BDA's RERA-promoter status as genuinely contested rather than settled either way — the Tribunal and the HC stay point in opposite directions, and NBR should watch for the HC's final ruling before relying on either outcome for any project involving BDA-acquired land."],
      ["K-RERA Enforcement — Orders, Recovery, Routes & Limits", "Beyond the Adarsh Nivas (₹79.84 lakh) and Casagrand Garden City (₹52.74 lakh) orders, the ₹92 Cr/₹758 Cr recovery gap, and the 8-week recovery-certificate enforcement mandate already on file, Karnataka HC has separately ruled that K-RERA orders CANNOT be executed through civil courts — reinforcing that the revenue-recovery-certificate route is the only proper enforcement channel; the HC has also struck down an entire K-RERA circular that imposed retroactive delay fees for late quarterly/annual filings, for lacking legislative backing (see Judicial Developments for detail). Most recently (Jul 2026), Karnataka HC upheld a smaller order directing a Bengaluru developer to refund ~₹1.77 lakh with interest to homebuyers — a modest sum, but a further data point on courts consistently backing RERA refund enforcement.", "The civil-court ruling closes off an alternative enforcement route homebuyers or K-RERA might otherwise try — confirms that any K-RERA order against NBR (or any peer) will be enforced only via the recovery-certificate/revenue-recovery mechanism, now itself subject to the 8-week mandate. Separately, confirm NBR has no outstanding demand traceable to the now-struck-down retroactive-fee circular. The pattern across cases of all sizes (₹1.77 lakh to ₹79.84 lakh) suggests refund/interest exposure is a near-certain outcome once a K-RERA order is upheld — budget accordingly rather than assuming smaller claims will lapse."],
      ["BBMP → GBA Transition", "A concrete transition date has now emerged: GBA is set to formally replace BBMP from 2 September 2026, with Maheshwar Rao named as the incoming commissioner — the first firm milestone since the Supreme Court's 30 June 2026 completion deadline passed without a confirmed poll date. This sits alongside the revised draft voter list (88.92 lakh voters, 5 corporations, 369 wards, 8,024 polling stations) and the government's dismissal of \"delay\" claims over the proposed 650+-member nomination process already on file.", "The 2 September handover date gives NBR a firm near-term marker for when civic-service, property-tax and governance responsibility formally shifts to the five-corporation structure — flag any pending NBR filings (khata, tax, approvals) that should be lodged with BBMP before the cutover rather than risk being caught mid-transition. Elections themselves remain a separate, still-unconfirmed timeline."],
      ["e-Khata / e-Aasthi / Kaveri 2.0", "GBA has committed to auto-approving e-Khata applications left unprocessed beyond 5 working days; a verified e-Khata reference number is now mandatory for any sale-deed registration within GBA limits; e-Aasthi began accepting SAS Property Tax ID-based e-Khata downloads on 25 Apr 2026 (~13 lakh records covered). Processing has largely stabilised, though applicants near inter-corporation boundaries should expect extra time.", "Pre-clear e-Khata reference numbers on all NBR land parcels and unsold inventory now — registration will simply be blocked without one; flag any parcel near a corporation boundary for early application given the noted boundary-area delays."],
      ["B-Khata → A-Khata Window", "Conversion charge remains cut from 5% to 2% of guidance value for the special window (15 May – 23 Aug 2026) — now roughly one month from closing.", "Fast-track conversion for any NBR-held or under-acquisition B-Khata parcels; with ~4 weeks left, this should now be treated as a near-term legal-team deadline, not a background item."],
      ["Guidance Value Revision", "The Feb 2026 revision (+6–15% Bengaluru urban, +4–8% rural) is in force; as of June 2026 a further proposed 10–15% increase remains un-gazetted, so current rates continue to apply pending formal notification.", "Consider sequencing any pending LOIs/registrations ahead of the next gazette notification — locking in acquisitions at today's guidance value could avoid a further step-up in stamp duty/registration cost."],
      ["Land Reforms Amendment 2025-26", "Karnataka Land Reforms & Certain Other Law (Amendment) Bill, 2025 expands Deputy Commissioners' Section 109 exemption power (0.5 ha → 4 ha, outside Bengaluru Urban/Rural); resale-after-7-years cases now need High-Power Committee (Chief Secretary) approval instead of State Government. A further 2026 amendment has softened the penalty regime for using agricultural land for non-agricultural purposes without conversion — replacing up to 3 years' imprisonment plus a ₹10,000 fine with a flat ₹1 lakh civil fine only.", "For any agricultural-to-non-agricultural conversion outside core Bengaluru districts, exemption approvals may now be faster at DC level. The shift from criminal exposure to a fixed civil fine also lowers the personal-liability stakes of a conversion lapse for NBR's land team — but a flat ₹1 lakh fine is a low deterrent state-wide, worth watching for any knock-on effect on informal conversions by smaller landholders NBR may be acquiring from."],
    ],
    landAcquisition: [
      ["BDA Premium Sites e-Auction 2026", "13th BDA e-auction this year (175 sites), following two prior bulk-land rounds that raised a record ~₹2,097 Cr; Deccan Herald reports the rapid pace of consecutive auctions is now \"raising policy concerns\" in some quarters, alongside the purely commercial framing (site availability, bidding mechanics) already on file.", "Continue evaluating BDA e-auction sites in NBR's active micro-markets on their commercial merits, but be aware of the emerging critical commentary around the state's land-sale pace — this could eventually translate into policy scrutiny or slower future auction cadence, which would affect NBR's assumptions about ongoing public-land supply."],
      ["Bengaluru Business Corridor (ex-PRR)", "BDA has revised the corridor's land-use plan: the road itself stays at 65m width (matching the Bengaluru-Mysuru highway), but a further 35m strip once earmarked for commercial acquisition is now being handed BACK to farmers rather than acquired — shrinking the total acquisition footprint versus the plan on file. A fresh 2-month public-objection window has been opened, and over 1,000 objections have already been submitted from farmers across the 74 km stretch on both legality and compensation grounds. This sits alongside the ~₹8,385/sq ft commercial-plot compensation rate and ₹2.50–15.60 Cr/acre residential range already on file.", "This is the most tangible pro-landowner concession on the corridor to date — a genuine reduction in acquisition footprint, not just a compensation-rate adjustment. Re-run any BBC-adjacent parcel economics assuming a narrower final acquisition width, but continue tracking the objection window (2 months from opening) before treating the revised plan as final."],
      ["BDA Approval for 6,217-Acre Acquisition", "State government approved BDA's plan to acquire 6,217 acres across Bengaluru for planned development.", "Track layout notifications under this acquisition — potential TDR/compensation-linked land opportunities for NBR to evaluate on a case-by-case basis."],
      ["NICE / BMICP Land Acquisitions — Pattern of Quashings", "Beyond the 29 Jul appellate ruling (see Judicial Developments), separate Thalaghattapura-area BMICP acquisitions have also been struck down: a 1-acre parcel notified ~17 years ago was set aside (5 Dec 2025, Justice K.S. Hemalekha) for 17+ years of acquisition dormancy with no award passed, and a 43-acre-25-gunta tract was ordered denotified after the Authority itself had recommended deletion back in 2008 and the State never acted on it.", "The NICE/BMICP project now has a documented pattern of courts striking down its old, dormant land acquisitions — a much stronger cautionary signal than a single ruling; NBR should treat any parcel with a decades-old, never-executed acquisition notation in its history as high-risk for a similar dormancy-based challenge (in NBR's favour if NBR is the landowner, against NBR if it is relying on such an acquisition)."],
      ["Devanahalli KIADB Land Resistance", "A land-acquisition resistance committee is actively opposing KIADB's move to acquire farmers' land in Channarayapatna hobli, Devanahalli taluk, Bengaluru Rural — a North Bengaluru growth corridor near the airport.", "Devanahalli is an active growth corridor for several Grade A/mid-tier peers (Century, Adarsh, Sumadhura all have projects there) — ongoing resistance to KIADB acquisition in the area is a reason for extra title/consent diligence on any NBR-adjacent parcel in that taluk."],
    ],
    judicial: [
      ["NICE Ltd. v. State of Karnataka (KIADB/BMICP) — Appeal Dismissed", "Karnataka HC Division Bench (29 Jul 2026, headed by Justice D.K. Singh) dismissed NICE's appeal and upheld a single-judge order quashing the KIADB land acquisition notification for the Bangalore-Mysuru Infrastructure Corridor Project; the Bench cited serious allegations of encroachment on lakes and private land during project execution, and noted the operator had collected substantial toll revenue while affected farmers had not received adequate compensation. Coverage of the judgment text has since surfaced sharper detail: the Bench remarked \"there is nothing nice about the NICE project except that the farmers have been robbed of their lands,\" and noted that of the planned 111 km expressway, NICE had built only 5 km in 25-26 years while sitting on a land bank of 20,000+ acres.", "A major legal victory for landowners against a large, long-running infrastructure-linked acquisition — corrects our earlier note that NICE had merely withdrawn its case (it litigated an appeal and lost). The scale detail (20,000+ acre land bank against 5 km of delivered expressway) is a striking illustration of dormancy-based risk — any NBR land near NICE/BMICP-linked corridors should be checked for a similar unresolved compensation or encroachment history."],
      ["Hosahalli Village Acquisition Quashed", "Karnataka HC quashed a 10+ acre land acquisition in Hosahalli village, Bengaluru South Taluk — proceedings dating to 1989 — citing fraud, suppression of material facts and manipulation of official records; the court relied on the G.V.K. Rao Committee's findings (which had recommended the acquisition be dropped) and set aside a single-judge order that had treated the acquisition as final.", "Reinforces that title defects and procedural fraud can unwind even a nearly 40-year-old \"final\" acquisition — NBR's diligence on any parcel with a house-building cooperative society or old acquisition history in its chain of title should go beyond face-value finality and check for underlying committee findings or record irregularities."],
      ["IMTMA v. State of Karnataka (KIAD Act)", "Karnataka HC Division Bench set aside acquisition of land for a private entity's expansion, calling it a \"colourable exercise of power\" and \"fraud on the statute\"; 12 landowner appeals allowed.", "Any NBR acquisition routed through KIADB/industrial-area mechanisms must have unambiguous public-purpose documentation — this ruling raises the evidentiary bar and litigation risk for such routes."],
      ["Banashankari VI Stage (BDA)", "Division Bench upheld BDA's acquisition for the Banashankari VI Stage township, overturning a Jan 2025 single-bench order favouring landowners; ₹50,000 costs imposed on petitioners.", "Confirms courts are increasingly upholding BDA's public-purpose township acquisitions on appeal — relevant precedent if NBR is affected by adjacent BDA layout notifications."],
      ["Bidadi Integrated Township — PIL Dismissed", "Karnataka HC dismissed a PIL challenging acquisition of 516 acres for the Bidadi Integrated Township Project, finding no legal violations and confirming the process followed the Right to Fair Compensation Act (RFCTLARR), 2013.", "A useful counterpoint to Hosahalli: acquisitions run strictly per RFCTLARR procedure and documentation are being upheld even at large scale (516 acres) — reinforces that clean process, not just age or size, is what determines litigation survivability."],
      ["P. Nagaraju v. Special LAO (NH Act Compensation)", "Karnataka HC ruled that courts cannot invoke Section 34 of the Arbitration Act to modify a land acquisition award or fix compensation under the National Highways Act — compensation disputes under that Act must follow their own statutory route, not general arbitration-law challenge.", "Relevant if any NBR parcel is ever affected by a National Highways Act acquisition (e.g., near a highway-widening project) — confirms the correct legal remedy is the Act's own mechanism, not an arbitration-style challenge, which affects how legal budgets and timelines for such disputes should be planned."],
      ["K-RERA Retroactive Delay-Fee Circular Struck Down", "Karnataka HC struck down an entire K-RERA CIRCULAR (not a single case-specific order) that sought to impose retroactive delay fees on developers/promoters for late submission of quarterly updates and annual audit statements, holding the circular lacked legislative backing and such fees cannot be levied by executive directive alone.", "This is broader than a single-developer precedent — it invalidates the underlying circular for ALL developers, including NBR; confirm no NBR project has an outstanding demand traceable to this now-struck-down circular."],
      ["RERA & Karnataka Apartment Ownership Act — Not Repugnant", "Karnataka HC held that RERA, 2016 and the Karnataka Apartment Ownership Act, 1972 are not repugnant to each other — the two statutes govern different stages of a project's life (RERA: pre-completion regulatory compliance; KAOA: post-completion apartment ownership/association/maintenance framework) and both apply.", "Confirms NBR must comply with both frameworks at their respective stages — RERA obligations don't lapse once KAOA-governed owners' associations take over, and vice versa; useful clarity for structuring post-handover documentation and association formation on NBR projects."],
      ["BDA Site Allotment Cancelled Over Favouritism", "Karnataka HC nullified a 2010 BDA site allotment made to the daughter of former MP D.B. Chandregowda, ruling the allotment illegal and a product of political favouritism that violated mandatory eligibility rules.", "Adds to a pattern of judicial scrutiny of BDA's land dealings (alongside the NICE/BMICP quashings and the contested RERA-promoter status) — reinforces that NBR should independently verify the allotment process and beneficiary-eligibility basis behind any BDA-sourced site it evaluates, not just its paperwork."],
      ["2001/2004 BDA Acquisition Annulled — Scheme Abandoned", "Karnataka HC annulled a BDA land acquisition (initiated 2001, notified 2004) for a North Bengaluru residential project, ruling it \"bad in law\" and finding BDA had abandoned the scheme and failed to complete acquisition within the legally mandated timeframe.", "A further instance of a Karnataka court striking down a BDA acquisition for delay/abandonment (alongside NICE/BMICP and Hosahalli) — if NBR is ever a landowner facing a stalled/dormant government acquisition, this line of cases is directly useful precedent; if relying on such an acquisition as a buyer, it's a fresh reason for caution."],
      ["KVAFSU Land Transfer PIL Dismissed", "Karnataka HC dismissed a PIL challenging the transfer of 7 acres of Karnataka Veterinary, Animal and Fisheries Sciences University land, earmarked for High Court judges' residential quarters and a new super-specialty hospital, rejecting arguments that the loss would harm the university's accreditation/funding.", "A useful counterpoint to the favouritism finding above: clearly-reasoned, public-purpose land allocations continue to survive challenge — reinforces that documented public purpose, not just process speed, is what protects an acquisition or land transfer from being unwound."],
      ["53-Years-Later Challenge Dismissed", "Karnataka HC upheld a decades-old BDA land acquisition and dismissed a petition filed by a challenger 53 years after the original acquisition.", "A useful counterweight to the dormancy-favours-landowner pattern above: extreme delay can also defeat a PETITIONER's challenge (laches), not just an acquiring authority's position — the outcome turns on the specific facts (who delayed, and why), not delay alone."],
    ],
    gradeA: [
      ["1", "Prestige Group", "Secured BBMP/GBA approval to build a 1.5 km private flyover connecting Bellandur and Kariyammana Agrahara, on top of the Bengaluru Airport City integrated destination (₹1,800 Cr, construction from early 2027), Aerospace Park Phase 2 (30+ acres) and Old Madras Road JV parcel already on file.", "₹20,400 Cr GDV (Q1 FY26, multi-city) + ₹1,800 Cr Airport City destination + self-funded 1.5 km flyover", "A developer building its own private flyover to solve connectivity is a notable escalation in how far Grade A players will go to unlock value on a constrained site — worth watching whether regulators extend similar approvals to other developers, which could reshape access economics for NBR's own connectivity-constrained parcels."],
      ["2", "Brigade Group", "Outright purchase of a 2-acre parcel on Kanakapura Road, South Bengaluru (~₹400 Cr GDV) — its third distinct Bengaluru land move on file, alongside the Whitefield-Hoskote mixed-use development (₹5,200 Cr GDV) and Malur plotted development, plus the Prestige JV on Old Madras Road and Gunjur JDA.", "8+ deals / ~123 acres FY26; Whitefield-Hoskote alone valued at ₹5,200 Cr GDV; new Kanakapura Road parcel ~₹400 Cr GDV", "Brigade is now spreading bets across South (Kanakapura Road), East (Whitefield-Hoskote) and outer (Malur) Bengaluru simultaneously — a broader geographic spread than most peers tracked, worth noting when assessing where competitive pressure will land next for NBR."],
      ["3", "Sobha Limited", "Confirmed it plans to maintain its ~₹1,150–1,160 Cr FY26 land-spend pace into FY27 as well (~10 million sq ft/year target continuing) — a signal of sustained rather than one-off acquisition intensity — alongside the East Bengaluru pipeline (Sobha Altair, Hoskote World City) and Godrej Industries Group family office's ₹858 Cr stake divestment already on file.", "~10 msf/yr acquisition target sustained into FY27 + ₹858 Cr Godrej-family-office stake sale", "A confirmed FY27 continuation (not just FY26) makes Sobha's Bengaluru-concentrated pace a multi-year benchmark, not a single-year spike — useful for NBR's own multi-year land-bank pacing comparisons."],
      ["4", "Puravankara Limited", "Signed a further JDA — ₹1,300 Cr, 4 acres on Hennur Road, North Bengaluru, 0.84 msf saleable, launch targeted within 12 months — its sixth-plus fresh Bengaluru land move on file, alongside the SICL third-party construction contracts (Sattva AANGANE ₹311.18 Cr, Embassy Eden ₹133.34 Cr), its 25.61 msf Bengaluru land bank, and the KVN Property Holdings JV (24.59 acres) already on file.", "25.61 msf Bengaluru land bank; ~124 acres across 7 recent land deals; new ₹1,300 Cr / 4-acre Hennur Road JDA; SICL holds ₹444.5 Cr+ in third-party construction contracts", "Puravankara remains the most consistently acquisitive Grade A peer tracked in this digest — a new land move roughly every few editions, on top of its construction-services diversification through SICL."],
      ["5", "Godrej Properties", "Q1 FY27 results (Apr-Jun 2026, reported 4 Aug) confirm the Bengaluru momentum already on file was no launch-week blip: national sales bookings hit a record ₹8,651 Cr (+22% YoY, its highest-ever Q1), with Bengaluru alone contributing 44% (~₹3,800 Cr) led by the Godrej Vanantara launch (₹3,237 Cr) — on top of Godrej Woodscapes (Whitefield-Budigere Cross, 2,000+ homes / ₹3,150 Cr+) and the 14-acre Hoskote parcel already on file. Net profit fell 42% to ₹349 Cr even as bookings and collections (+18% YoY to ₹4,348 Cr) both grew.", "₹8,651 Cr Q1 FY27 national bookings (+22% YoY), Bengaluru = 44% share (~₹3,800 Cr); 17 deals / 443.5 acres FY26 Bengaluru-area land run; ₹29,444 Cr FY24-25 national bookings", "A confirmed 44% Bengaluru share of a record national quarter is the strongest evidence yet that Godrej's growth is disproportionately Bengaluru-driven — reinforces it as the benchmark to watch most closely for both pricing and absorption-rate comparisons, though the falling profit margin despite rising bookings is worth tracking for cost/pricing-pressure signals."],
      ["6", "Embassy Group / Embassy Developments", "An 18.6-acre residential project off Whitefield on SH-648 (reported ongoing from March 2026) is planned with six high-rise towers and 1,200+ apartments; its Embassy Eden project has now engaged Puravankara's construction arm (₹133.34 Cr LOI, 34-month execution) alongside the previously reported ₹10,300 Cr FY26 North Bengaluru launch pipeline.", "₹10,300 Cr planned launches FY26 + 18.6-acre SH-648 project", "Embassy's SH-648/Whitefield project adds another large-scale East Bengaluru entrant alongside Prestige and Brigade — East Bengaluru corridor density is rising fast."],
      ["7", "Salarpuria Sattva Group", "Launched a 50-acre, 2,500-unit housing township in North Bengaluru targeting ₹8,600 Cr revenue — a step-change in scale beyond the 1-acre Sattva Vasanta Cove infill and other premium listings (Sattva Luxuria, Sattva Magnificia) already on file.", "₹8,600 Cr revenue target (50-acre North Bengaluru township) + multiple ongoing premium launches", "This is Sattva's largest single project flagged in this digest to date — a marked shift from the infill/redevelopment pattern previously tracked toward large-scale greenfield township development."],
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
      ["9", "Sumadhura Group", "Signed a fresh JDA for a 17-acre parcel on the Whitefield-Kannamangala corridor, East Bengaluru (announced 4 Aug) — ₹3,500 Cr GDV, 2.5+ msf saleable residential — on top of the ~₹5,000 Cr (₹50 billion) 3-year Sumadhura Infracon investment plan, ₹250 Cr twin-tower project, 40-acre pipeline (~₹6,000 Cr potential), Sumadhura Aikya (Soukya Road) and Panathur plotted foray already on file.", "New ₹3,500 Cr / 17-acre Whitefield-Kannamangala JDA; ~₹5,000 Cr 3-year investment plan (Sumadhura Infracon); Residential, plotted, multi-city (Bengaluru + Hyderabad)"],
      ["10", "DivyaSree Developers", "No new July news; Bengaluru-origin developer with residential and commercial/IT-park legacy portfolio.", "Residential, commercial/IT parks"],
      ["11", "Total Environment", "No new July news; design-led boutique developer known for premium, low-density residential projects.", "Premium boutique residential"],
      ["12", "Casagrand", "Launching a new luxury residential project directly on Sarjapur Road — the latest peer (after Puravankara, Godrej, Adarsh, Assetz) to enter NBR's core home corridor — on top of its award-winning portfolio (17th Realty+ Excellence Awards, ET Real Estate Conclave & Awards 2026) and Casagrand Moondance, even as K-RERA's ₹52.74 lakh GST-overcharge refund order against Casagrand Garden City remains a live enforcement action.", "New Sarjapur Road luxury launch; Residential, multi-city (South India) — under active K-RERA enforcement"],
      ["13", "Birla Estates", "JV with M S Ramaiah Realty LLP for a 52-acre integrated township in Devanahalli (Birla Trimaya, ₹3,000 Cr target) has booked ~₹2,459 Cr across phases (Phase 4 alone ~₹650 Cr, 85%+ sold); its separate JV with Mitsubishi Estate Co. for a Southeast Bengaluru residential project is now confirmed at a ₹560 Cr investment.", "Premium residential, new entrant — two large JVs (Ramaiah Realty in Devanahalli; ₹560 Cr Mitsubishi Estate JV in Southeast Bengaluru)"],
      ["14", "Confident Group", "No new July news; Bengaluru-focused mid-segment residential and villa portfolio.", "Mid-segment residential, villas"],
      ["15", "Bren Corporation", "No new July news; mid-tier developer active in apartment and villa developments across peripheral corridors.", "Apartments, villas"],
      ["16", "Vaswani Group", "No new July news; established Bengaluru developer with residential and commercial project portfolio.", "Residential, commercial"],
    ],
    recommendations: [
      { head: "Land Acquisition Strategy", color: NAVY, items: [
        "GBA's confirmed 2 September handover date is a firm near-term marker — flag any pending NBR khata, tax or approval filings that should be lodged with BBMP before the cutover rather than risk them being caught mid-transition.",
        "BDA's BBC land-use revision (65m road, 35m strip returned to farmers) remains the strongest pro-landowner signal on this corridor — continue re-running BBC-adjacent parcel economics on a narrower acquisition footprint, but hold off treating it as final until the 2-month objection window closes.",
        "Sequence pending land LOIs and registrations ahead of the next (un-gazetted) guidance-value notification — locking in at today's rate avoids a further step-up once the proposed 10–15% hike is formalised.",
      ]},
      { head: "Legal & Compliance", color: RED_RISK, items: [
        "With the B-Khata→A-Khata 2%-conversion window now just over two weeks from closing (23 Aug), fast-track conversion filings for any NBR-held or under-acquisition B-Khata parcels — this should be treated as a near-term legal-team deadline, not a background item.",
        "The softened Land Reforms penalty (jail term replaced by a flat ₹1 lakh civil fine for unconverted non-agricultural use) lowers personal-liability stakes for NBR's own land team, but also lowers the deterrent for informal conversions by smaller landholders NBR may be acquiring from — factor this into title/conversion-history diligence on outstation parcels.",
        "Confirm NBR has no outstanding demand traceable to the now-struck-down K-RERA retroactive-delay-fee circular — the ruling invalidated the circular itself, not just one developer's case, so it applies across the board.",
      ]},
      { head: "Competitive Positioning", color: GREEN_POS, items: [
        "Godrej's Q1 FY27 results confirm Bengaluru drove 44% of a record national quarter (~₹3,800 Cr of ₹8,651 Cr) — the clearest quantified evidence yet of how disproportionately Bengaluru weighs in a Grade A peer's national numbers; use this ratio as a fresh benchmark when framing NBR's own market-share ambitions to leadership.",
        "Godrej's 42%-YoY profit decline despite rising bookings and collections is worth tracking as a possible early margin-pressure signal across the Grade A tier — if replicated by peers, it could indicate rising land/construction costs are outpacing price realisation even in a strong-demand market.",
        "Sumadhura's second East Bengaluru land move in two weeks (₹3,500 Cr / 17-acre Whitefield-Kannamangala JDA) confirms it has stepped up from the smaller-launch pattern tracked earlier — reassess whether NBR's mid-tier competitive tiering still reflects Sumadhura's actual scale.",
      ]},
    ],
  },

  layer2: {
    regional: [
      ["Tamil Nadu / Chennai", "TNRERA's governance vacuum now has a named cause: ex-chairperson Shiv Das Meena resigned alongside two members, and a three-member interim panel (K. Phanindra Reddy — retired IAS, chairperson; A. Nazir Ahamed — retired district judge; Reeta Harish Thakkar — retired IAS) is running the Authority for 6 months from 14 Jul 2026. Separately, and unaffected by the leadership question, a Dec-2025 \"Three-Account System\" mandate requires developers to maintain three segregated project bank accounts with automated fund transfers and controlled withdrawals — layering onto the Jan 2026 guideline-value revision and 11% all-in stamp duty already on file.", "The Three-Account System is a compliance mechanic (not a governance question) that would apply to NBR immediately upon any Chennai project — factor its automated-transfer/controlled-withdrawal structure into cash-flow planning for any future TN entry, independent of the ongoing interim-panel situation."],
      ["Telangana / Hyderabad", "The Telangana HC has sharply escalated the HYDRAA fight: contempt notices were issued (8 Aug 2026) to five senior IAS/IPS officers — including Special Chief Secretary Jayesh Ranjan, not just Commissioner Ranganath — extending the court's original 63-petition basis for the removal order well beyond a single official. Separately, HYDRAA reclaimed 861 acres of government land (₹15,000 Cr value) in a fresh Sangareddy district operation, and the Section 22A land-freeze designation (already on file) continues stalling property transactions without prior notice.", "Contempt notices reaching the Special Chief Secretary signal the courts are treating this as a systemic administrative-compliance failure, not an isolated commissioner issue — this raises rather than lowers the odds of continued aggressive HYDRAA enforcement regardless of who holds the commissioner post; the Sangareddy reclamation shows the agency's operational reach extending beyond central Hyderabad."],
      ["Andhra Pradesh / Amaravati", "Andhra Pradesh has set its first firm capital-delivery target — Amaravati's Phase 1 build-out is due for completion by August 2028, per Minister Ponguru Narayana. APCRDA has begun a fresh round of Gram Sabhas across capital-region villages (since 4 Aug) to finalise the Master Plan and village-road network ahead of a government-set land-pooling completion timeline; Tadikonda and Mothadaka farmers have agreed in principle to the Land Pooling Scheme, while Kantheru farmers have sought more time. This layers onto the 20,494-acre acquisition breakdown (16,562+ acres from private patta holders) already on file.", "The August 2028 target gives NBR a concrete multi-year horizon for when Amaravati's core infrastructure should be substantially complete — useful for framing any long-horizon AP market-entry discussion, though land-pooling consent (Kantheru still undecided) remains the near-term watch item, not the completion date itself."],
      ["Kerala", "Kerala's \"Land Reforms 2.0\" review (already on file) is now drawing early political controversy: critics allege draft amendments to land-encroachment rules could effectively benefit illegal occupiers of government/assigned land and resort/commercial operators, rather than primarily protecting farmers as the initiative claims — a tension the 6-month expert committee will need to resolve before recommendations are finalised.", "The controversy is a reminder that Land Reforms 2.0's final shape is genuinely contested, not a settled pro-farmer reform — NBR's legal team should track which direction the expert committee leans before assuming any assignment-land conversion benefit, since a resort/encroachment-favouring outcome could just as easily tighten scrutiny on legitimate land transactions as a political counter-reaction."],
    ],
    deals: [
      ["Godrej Properties / Godrej Fund Management", "Bengaluru (South India-wide franchise)", "Godrej Fund Management's purchase of a Bengaluru parcel from Puravankara illustrates the institutional-capital secondary market for land now active across Godrej's South India operations.", "Fund-backed acquisition from a peer developer"],
      ["Listed developers (South India, FY26)", "Chennai & Hyderabad", "5 land deals each in Chennai (74+ acres) and Hyderabad (~38 acres) among listed players in FY26; Coimbatore also saw listed-player activity.", "~112 acres combined (Chennai+Hyderabad)"],
      ["Casagrand", "Chennai (HQ) + Bengaluru expansion", "Continues multi-city project delivery (incl. Casagrand Moondance, Bengaluru) alongside a K-RERA GST-refund enforcement order in Bengaluru this week.", "Multi-project, multi-city — under K-RERA enforcement"],
      ["Sumadhura Group", "Bengaluru + Hyderabad", "New plotted-development vertical (~₹1,500 Cr topline target) adds to its concurrent Bengaluru/Hyderabad pipeline, reinforcing cross-city land banking by mid-tier South India players.", "New plotted vertical, ~₹1,500 Cr target"],
    ],
    recommendations: [
      { head: "Regional Positioning", color: NAVY, items: [
        "Contempt notices reaching Telangana's Special Chief Secretary (not just the HYDRAA commissioner) signal courts now treat this as a systemic compliance failure — expect continued aggressive HYDRAA enforcement regardless of who holds the commissioner post; Section 22A land-freeze designations remain a distinct, separate Hyderabad diligence risk to check on any parcel.",
        "Amaravati's confirmed August 2028 Phase 1 completion target gives NBR a concrete multi-year horizon for the capital's build-out — useful context for any long-horizon AP market-entry discussion, though near-term land-pooling consent (Kantheru still undecided) remains the item to actually track.",
        "Kerala's \"Land Reforms 2.0\" is now shown to be genuinely contested (encroachment/resort-benefit allegations vs. the initiative's pro-farmer framing) — track which direction the expert committee leans rather than assume a straightforwardly pro-farmer outcome.",
      ]},
      { head: "Legal Watch", color: RED_RISK, items: [
        "TNRERA's Dec-2025 Three-Account System (segregated project bank accounts, automated transfers, controlled withdrawals) applies independent of the ongoing interim-panel situation — factor its cash-flow mechanics into any future Chennai project planning now, not just the governance-continuity risk.",
        "TNRERA is still operating under a 6-month interim panel (from 14 Jul 2026, following ex-chairperson Shiv Das Meena's resignation) — treat any TNRERA approval or registration sought during this window as subject to extra follow-up once permanent leadership is restored, if NBR ever pursues a Chennai entry or JV.",
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
    "Asianet Newsable — Bengaluru Business Corridor commercial-plot compensation rate (~₹8,385/sq ft) coverage",
    "Whalesbook — Karnataka HC BDA site-allotment favouritism cancellation and KVAFSU land-transfer PIL dismissal coverage",
    "BookNewProperty / Business Standard — Brigade Group Kanakapura Road land-parcel acquisition coverage",
    "Deccan Herald / PropNewsTime — D.B. Chandregowda BDA site-allotment nullification (13 Jun 2026) coverage",
    "Deccan Herald / TheWalkers — 2001/2004 BDA acquisition annulment and 53-years-later petition-dismissal coverage",
    "RealtyPromoo / LandConflictWatch — Casagrand Sarjapur Road launch and Banashankari VI Stage case background",
    "ConstructionWeekOnline / India Infoline — Sattva 50-acre township, Sumadhura Infracon investment plan, Godrej land coverage",
    "Company disclosures/press releases — Prestige, Brigade, Sobha, Puravankara, Godrej Properties, Embassy, Sattva, Shriram Properties, Sumadhura, Century Real Estate, Assetz, Casagrand, Adarsh Developers, Birla Estates",
    "Deccan Herald / The Hindu — Bengaluru Business Corridor land-use revision (65m road, 35m strip returned to farmers) and 2-month objection-window coverage",
    "Business Standard — Puravankara ₹1,300 Cr Hennur Road JDA coverage",
    "LiveLaw.in — Karnataka HC ₹1.77 lakh homebuyer-refund order coverage",
    "The News Minute — GBA-replaces-BBMP (2 Sep 2026) transition and incoming commissioner Maheshwar Rao coverage",
    "The420.in — NICE Ltd. v. State of Karnataka judgment-text detail (\"nothing nice about the NICE project\", 20,000+ acre land bank) coverage",
    "Construction Week India / BigInfo — Sumadhura Group Whitefield-Kannamangala 17-acre JDA (4 Aug 2026) coverage",
    "Verified.RealEstate / Community — TNRERA chairperson and member resignations, interim panel appointment coverage",
    "TheWeek.in / Siasat / Telangana Today — HYDRAA-commissioner removal-order protest and HYDRAA-land-monetisation (HMDA/TGIIC auction) coverage",
    "Lexology — Karnataka Land Reforms 2026 amendment (non-agricultural-use penalty softening) coverage",
    "Deccan Herald / TheNewsMinute — Amaravati APCRDA Gram Sabhas and land-pooling-consent progress coverage",
    "OnManorama — Kerala Land Assignment (Pattaya Land) Rules 2026 coverage",
    "TheNewsMinute — APCRDA Gram Sabhas (4 Aug 2026) and Amaravati August-2028 Phase 1 completion-target coverage",
    "Telangana Today — Section 22A property-owner concerns and HYDRAA-land-monetisation (HMDA/TGIIC) coverage",
    "TheWeek.in / ThePrint — Telangana HC HYDRAA-commissioner removal order (63 pending contempt petitions) coverage",
    "Outlook India / Business Standard — Kerala Budget 2026-27 \"Land Reforms 2.0\" initiative coverage",
    "Business Standard / Outlook Business / InvestyWise — Godrej Properties Q1 FY27 results (₹8,651 Cr bookings, Bengaluru 44% share) coverage",
    "The Hans India / Telangana Today — HYDRAA Saroornagar recruitment-drive and Sangareddy 861-acre reclamation coverage",
    "Siasat / Telangana Today — Telangana HC contempt notices to five IAS/IPS officers incl. Special Chief Secretary Jayesh Ranjan coverage",
    "Law.asia / Verified.RealEstate Community — TNRERA Three-Account System and ex-chairperson Shiv Das Meena resignation coverage",
    "Janayugom Online — Kerala land-law amendment / encroachment-benefit controversy coverage",
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
