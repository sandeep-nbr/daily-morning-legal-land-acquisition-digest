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
  dateLabel: "Saturday, 29 August 2026",
  editionNo: "Vol. I — Edition 034",

  execSummary: [
    { tag: "WATCH", color: AMBER_WATCH, text: "The Karnataka government has begun formal valuation of 24 \"enemy properties\" statewide (held by the Custodian of Enemy Property for India under the Enemy Property Act, 1968), including 6 in prime central Bengaluru — Raj Bhavan Road, Vittal Mallya Road, Victoria Road, Kalasipalya and Grant Road, ~1.5 lakh sq ft, collectively worth ~₹500 Cr. A potential future central-Bengaluru disposal/auction pipeline worth tracking, though no sale timeline has been announced." },
    { tag: "INFO", color: NAVY, text: "A quiet stretch otherwise: no new Grade A/mid-tier Bengaluru land deal since Puravankara's Electronic City JDA and Brigade's Whitefield buy, no fresh Karnataka HC land ruling since the Lalitha S. GST-compensation decision, no public BDA response to the BBC Governor petition, no K-RERA leadership change, and no developments in TNRERA, HYDRAA, Amaravati or Kerala beyond what is already on file. No new NBR-specific news." },
  ],

  layer1: {
    regulatory: [
      ["K-RERA Registrations & Scope", "K-RERA has now registered 28,000+ projects and 35,000+ agents since inception; residential projects which applied for an Occupancy Certificate before 1 May 2017 remain exempt from mandatory RERA registration; N. Jayaram has been appointed a K-RERA Member for a 5-year term. The underlying dispute over BDA's RERA status now has two pulling-apart rulings: the Karnataka RERA Appellate Tribunal (KREAT, Mar 2026) classified BDA as a \"promoter\" under RERA and ordered it to register the Nadaprabhu Kempegowda Layout project, while Karnataka HC (3 Jun 2026) stayed RERA applicability to BDA layouts with pre-2016 land acquisition notifications — interim relief ran only until 6 Jul 2026, a date now passed without a confirmed follow-up order.", "Confirm registration status (or exemption basis) for every NBR project before marketing. Treat BDA's RERA-promoter status as genuinely contested rather than settled either way — the Tribunal and the HC stay point in opposite directions, and NBR should watch for the HC's final ruling before relying on either outcome for any project involving BDA-acquired land."],
      ["K-RERA Enforcement — Orders, Recovery, Routes & Limits", "Beyond the Adarsh Nivas (₹79.84 lakh) and Casagrand Garden City (₹52.74 lakh) orders, the ₹92 Cr/₹758 Cr recovery gap, and the 8-week recovery-certificate enforcement mandate already on file, Karnataka HC has separately ruled that K-RERA orders CANNOT be executed through civil courts — reinforcing that the revenue-recovery-certificate route is the only proper enforcement channel; the HC has also struck down an entire K-RERA circular that imposed retroactive delay fees for late quarterly/annual filings, for lacking legislative backing (see Judicial Developments for detail). Most recently (Jul 2026), Karnataka HC upheld a smaller order directing a Bengaluru developer to refund ~₹1.77 lakh with interest to homebuyers — a modest sum, but a further data point on courts consistently backing RERA refund enforcement.", "The civil-court ruling closes off an alternative enforcement route homebuyers or K-RERA might otherwise try — confirms that any K-RERA order against NBR (or any peer) will be enforced only via the recovery-certificate/revenue-recovery mechanism, now itself subject to the 8-week mandate. Separately, confirm NBR has no outstanding demand traceable to the now-struck-down retroactive-fee circular. The pattern across cases of all sizes (₹1.77 lakh to ₹79.84 lakh) suggests refund/interest exposure is a near-certain outcome once a K-RERA order is upheld — budget accordingly rather than assuming smaller claims will lapse."],
      ["BBMP → GBA Transition", "Two separate tracks are now clearly diverging: administrative handover (GBA formally replacing BBMP from 2 September 2026, Maheshwar Rao as incoming commissioner) remains on schedule, but civic ELECTIONS have been pushed back a third time — the Supreme Court extended the poll deadline from 31 August to 31 December 2026, citing the ongoing Special Intensive Revision (SIR) of electoral rolls (concluding only by November), with a stern warning against further delay. Karnataka BJP has filed a separate SC plea alleging the state government deliberately delayed polls fearing electoral loss.", "NBR should treat the 2 September administrative handover as the operative near-term date for khata/tax/approval filings, but now expect elected local governance (and any associated policy shifts) to remain unsettled until at least December 2026 — a materially longer governance-uncertainty window than previously assumed. The BJP's political challenge adds a layer of uncertainty to whether even the December deadline holds."],
      ["e-Khata / e-Aasthi / Kaveri 2.0", "GBA has committed to auto-approving e-Khata applications left unprocessed beyond 5 working days; a verified e-Khata reference number is now mandatory for any sale-deed registration within GBA limits; e-Aasthi began accepting SAS Property Tax ID-based e-Khata downloads on 25 Apr 2026 (~13 lakh records covered). Processing has largely stabilised, though applicants near inter-corporation boundaries should expect extra time.", "Pre-clear e-Khata reference numbers on all NBR land parcels and unsold inventory now — registration will simply be blocked without one; flag any parcel near a corporation boundary for early application given the noted boundary-area delays."],
      ["B-Khata → A-Khata Window", "The 2%-of-guidance-value conversion window (15 May – 23 Aug 2026) has now CLOSED, with no extension announced. From today (24 Aug) the standard 5%-of-guidance-value fee applies to all new B-Khata conversion filings.", "Confirm with the legal team whether every eligible NBR B-Khata filing was submitted before yesterday's close; any that were not now face the full 5% rate (a multi-lakh cost difference on a typical site). Watch for any government signal of a fresh reduced-rate window, given only a fraction of the ~7 lakh eligible properties statewide converted during this one."],
      ["Guidance Value Revision", "The Feb 2026 revision (+6–15% Bengaluru urban, +4–8% rural) is in force; as of June 2026 a further proposed 10–15% increase remains un-gazetted, so current rates continue to apply pending formal notification.", "Consider sequencing any pending LOIs/registrations ahead of the next gazette notification — locking in acquisitions at today's guidance value could avoid a further step-up in stamp duty/registration cost."],
      ["Land Reforms Amendment 2025-26", "Karnataka Land Reforms & Certain Other Law (Amendment) Bill, 2025 expands Deputy Commissioners' Section 109 exemption power (0.5 ha → 4 ha, outside Bengaluru Urban/Rural); resale-after-7-years cases now need High-Power Committee (Chief Secretary) approval instead of State Government. A further 2026 amendment has softened the penalty regime for using agricultural land for non-agricultural purposes without conversion — replacing up to 3 years' imprisonment plus a ₹10,000 fine with a flat ₹1 lakh civil fine only.", "For any agricultural-to-non-agricultural conversion outside core Bengaluru districts, exemption approvals may now be faster at DC level. The shift from criminal exposure to a fixed civil fine also lowers the personal-liability stakes of a conversion lapse for NBR's land team — but a flat ₹1 lakh fine is a low deterrent state-wide, worth watching for any knock-on effect on informal conversions by smaller landholders NBR may be acquiring from."],
    ],
    landAcquisition: [
      ["BDA Premium Sites e-Auction 2026", "A further round (69 premium sites across Anjanapura Township, Arkavathy Layout, Banashankari 6th Stage and Sir M. Visvesvaraya Layout; registration closed 21 Aug) follows the 13th e-auction (175 sites) and two prior bulk-land rounds that raised a record ~₹2,097 Cr; Deccan Herald reports the rapid pace of consecutive auctions is now \"raising policy concerns\" in some quarters, alongside the purely commercial framing (site availability, bidding mechanics) already on file.", "Continue evaluating BDA e-auction sites in NBR's active micro-markets on their commercial merits, but be aware of the emerging critical commentary around the state's land-sale pace — this could eventually translate into policy scrutiny or slower future auction cadence, which would affect NBR's assumptions about ongoing public-land supply."],
      ["\"Enemy Properties\" Valuation — Central Bengaluru", "Karnataka has begun formal valuation of 24 \"enemy properties\" statewide (held by the Custodian of Enemy Property for India under the Enemy Property Act, 1968 — left behind by people who took Pakistani/Chinese citizenship after the 1962/1965/1971 wars), including 6 in prime central Bengaluru (Raj Bhavan Road, Vittal Mallya Road, Victoria Road, Kalasipalya, Grant Road; ~1.5 lakh sq ft; ~₹500 Cr combined). Valuation involves the PWD (building value) and the Tahsildar's office (guidance value), compiled by the Deputy Commissioner (Bengaluru Urban), who doubles as Deputy Custodian. No disposal timeline has been announced.", "A distinct, one-off potential land-supply source in the highest-value part of central Bengaluru, unrelated to BDA/BBMP-linked land — worth monitoring for any future government auction or disposal notification, though at this stage it is valuation only, with no confirmed timeline for sale or a stated disposal mechanism (auction vs. negotiated sale)."],
      ["Dr K Shivaram Karanth Layout — Site Price Hike", "BDA has raised the price of sites in the 3,069-acre Dr K Shivaram Karanth Layout (North Bengaluru, ~34,000 sites eventually planned) to ₹6,000/sq ft — a ~22% hike from the ₹4,900/sq ft rate fixed in late 2023, making a standard 30x40 site cost ~₹72 lakh. Allotment will proceed in two phases (first phase ~3,000 sites) owing to ongoing legal proceedings.", "A ~22% BDA site-price hike in a major North Bengaluru layout is a useful fresh benchmark for guidance-value trajectory in that corridor — factor it into comparative land-cost assumptions for any NBR parcel evaluation near Karanth Layout, and note that a large public-land supply (34,000 sites eventually) is still working through legal proceedings before full release."],
      ["Bengaluru Business Corridor (ex-PRR)", "Phase 1 land acquisition has progressed to ~83% complete (523 of 665 acres, per BBCL) — up from the ~78% figure previously on file — even as the corridor's most direct legal challenge remains unresolved: the PRR Farmers and Site Owners Association's Governor petition (arguing the June 2007 notification lapsed under the old 1894 Act for want of an award within 5 years) and the 250+-farmer Madavara-intersection HC petition have both drawn no public BDA response yet. This layers onto the 65m/35m land-use revision and ~₹8,385/sq ft commercial-plot compensation rate already on file.", "Rising physical acquisition progress (83%) alongside an unresolved legal challenge to the underlying notification's validity is a genuine tension NBR should track carefully — continued on-the-ground progress does not resolve the dormancy-doctrine risk; model a scenario where courts could still act on the Governor/Madavara petitions even after most land is physically acquired, and re-check any BBC-adjacent parcel for reliance on the 2007 notification's continued validity."],
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
      ["Shailesh B. Charati v. Arya Gruha Pvt. Ltd. — Common Amenities Sharing", "Karnataka HC Division Bench (Justices Jayant Banerji and Tara Vitasta Ganju) dismissed an appeal by flat owners in the \"Arya Hamsa\" project who sought exclusive rights over common amenities, upholding K-RERA and KREAT orders that residents of the neighbouring \"Arya Hamsa Grande\" phase could continue sharing roads, the clubhouse and entertainment facilities, since the registered sale deed itself permitted multi-phase sharing.", "Directly relevant to any NBR project developed in phases: the specific wording of the registered sale deed — not informal expectation — determines whether later-phase residents can be given amenity access; legal team should review sale-deed amenity-sharing clauses on all multi-phase NBR projects now, before any post-handover dispute arises."],
      ["Art of Living Trust — SIT Probe on 290+ Acre Encroachment Stayed", "Karnataka HC (Justice M.G.S. Kamal) stayed a state government order constituting a Special Investigation Team to probe alleged encroachment of 290+ acres of gomala (grazing) land by Ved Vignan Maha Vidya Peeth, a trust linked to the Art of Living Foundation, holding that the mandatory joint survey under the Karnataka Land Revenue Act, 1964 was not conducted before authorities concluded encroachment had occurred.", "A further reminder that Karnataka courts require strict procedural compliance — specifically a joint survey with the landholder present — before ANY government-land encroachment finding can stand, regardless of the underlying facts; NBR should ensure any government/gomala-adjacent parcel diligence documents whether a proper joint survey was conducted, since its absence is itself grounds to unwind an adverse encroachment finding."],
      ["M. Govind Reddy v. State of Karnataka — JDA Landowners Are Co-Promoters", "Karnataka HC (11 Aug 2026, Justice Suraj Govindraj) upheld a RERA recovery-certificate auction notice issued against a landowner who had entered a JDA with a developer (Venkat Estates) then undergoing Corporate Insolvency Resolution Process, holding that landowners under a JDA are CO-PROMOTERS under RERA and therefore share enforcement exposure — the landowner's land could be auctioned to satisfy the recovery certificate despite the landowner not having caused the default.", "This is directly material to NBR's own JDA-based land acquisition model: any landowner partner in an NBR JDA is a co-promoter under RERA, meaning their land carries recovery-certificate/auction exposure tied to project-level RERA compliance — not just the developer's own conduct. Review NBR's standard JDA terms to confirm indemnity/risk-allocation clauses adequately address this landowner exposure, and flag it clearly to any prospective JDA landowner partner during negotiations."],
      ["Lalitha S. v. Deputy Commissioner — No GST on Land Acquisition Compensation", "Karnataka HC (Justice R. Natraj, WP 24040/2026 (LA-RES), order dated 13 Aug 2026) held that GST cannot be deducted from compensation paid for compulsory land acquisition, since such acquisition is not a \"supply of goods or services\" under GST law — the respondents had \"acted in excess of their power\" in deducting it. The Court ordered a refund of the wrongly deducted GST with 15% interest, with recovery to be made personally from the officer responsible for the deduction.", "A directly quantifiable, personally-enforced protection for any NBR-owned land compulsorily acquired by a government authority — GST must never be deducted from the compensation payable, and any officer who does so faces personal financial liability rather than just departmental correction. Legal team should flag this precedent immediately against any pending or past NBR compensation claim showing a GST deduction."],
      ["Sobha Ltd. v. Karnataka RERA — K-REAT Email Consent Valid for Project Modification (Appeal K-REAT 84/2025)", "Karnataka Real Estate Appellate Tribunal (7 Aug 2026) held, in an appeal brought by Sobha Ltd. as the promoter, that authenticated email communications from allottees can constitute valid written consent for project changes, satisfying the two-thirds-allottee-consent requirement under Section 14(2)(ii) of RERA — Sobha had relied on such email consent to implement project alterations, and the Tribunal upheld it.", "A useful procedural clarification for any NBR project requiring allottee consent for design or scope changes — authenticated email can be relied upon as valid written consent, without needing physical signatures from two-thirds of allottees, easing the operational burden of obtaining project-modification approval. That a Grade A peer (Sobha) was the party litigating and winning this point is a further reason to adopt the same practice with confidence."],
    ],
    gradeA: [
      ["1", "Prestige Group", "Q1 FY27 results (reported this week) show a stark reversal from the momentum already on file (private flyover approval, Airport City, Aerospace Park Phase 2, Old Madras Road JV): residential pre-sales fell ~46% YoY to ₹6,579.3 Cr against a high ₹12,126.4 Cr Q1 FY26 base, and net profit declined to ₹236 Cr from ₹290 Cr — even as revenue (+16% to ₹2,675 Cr) and collections (+6% to ₹4,802.2 Cr) both grew.", "₹6,579.3 Cr Q1 FY27 pre-sales (-46% YoY) / ₹236 Cr net profit; ₹20,400 Cr GDV (Q1 FY26, multi-city) + ₹1,800 Cr Airport City destination", "The pre-sales decline is against an unusually high prior-year base, so it may be normalisation rather than genuine weakness — but it is a clear divergence from Godrej's 22%-YoY growth quarter, worth tracking whether it is company-specific or an early sign of broader Grade A demand cooling; revenue/collections growth despite the booking decline suggests execution on the existing pipeline remains healthy."],
      ["2", "Brigade Group", "Has acquired a 2-acre parcel in Whitefield, East Bengaluru for a premium residential project (~₹475 Cr GDV). Q1 FY27 results (approved 13 Aug) show a mixed print: net profit rose 33% YoY to ₹200.41 Cr and average realisation rose 21% YoY (₹14,256/sq ft) on lower costs (total expenses -17.75% YoY), but residential pre-sales value fell 5.09% YoY (₹1,061 Cr) with volume down 22.1% YoY — closer to Prestige's pattern than Godrej/Sobha/Puravankara's growth. Balance sheet remains conservative: net debt ₹2,218 Cr (0.26x D/E), 543-acre land bank (₹4,672 Cr, ₹3,734 Cr paid). This layers onto the Kanakapura Road parcel, Whitefield-Hoskote development and ₹1,500 Cr NCD fundraise already on file.", "New 2-acre / ~₹475 Cr Whitefield land deal; ₹200.41 Cr Q1 FY27 profit (+33% YoY); pre-sales ₹1,061 Cr (-5.09% YoY), volume -22.1% YoY; 543-acre land bank (₹4,672 Cr); net D/E 0.26x", "Brigade joining Prestige in a volume decline (even as both show rising realisation/profit) turns this into a genuine 3-vs-2 Grade A earnings split, not a single-peer outlier — NBR should treat this quarter's Bengaluru pre-sales picture as genuinely mixed, not uniformly strong. The fresh Whitefield buy shows Brigade continuing to deploy capital into East Bengaluru despite the softer pre-sales print."],
      ["3", "Sobha Limited", "Q1 FY27 results (reported 21 Jul) confirm the FY27 land-spend continuity already on file with strong operating numbers: record sales of ₹3,656 Cr (+76% YoY), a shift to net cash positive (₹659 Cr), and total income up 48% YoY to ₹1,330 Cr, alongside the East Bengaluru pipeline (Sobha Altair, Hoskote World City) and Godrej Industries Group family office's ₹858 Cr stake divestment already on file.", "₹3,656 Cr Q1 FY27 sales (+76% YoY); net cash positive (₹659 Cr); ~10 msf/yr acquisition target sustained into FY27 + ₹858 Cr Godrej-family-office stake sale", "The 76%-YoY sales jump and net-cash-positive milestone make Sobha the strongest of the three Grade A peers reporting so far this earnings season (alongside Godrej +22%, against Prestige -46%) — a useful benchmark for NBR's own balance-sheet discipline alongside growth."],
      ["4", "Puravankara Limited", "Has signed a fresh JDA for a 7.83-acre parcel in South-East Bengaluru near the Electronic City corridor (Hosa Road/Hosur Road, served by Namma Metro's Yellow Line) — ₹1,100 Cr GDV, ~0.89 msf saleable residential — its 5th land transaction of FY27, taking cumulative FY27 business development to ~49.76 acres / ~₹6,300 Cr GDV / ~5.12 msf. Q1 FY27 results had already confirmed continued acquisition momentum on top of the Hennur Road JDA already on file: pre-sales rose 28% YoY to ₹1,439 Cr, collections +40% YoY to ₹1,199 Cr, average price realisation +18% YoY (₹10,589/sq ft) — alongside the SICL third-party construction contracts and 25.61 msf Bengaluru land bank already on file.", "New 7.83-acre / ₹1,100 Cr South-East Bengaluru JDA (5th FY27 deal; cumulative ~49.76 acres/~₹6,300 Cr GDV); ₹1,439 Cr Q1 FY27 pre-sales (+28% YoY); 25.61 msf Bengaluru land bank", "Puravankara remains the most consistently acquisitive Grade A peer tracked in this digest — 5 land deals in FY27 alone reinforces it as the benchmark for land-acquisition cadence; the Electronic City-corridor location is a further data point on where Grade A capital is actively flowing in South-East Bengaluru."],
      ["5", "Godrej Properties", "Has acquired a further ~3.8 acres on Sarjapur Road, South Bengaluru (~₹2,400 Cr revenue potential) directly adjacent to land it already holds — the combined 30-acre tract offers ~3 million sq ft buildable (~₹3,500 Cr total revenue potential). This is on top of the confirmed Q1 FY27 results (Apr-Jun 2026, reported 4 Aug): record national sales bookings of ₹8,651 Cr (+22% YoY, highest-ever Q1), with Bengaluru alone contributing 44% (~₹3,800 Cr) led by the Godrej Vanantara launch (₹3,237 Cr), plus Godrej Woodscapes and the 14-acre Hoskote parcel already on file. Net profit fell 42% to ₹349 Cr even as bookings and collections (+18% YoY to ₹4,348 Cr) both grew.", "New ~3.8-acre / ₹2,400 Cr Sarjapur Road deal (30-acre / ~₹3,500 Cr combined tract); ₹8,651 Cr Q1 FY27 national bookings (+22% YoY), Bengaluru = 44% share (~₹3,800 Cr); 17 deals / 443.5 acres FY26 Bengaluru-area land run", "The fresh Sarjapur Road assembly is directly inside NBR's core Sarjapur-Mullur corridor — Godrej is now the latest (after Puravankara, Assetz, Adarsh, Casagrand) Grade A/mid-tier peer actively assembling land in NBR's home turf, and its land-banking pace there deserves closer tracking than a one-off deal. The falling profit margin despite rising bookings remains worth tracking for cost/pricing-pressure signals."],
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
      ["9", "Sumadhura Group", "Has outlined plans for 6 new residential projects across Bengaluru — a ~₹6,000 Cr investment for ~4,000 apartments/plots across ~90 acres — a materially larger stated near-term pipeline than the 17-acre Whitefield-Kannamangala JDA (₹3,500 Cr GDV, announced 4 Aug) alone. This layers onto the ~₹5,000 Cr (₹50 billion) 3-year Sumadhura Infracon investment plan, ₹250 Cr twin-tower project, Sumadhura Aikya (Soukya Road) and Panathur plotted foray already on file; also showcasing 10+ Bengaluru projects at a NoBroker property expo (15-16 Aug).", "New ~₹6,000 Cr / ~90-acre / 6-project Bengaluru pipeline; ₹3,500 Cr / 17-acre Whitefield-Kannamangala JDA; Residential, plotted, multi-city (Bengaluru + Hyderabad)"],
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
        "The 24-property \"enemy properties\" valuation (6 in prime central Bengaluru) is an early-stage, one-off land-supply source distinct from BDA/BBMP land — flag it for periodic monitoring given the high-value central locations, but treat it as informational only until a disposal mechanism and timeline are actually announced.",
        "The Lalitha S. GST-on-compensation ruling is a directly actionable protection to build into NBR's standard compensation-claim checklist going forward — any government compulsory-acquisition payout to NBR should be checked line-by-line for an improper GST deduction before being accepted as final.",
        "The Governor petition's lapsed-notification argument (2007 BBC notification void for want of award within 5 years under the old 1894 Act) is the strongest legal threat to the corridor's validity flagged in this digest to date — model a genuine full-or-partial-quashing scenario, not just a negotiated-compensation outcome, for any BBC-adjacent parcel economics.",
        "GBA's 2 September administrative handover date remains the firm near-term marker for khata/tax/approval filings, but elected local governance is now pushed to a third Supreme Court deadline (31 December 2026) — plan for governance uncertainty (and any policy shifts a new council might bring) to persist through year-end, materially longer than earlier assumed.",
        "Sequence pending land LOIs and registrations ahead of the next (un-gazetted) guidance-value notification — locking in at today's rate avoids a further step-up once the proposed 10–15% hike is formalised, though it remains un-gazetted nearly four months after the April 2026 announcement.",
      ]},
      { head: "Legal & Compliance", color: RED_RISK, items: [
        "The Lalitha S. ruling confirms GST can never be deducted from land acquisition compensation, with a refund-plus-15%-interest remedy and personal recovery from the responsible officer — check NBR's own compensation-receipt history (and any pending claim) for a GST deduction and raise it immediately if found; the personal-liability angle also gives NBR's compensation team useful leverage in direct dealings with acquiring officials.",
        "The B-Khata→A-Khata 2%-conversion window closed yesterday (23 Aug) with only a fraction of the ~7 lakh eligible properties statewide converted — confirm every eligible NBR filing was submitted before close; anything missed now faces the full 5% rate with no confirmed date for a similar reduced-rate window to recur.",
        "The M. Govind Reddy ruling confirms JDA landowners are RERA co-promoters and carry recovery-certificate/auction exposure tied to the developer's compliance, not just their own conduct — review NBR's standard JDA terms for adequate indemnity/risk-allocation language, and disclose this exposure clearly to prospective landowner partners during negotiations.",
        "The Arya Hamsa ruling confirms registered sale-deed wording — not informal expectation — governs whether later-phase residents can access common amenities; legal team should review amenity-sharing clauses on all multi-phase NBR projects now, before any post-handover dispute arises.",
        "The BBC dormancy challenge uses the SAME legal doctrine (lapsed/dormant acquisition notification) that has repeatedly succeeded for landowners in NICE/BMICP, Hosahalli and the 2001/2004 BDA annulment — treat any NBR-adjacent parcel with an old, unexecuted acquisition notation in its history as carrying real challenge risk, in either direction depending on NBR's position.",
      ]},
      { head: "Competitive Positioning", color: GREEN_POS, items: [
        "Puravankara's 5th FY27 land deal (7.83 acres, South-East Bengaluru/Electronic City corridor) and Brigade's fresh 2-acre Whitefield buy show Grade A land-acquisition activity remains brisk even as this quarter's pre-sales picture is genuinely mixed — capital deployment into new land has not slowed despite softer volumes at some peers (Prestige, Brigade itself).",
        "Godrej's ~3.8-acre Sarjapur Road acquisition (30-acre combined tract) and Sumadhura's ~₹6,000 Cr / 90-acre / 6-project Bengaluru pipeline, both from the prior edition, remain the most concrete land-banking data points this month — reinforcing Sarjapur Road specifically as an increasingly contested corridor for NBR.",
        "With Brigade's results now in, this Q1 FY27 earnings season is a genuine 3-vs-2 split, not a single-peer outlier: Godrej (+22%), Sobha (+76%) and Puravankara (+28%) all grew pre-sales, while Prestige (~-46%) and Brigade (-22.1% volume) both declined even as their realisation and profit rose — correct any prior assumption that Bengaluru Grade A demand is uniformly strong this quarter; NBR's own near-term absorption-rate planning should reflect this genuinely mixed picture.",
        "Both declining peers (Prestige, Brigade) still grew profit and price realisation despite falling volumes — suggesting price discipline/premiumisation, not distress, may explain their softer pre-sales; worth distinguishing from a genuine demand-side slowdown when interpreting NBR's own sales pipeline against this benchmark.",
        "Sobha's shift to net cash positive (₹659 Cr) alongside 76%-YoY sales growth remains the standout balance-sheet result this earnings season — a useful benchmark for pairing growth with financial discipline in NBR's own land-acquisition funding strategy.",
      ]},
    ],
  },

  layer2: {
    regional: [
      ["Tamil Nadu / Chennai", "TNRERA's governance vacuum now has a named cause: ex-chairperson Shiv Das Meena resigned alongside two members, and a three-member interim panel (K. Phanindra Reddy — retired IAS, chairperson; A. Nazir Ahamed — retired district judge; Reeta Harish Thakkar — retired IAS) is running the Authority for 6 months from 14 Jul 2026. The interim panel has now made its first substantive regulatory move: a circular dated 10 Aug 2026 grants a 4-month completion-deadline extension to registered projects affected by West Asia conflict-linked supply-chain disruptions (for projects whose completion date falls on/after 28 Feb 2026), treating the situation as \"war\" for RERA force-majeure purposes per a 31 Jul 2026 MoHUA advisory — no separate application or fee required. This layers onto the Dec-2025 \"Three-Account System\" mandate and \"STAR 3.0\" registration modernisation already on file, plus the Jan 2026 guideline-value revision and 11% all-in stamp duty.", "The force-majeure extension is a useful precedent confirming the interim panel is functioning normally on substantive regulatory matters despite the leadership vacuum — reduces (without eliminating) the governance-continuity risk flagged in prior editions. The Three-Account System remains the more immediately actionable compliance mechanic for any future NBR Chennai project — factor its automated-transfer/controlled-withdrawal structure into cash-flow planning independent of the governance question."],
      ["Telangana / Hyderabad", "A 7-storey building collapse in Anjaiah Nagar, Gachibowli killed 2 workers, prompting Commissioner Ranganath to order identification and demolition of unsafe buildings citywide — notable because this building-safety drive is proceeding even while the Telangana HC's broader SOP-conditioned restraint on HYDRAA's encroachment-linked demolitions remains in force, suggesting the two enforcement tracks (structural safety vs. encroachment) are being treated as legally distinct. This layers onto the ₹72 Cr HYDRAA funding tranche, Ranganath's public land-mafia allegations (Lothukunta case), the G.O.Ms.No.218 HMDA TDR geographic-flexibility amendment, TG RERA's realtor penalties, the removal-order appeal, the Malkajgiri contempt proceeding, and five IAS/IPS contempt notices already on file.", "The building-safety drive proceeding in parallel with the demolition restraint is a useful clarification: HYDRAA's enforcement posture is not uniformly paused — a structural-safety hazard can still trigger immediate demolition action distinct from the encroachment-enforcement track courts have restrained. NBR should not assume any HYDRAA-adjacent Hyderabad asset is fully insulated from action merely because the encroachment-demolition SOP is still pending."],
      ["Andhra Pradesh / Amaravati", "Andhra Pradesh has green-lit a fresh 2,344.12-acre land-pooling round for Amaravati's Inner Ring Road and Erraballen-Namburu Railway Line (20 Aug), with 689.61 acres of government land collected in Tadikonda Mandal's Motadaka area; Pedamaddur's Phase II pooling has separately reached 70.66% (719.35 of 1,018 acres). The Supreme Court (14 Aug 2026) declined to interfere with the AP High Court's quashing of a 2021 criminal case against Chief Minister Chandrababu Naidu and former minister Ponguru Narayana over alleged Amaravati land-pooling irregularities — removing a lingering legal cloud over the scheme's leadership. This adds to the State Investment Promotion Board's ₹2,08,406 Cr investment approval (25 projects, 11 Aug) and the Seed Access Road launch already on file.", "Continued incremental land-pooling progress (Pedamaddur at 70.66%, a fresh 2,344-acre round approved) plus the Supreme Court clearing the CM's legal exposure over the scheme confirms Amaravati's capital build-out remains an active, legally de-risked, multi-front effort rather than stalled — still not an actionable NBR opportunity given no Karnataka/Bengaluru nexus, but worth tracking as a barometer of AP's execution capacity if a South India footprint beyond Karnataka is ever revisited."],
      ["Kerala", "Kerala's \"Land Reforms 2.0\" review (already on file, and already drawing controversy over whether it favours farmers or encroachers/resort operators) has now moved forward: CM V.D. Satheesan announced (17 Aug 2026) a \"second phase\" of land reforms, including timely amendments to revive plantations closed for years and a proposed law amendment to speed up land acquisition for the Vizhinjam International Seaport expansion — cutting the acquisition timeline from the current 15 months to a targeted 6 months.", "The Vizhinjam-linked acquisition-speedup amendment is a concrete signal that Kerala is willing to compress land-acquisition timelines materially (15→6 months) for priority infrastructure — a useful benchmark for how fast the state can move when it chooses to, though not yet applicable outside the port-linked context. NBR's legal team should continue tracking which direction the broader Land Reforms 2.0 expert committee leans (pro-farmer vs. encroachment/resort-benefiting) before assuming any assignment-land conversion benefit."],
    ],
    deals: [
      ["Godrej Properties / Godrej Fund Management", "Bengaluru (South India-wide franchise)", "Godrej Fund Management's purchase of a Bengaluru parcel from Puravankara illustrates the institutional-capital secondary market for land now active across Godrej's South India operations.", "Fund-backed acquisition from a peer developer"],
      ["Listed developers (South India, FY26)", "Chennai & Hyderabad", "5 land deals each in Chennai (74+ acres) and Hyderabad (~38 acres) among listed players in FY26; Coimbatore also saw listed-player activity.", "~112 acres combined (Chennai+Hyderabad)"],
      ["Casagrand", "Chennai (HQ) + Bengaluru expansion", "Continues multi-city project delivery (incl. Casagrand Moondance, Bengaluru) alongside a K-RERA GST-refund enforcement order in Bengaluru this week.", "Multi-project, multi-city — under K-RERA enforcement"],
      ["Sumadhura Group", "Bengaluru + Hyderabad", "New plotted-development vertical (~₹1,500 Cr topline target) adds to its concurrent Bengaluru/Hyderabad pipeline, reinforcing cross-city land banking by mid-tier South India players.", "New plotted vertical, ~₹1,500 Cr target"],
      ["Hyderabad-Amaravati-Chennai Bullet Train (NHSRCL)", "Telangana (180 km) + Andhra Pradesh (518 km) + Tamil Nadu (61 km)", "One of seven new HSR corridors from Union Budget 2026-27; 18 stations incl. Hyderabad, Amaravati, Guntur, Nellore, Tirupati and Chennai. DPR work underway (target Mar 2027); land acquisition, funding and construction programme not yet finalised.", "760 km, ~₹2.24 lakh Cr — pre-land-acquisition planning stage"],
    ],
    recommendations: [
      { head: "Regional Positioning", color: NAVY, items: [
        "HYDRAA's Gachibowli building-safety demolition drive proceeding in parallel with the still-pending SOP-conditioned encroachment-demolition restraint clarifies that the two enforcement tracks are legally distinct — do not assume any Hyderabad-area asset is insulated from HYDRAA action solely because the encroachment-focused restraint remains in force.",
        "The Telangana HC's blanket restraint on HYDRAA demolitions (pending a formal SOP) is the broadest operational curb flagged in this digest to date — any active or threatened HYDRAA enforcement on a Hyderabad parcel is likely paused for now, but the underlying encroachment risk is unchanged and enforcement will resume once an SOP is filed; do not treat the pause as a permanent resolution. Section 22A land-freeze designations remain a distinct, separate diligence risk.",
        "The Hyderabad-Amaravati-Chennai bullet train corridor is still at the DPR stage (target March 2027) with no land-acquisition programme yet finalised — too early for NBR to act on, but worth tracking as it could eventually generate fresh land-acquisition and compensation activity along the Telangana/AP/TN route, similar to what Amaravati's own capital build-out has produced.",
        "Amaravati's confirmed August 2028 Phase 1 completion target gives NBR a concrete multi-year horizon for the capital's build-out — useful context for any long-horizon AP market-entry discussion, though near-term land-pooling consent (Kantheru still undecided) remains the item to actually track.",
        "Kerala's \"Land Reforms 2.0\" is now shown to be genuinely contested (encroachment/resort-benefit allegations vs. the initiative's pro-farmer framing) — track which direction the expert committee leans rather than assume a straightforwardly pro-farmer outcome.",
      ]},
      { head: "Legal Watch", color: RED_RISK, items: [
        "TNRERA's 10 Aug 2026 force-majeure circular (4-month completion extension for West Asia-conflict-affected projects) shows the interim panel is actively issuing substantive regulatory relief despite the ongoing leadership vacuum — a positive functioning signal, though permanent leadership continuity remains worth tracking for any future NBR Chennai entry.",
        "The HMDA TDR geographic-flexibility amendment (G.O.Ms.No.218) is a useful data point on how South India authorities are easing monetisation of development-rights instruments given to landowners for public-project land surrenders — worth comparing against Karnataka's own TDR framework if a similar compensation option is ever offered to NBR.",
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
    "OneIndia News — Bengaluru Business Corridor Governor-petition and lapsed-1894-Act-notification coverage",
    "LandConflictWatch — PRR Farmers and Site Owners Association legal-resistance and Madavara-intersection 250+-farmer HC-petition coverage",
    "PropNewz — Karnataka guidance-value April-2026 announcement / not-yet-gazetted status coverage",
    "Law.asia — Tamil Nadu STAR 3.0 registration-department modernisation coverage",
    "Sahi.com / Whalesbook — Prestige Estates Q1 FY27 results (pre-sales, profit, revenue, collections) coverage",
    "ScanX / Sahi.com — Brigade Enterprises Q1 FY27 results-date (13 Aug 2026 AGM) coverage",
    "TheNewsMinute / Telugu360 / LiveLaw — Telangana government's HYDRAA-removal-order appeal decision and Article 311 legal-argument coverage",
    "Investing.com / Upstox / Whalesbook — Sobha Q1 FY27 results (record sales, net cash position) coverage",
    "Business Standard / Whalesbook — Puravankara Q1 FY27 results (pre-sales, collections, land deals) coverage",
    "LiveLawBiz — RERA Cases Weekly Digest (2-8 Aug 2026): Shailesh B. Charati v. Arya Gruha Pvt. Ltd. coverage",
    "Andhra Jyothy / Siasat — Telangana HC suo motu contempt proceeding (Malkajgiri Vani Cooperative Housing Society) coverage",
    "Gulte / NewsOnAir — Amaravati Seed Access Road (E3) launch coverage",
    "5paisa / ScanX — Brigade Enterprises Q1 FY27 board-meeting and NCD-fundraise coverage",
    "Business Standard — Brigade Enterprises Q1 FY27 results (profit, realisation, pre-sales, land bank) coverage",
    "The Hans India — Bengaluru Business Corridor Phase 1 land-acquisition progress (83% complete) coverage",
    "The South First / LiveLaw / Deccan Herald — Supreme Court GBA election-deadline extension to 31 December 2026 (SIR of electoral rolls) coverage",
    "The Hans India — BJP Supreme Court plea for immediate BBMP/GBA elections coverage",
    "BookNewProperty / RealtyPromoo — BDA Dr K Shivaram Karanth Layout site-price hike (₹4,900 to ₹6,000/sq ft) coverage",
    "Taxscan / ibclaw.in — Sobha Ltd. v. Karnataka RERA (K-REAT Appeal 84/2025) case-citation coverage",
    "PKP Advocates / OneCityProperty — B-Khata to A-Khata 2%-conversion-window final-day coverage",
    "Taxscan / TaxO / SAG Infotech — Lalitha S. v. Deputy Commissioner (Karnataka HC, no-GST-on-compensation) coverage",
    "Aparna Decors / Verified.RealEstate Community — TNRERA West Asia-conflict force-majeure 4-month extension coverage",
    "Hyderabad Mail / Telangana Today — HYDRAA land-monetisation and TG RERA realtor-penalty coverage",
    "RegNews — Telangana Building Rules G.O.Ms.No.218 TDR geographic-flexibility amendment coverage",
    "The News Minute — Supreme Court Amaravati land-pooling criminal-case quashing (CM Naidu) coverage",
    "Punjab Kesari / Lokmat Times / Kerala Kaumudi — Kerala CM Satheesan second-phase land reforms and Vizhinjam Port land-acquisition-law amendment coverage",
    "India Infoline / MarketScreener — Godrej Properties Sarjapur Road ~3.8-acre / ₹2,400 Cr land-deal coverage",
    "BigInfo / PropNewsTime — Sumadhura Group 6-project / ₹6,000 Cr / 90-acre Bengaluru pipeline and NoBroker property-expo coverage",
    "Telangana Today — HYDRAA Gachibowli building-collapse and unsafe-building demolition-drive coverage",
    "eauctionsindia.com / BiddingPulse — BDA 69-site premium e-auction round (registration closed 21 Aug 2026) coverage",
    "Business Standard (Capital Market News) — Puravankara South-East Bengaluru 7.83-acre JDA (Electronic City corridor) coverage",
    "PropNewsTime — Brigade Enterprises Whitefield 2-acre land-acquisition coverage",
    "Deccan Herald / ProPPulse — Karnataka \"enemy properties\" valuation (24 statewide, 6 in central Bengaluru) coverage",
    "NewsMeter — Telangana districts' request for statewide HYDRAA expansion coverage",
    "Trade Brains / Business Today / The South First — Hyderabad-Amaravati-Chennai bullet train corridor (Budget 2026-27, 760 km, 18 stations) coverage",
    "Telangana Today / NewsMeter / Hyderabad Mail — Telangana HC demolition-restraint order on HYDRAA (no SOP filed) coverage",
    "Deccan Herald / LiveLaw / Bar and Bench / The News Minute — Karnataka HC SIT-probe stay on Art of Living Foundation-linked trust (290+ acre encroachment) coverage",
    "The Hans India — Andhra Pradesh SIPB ₹2,08,406 Cr investment-approval coverage",
    "Landeed / PKP Advocates / PropNewz — B-Khata→A-Khata Bhu Guarantee window closing-date and conversion-progress coverage",
    "Deccan Chronicle / The Hans India / Siasat — Telangana ₹72 Cr HYDRAA funding-release and Ranganath Lothukunta-case defence coverage",
    "Taxscan — M. Govind Reddy v. State of Karnataka (JDA landowners as RERA co-promoters) coverage",
    "TaxGuru / LiveLaw Law Firms — RERA landowner/promoter joint liability analysis",
    "Telugu Post / PropNewsTime — Amaravati 2,344.12-acre land-pooling green signal and Pedamaddur Phase II 70.66% progress coverage",
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
