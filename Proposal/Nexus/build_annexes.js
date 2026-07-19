const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageOrientation, ImageRun
} = require("docx");
const fs = require("fs");
const path = require("path");
const { dims } = require("./imgdim");
const IMGDIR = process.env.IMGDIR || "/home/user/JST-Nexus/Proposal/Nexus/images";

const USABLE = 9936; // DXA usable width at 0.8" margins on US Letter
const NAVY = "1F3864", GREY = "F2F2F2", HEAD = "D9E2F3", PLACE = "FFF2CC";

// ---------- helpers ----------
const P = (text, opts = {}) => new Paragraph({
  spacing: { after: opts.after ?? 120, before: opts.before ?? 0 },
  alignment: opts.align,
  children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, size: opts.size ?? 21, color: opts.color, font: "Calibri" })],
});
const runs = (children, opts = {}) => new Paragraph({ spacing: { after: opts.after ?? 120 }, alignment: opts.align, children });
const T = (t, o = {}) => new TextRun({ text: t, size: o.size ?? 21, bold: o.bold, italics: o.italics, color: o.color, font: "Calibri" });
const sub = (t) => new TextRun({ text: t, size: 21, subScript: true, font: "Calibri" });
const sup = (t) => new TextRun({ text: t, size: 21, superScript: true, font: "Calibri" });

const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 260, after: 130 },
  children: [new TextRun({ text: t, bold: true, size: 28, color: NAVY, font: "Calibri" })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 },
  children: [new TextRun({ text: t, bold: true, size: 24, color: NAVY, font: "Calibri" })] });

const noBorder = { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} };

// image placeholder box
function placeholder(cap, fname) {
  const fp = path.join(IMGDIR, fname);
  if (fs.existsSync(fp)) {
    const d = dims(fp) || { w: 1000, h: 650, type: fname.split(".").pop() };
    const maxW = 540; // px (~5.6 in)
    const w = Math.min(maxW, d.w), h = Math.round(d.h * (w / d.w));
    return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
      children: [ new ImageRun({ type: d.type === "jpg" ? "jpg" : "png", data: fs.readFileSync(fp),
        transformation: { width: w, height: h } }) ] });
  }
  return new Table({
    width: { size: USABLE, type: WidthType.DXA }, columnWidths: [USABLE],
    rows: [ new TableRow({ children: [ new TableCell({
      width: { size: USABLE, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: PLACE },
      margins: { top: 160, bottom: 160, left: 160, right: 160 },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:60}, children:[ new TextRun({ text: "[ IMAGE PLACEHOLDER ]", bold:true, size:20, color:"7F6000", font:"Calibri" }) ]}),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:60}, children:[ new TextRun({ text: cap, italics:true, size:20, color:"7F6000", font:"Calibri" }) ]}),
        new Paragraph({ alignment: AlignmentType.CENTER, children:[ new TextRun({ text: "Drop image file here: Proposal/Nexus/images/"+fname, size:18, color:"7F6000", font:"Consolas" }) ]}),
      ],
    }) ] }) ],
  });
}
const caption = (t) => new Paragraph({ alignment: AlignmentType.CENTER, spacing:{ before: 60, after: 180 },
  children: [ new TextRun({ text: t, italics: true, size: 19, color: "595959", font:"Calibri" }) ] });

// generic table builder
function table(colWidths, headerCells, dataRows, opts = {}) {
  const mkCell = (content, w, o = {}) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill } : undefined,
    margins: { top: 40, bottom: 40, left: 90, right: 90 },
    children: (Array.isArray(content) ? content : [content]).map(txt =>
      new Paragraph({ alignment: o.align ?? AlignmentType.LEFT, spacing:{after:0},
        children: [ new TextRun({ text: String(txt), bold: o.bold, size: o.size ?? 19, color: o.color, font:"Calibri" }) ] })),
  });
  const rows = [];
  rows.push(new TableRow({ tableHeader: true, children: headerCells.map((h,i)=> mkCell(h, colWidths[i], { fill: HEAD, bold:true, align: i===0?AlignmentType.LEFT:AlignmentType.CENTER })) }));
  dataRows.forEach((r, ri) => {
    rows.push(new TableRow({ children: r.map((c,i)=> {
      const isTotal = opts.totalRows && opts.totalRows.includes(ri);
      return mkCell(c, colWidths[i], { fill: isTotal ? HEAD : (ri%2? GREY: "FFFFFF"), bold: isTotal, align: i===0?AlignmentType.LEFT:AlignmentType.CENTER });
    }) }));
  });
  return new Table({ width:{size: colWidths.reduce((a,b)=>a+b,0), type: WidthType.DXA }, columnWidths: colWidths, rows });
}

// ---------- build content ----------
const kids = [];

// Title
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:40}, children:[ new TextRun({ text:"JAPAN - INDONESIA JOINT CALL FOR PROPOSALS on “BIOENERGY”", bold:true, size:24, color:NAVY, font:"Calibri" }) ]}));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:40}, children:[ new TextRun({ text:"JST - BRIN NEXUS 2026", bold:true, size:20, color:NAVY, font:"Calibri" }) ]}));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:60}, children:[ new TextRun({ text:"Supporting Annexes: Figures, Equations, Implementation Schedule, Research Roadmap, Budget Proposal, and Detailed Budget Plan (RAB)", bold:true, size:21, font:"Calibri" }) ]}));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:40}, children:[ new TextRun({ text:"Proton-Selective Fluorine-Free Composite Membranes for Microbial Electrolysis Cells: Converting Organic Waste into High-Purity Biohydrogen", italics:true, size:19, color:"595959", font:"Calibri" }) ]}));
kids.push(new Paragraph({ border:{ bottom:{ style:BorderStyle.SINGLE, size:12, color:NAVY, space:6 } }, spacing:{after:120}, children:[] }));
kids.push(P("This document collects the illustrations, governing equations, and the four detailed planning sections (implementation schedule, research roadmap, budget proposal, and detailed budget plan) that support the Joint Application Form (Nexus_Indonesia_Form2026). The tables and figures are consistent with Parts III-V of the form. Equations are reproduced from the source membrane study (J. Membr. Sci. 540 (2017) 165-173).", { italics:true, color:"595959", size:19 }));

// ---- Graphical abstract + working principle ----
kids.push(H1("1.  Concept, Working Principle, and Governing Equations"));
kids.push(H2("1.1  Graphical abstract"));
kids.push(P("The joint concept couples molecular dynamics on the Japan side with membrane fabrication and microbial-electrolysis-cell (MEC) experiments on the Indonesia side. The Japanese team simulates proton and cation transport through the SPAES coarse-grained membrane; the Indonesian team fabricates the SPAES/PIN composite and demonstrates biohydrogen production in the MEC. Predictions and measurements are exchanged in a closed loop."));
kids.push(placeholder("Figure 1. Combined graphical abstract - Japan side (molecular dynamics simulation; SPAES coarse-grained membrane) and Indonesia side (MEC experimental setup and SPAES/PIN fabrication) feeding the MEC reaction. Source of membrane concept: doi.org/10.1016/j.memsci.2017.06.048", "01_graphical_abstract.png"));
kids.push(caption("Figure 1. Graphical abstract of the joint research."));

kids.push(H2("1.2  Working principle of the MEC"));
kids.push(P("In the microbial electrolysis cell, exoelectrogenic bacteria at the anode oxidise the organic substrate to protons, electrons, and carbon dioxide. With a small applied voltage (about 1 V shown), electrons reach the cathode through the external circuit and protons migrate across the SPAES/PIN membrane, recombining at the cathode to form hydrogen. Competing cations (Na⁺, K⁺) also cross the membrane and must be rejected relative to protons. The half-reactions are:"));
kids.push(placeholder("Figure 2. MEC working principle at 1 V: anode (substrate → CO2 + H+), cathode (2H2O → H2 + OH-), reference electrodes, and transport of H+ vs Na+/K+ across the SPAES/PIN membrane.", "02_mec_reaction.png"));
kids.push(caption("Figure 2. Reaction scheme and selective proton transport across the SPAES/PIN membrane."));
// Eq 1 & 2
kids.push(runs([ T("Anode:  CH"), sub("3"), T("COOH + 2H"), sub("2"), T("O → 2CO"), sub("2"), T(" + 8H"), sup("+"), T(" + 8e"), sup("−"), T("      E° = −0.28 V (NHE)          (1)") ], { align: AlignmentType.CENTER }));
kids.push(runs([ T("Cathode:  8H"), sup("+"), T(" + 8e"), sup("−"), T(" → 4H"), sub("2"), T("      E° = −0.42 V (NHE)                       (2)") ], { align: AlignmentType.CENTER }));

kids.push(H2("1.3  MEC experimental setup"));
kids.push(P("The two-chamber MEC is operated with substrate feed, HCl dosing under pH control, and effluent and gas-measurement lines. Anode and cathode potentials, pH, and temperature are logged through a National Instruments data-acquisition (NI DAQ) system, and the electrolyte is recirculated. This instrumentation gives the controlled, monitored conditions required to benchmark the membrane."));
kids.push(placeholder("Figure 3. MEC experimental setup: reactor with membrane/anode/cathode, effluent and gas-measurement lines, substrate and HCl pumps, recirculation, pH controller, and NI DAQ logging anode/cathode potentials, pH and temperature.", "03_mec_setup.png"));
kids.push(caption("Figure 3. Instrumented two-chamber MEC test rig."));

kids.push(H2("1.4  MEC cell assembly"));
kids.push(P("The cell is a clamped stack: an anode block housing the carbon-felt anode, a membrane gasket holding the SPAES/PIN membrane, and a cathode block housing the titanium/platinum (Ti/Pt) cathode. This exploded view defines the electrode materials and the membrane's central, load-bearing position."));
kids.push(placeholder("Figure 4. Exploded view of the MEC cell: Anode Block, Anode (carbon felt), Membrane Gasket (SPAES/PIN), Cathode (Ti/Pt), and Cathode Block.", "04_cell_assembly.png"));
kids.push(caption("Figure 4. MEC cell assembly and electrode materials."));

kids.push(H2("1.5  SPAES/PIN membrane fabrication"));
kids.push(P("Polyimide nanofibres (PIN, ~200 nm) are produced by electrospinning to form a mechanical support web. A proton-conducting SPAES solution in N,N-dimethylacetamide (DMAc) is prepared and impregnated into the PIN web; the composite is formed by doctor blading, then dried at 80 °C for 6 h and acidified in 1 N H₂SO₄. The result is a composite in which the SPAES layer conducts protons while the embedded PIN carries mechanical load."));
kids.push(placeholder("Figure 5. SPAES/PIN fabrication route: PIN electrospinning → SPAES/DMAc solution → impregnation → doctor blading → drying (80 C, 6 h) and acidification (1 N H2SO4) → SPAES/PIN composite (SPAES proton-conductor layer + PIN mechanical supporter). Source: doi.org/10.1016/j.memsci.2017.06.048", "05_membrane_fabrication.png"));
kids.push(caption("Figure 5. Fabrication procedure and structure of the SPAES/PIN composite membrane."));

kids.push(H2("1.6  Membrane characterization - governing equations"));
kids.push(P("Membrane properties are quantified with the following relations from the source study, where W, A, and T are weight, area, and thickness in the wet and dry states; V and C are the consumed volume and concentration of NaOH; Em is the potential difference; F, R, and T are the Faraday constant, gas constant, and absolute temperature; and C₁, C₂ are the two HCl concentrations:"));
kids.push(runs([ T("Water uptake:   WU (%) = ( W"), sub("wet"), T(" − W"), sub("dry"), T(" ) / W"), sub("dry"), T(" × 100                         (3)") ], { align: AlignmentType.CENTER }));
kids.push(runs([ T("Area swelling:   ΔA (%) = ( A"), sub("wet"), T(" − A"), sub("dry"), T(" ) / A"), sub("dry"), T(" × 100                          (4)") ], { align: AlignmentType.CENTER }));
kids.push(runs([ T("Thickness swelling:   ΔT (%) = ( T"), sub("wet"), T(" − T"), sub("dry"), T(" ) / T"), sub("dry"), T(" × 100               (5)") ], { align: AlignmentType.CENTER }));
kids.push(runs([ T("Ion exchange capacity:   IEC (meq/g) = ( V"), sub("NaOH"), T(" × C"), sub("NaOH"), T(" ) / W"), sub("dry"), T("       (6)") ], { align: AlignmentType.CENTER }));
kids.push(runs([ T("Proton transport number:   t"), sub("+"), T(" = ½ { [ Em·F / ( R·T·ln(C"), sub("1"), T("/C"), sub("2"), T(") ) ] + 1 }     (7)") ], { align: AlignmentType.CENTER }));
kids.push(P("Baseline targets from the source study that the model must reproduce and the optimization must beat: t₊ = 0.96 (vs 0.91 for Nafion-211); 3-10-fold lower cation crossover; 2-2.5-fold lower gas permeability; tensile strength > 40 MPa; water uptake 35 ± 3 %; IEC 2.3 ± 0.3 meq/g; MEC hydrogen production 32.4 % higher with 90.3 % purity (vs 61.8 %).", { size:19, italics:true, color:"595959" }));

// ---- ANNEX A: SCHEDULE ----
kids.push(new Paragraph({ children:[], pageBreakBefore:true }));
kids.push(H1("2.  Implementation Schedule"));
kids.push(P("The 36-month common work plan (fiscal years 2027-2029) is summarised as a quarterly Gantt chart. Japanese (JP) molecular-dynamics work packages, Indonesian (ID) experimental work packages, and joint exchange (EX) activities are shown together so that the predict-verify handovers between the teams are visible. “X” marks an active quarter."));
const qcols = [4260]; for (let i=0;i<12;i++) qcols.push(473); // 4260 + 12*473 = 9936
const schedRows = [
 ["JP1  Build & benchmark GPU-HPC MD platform",[1,2]],
 ["JP2  Baseline atomistic SPAES model + first MD",[2,3,4]],
 ["JP3  EVB/MS-EVB module; DFT-MD validation",[3,4,5]],
 ["JP4  Mechanistic model: t+ & cation PMF",[5,6,7]],
 ["JP5  Continuum NP/PB fit; design screening",[6,7,8,9,10]],
 ["JP6  Design map, blind predictions, publish",[9,10,11,12]],
 ["ID1  Upgrade characterization (SAXS/TEM/PALS)",[1,2]],
 ["ID2  Baseline fabrication & measurement",[2,3,4]],
 ["ID3  First DS/IEC family synthesis",[3,4,5]],
 ["ID4  Waste-feedstock characterization",[2,3,4]],
 ["ID5  Full family characterization (validate)",[5,6,7]],
 ["ID6  Side-chain synthesis routes",[6,7,8]],
 ["ID7  MEC rig commissioning & acetate baseline",[7,8,9]],
 ["ID8  Optimum synthesis & blind validation",[9,10,11]],
 ["ID9  Real-waste MEC demonstration & durability",[10,11,12]],
 ["ID10 Data deposition & joint IP filing",[12]],
 ["EX  ECR doctoral study at Tohoku (continuous)",[1,2,3,4,5,6,7,8,9,10,11,12]],
 ["EX  Joint Workshops 1/2/3 (Sendai/Bogor/Sendai)",[4,8,12]],
 ["EX  Reciprocal visits & conference talks",[3,4,7,8,11,12]],
 ["EX  Joint publications & annual reporting",[8,12]],
];
const schedHeader = ["Work package (JP=Japan, ID=Indonesia, EX=exchange)","1","2","3","4","5","6","7","8","9","10","11","12"];
const schedData = schedRows.map(([lbl,qs]) => [lbl, ...Array.from({length:12},(_,i)=> qs.includes(i+1)?"X":"")]);
kids.push(table(qcols, schedHeader, schedData));
kids.push(caption("Table A1. Quarterly implementation schedule over 36 months. Columns 1-4 = FY2027, 5-8 = FY2028, 9-12 = FY2029. Each year closes with a bidirectional handover."));

// ---- ANNEX B: ROADMAP ----
kids.push(new Paragraph({ children:[], pageBreakBefore:true }));
kids.push(H1("3.  Research Roadmap"));
kids.push(P("The roadmap places the funded 36-month project within a longer trajectory from validated membrane design to domestic manufacture and regional deployment."));
kids.push(H2("Figure B1.  Ten-year research roadmap"));
const rmap = [
  ["Phase 1 (2027-2029) — THIS PROJECT","Validated MD structure-selectivity design rule (Gap 1) and an optimized DS/IEC membrane (Gap 2); the blind-tested optimum demonstrated in an MEC on real Indonesian waste versus Nafion-211."],
  ["Phase 2 (2030-2032)","Scale-up from coupons to continuously cast sheet; pilot MEC on palm oil mill / food-processing effluent; establish a domestic membrane fabrication route."],
  ["Phase 3 (2033-2036)","Field deployment at agro-industrial sites; extend the platform to microbial fuel cells, reverse electrodialysis, and water treatment; disseminate across ASEAN Member States."],
];
kids.push(table([2600,7336], ["Phase","Focus"], rmap));
kids.push(caption("Figure B1. Ten-year roadmap positioning model-guided, fluorine-free membrane design as a scalable route to a domestically manufacturable bioenergy technology."));
kids.push(H2("Figure B2.  Three-year research plan (funded project)"));
const rplan = [
  ["Year 1 (FY2027)","Japan builds/benchmarks the MD-HPC platform and the baseline SPAES model; Indonesia upgrades characterization (SAXS/TEM/PALS), reproduces the baseline, synthesises the first DS/IEC family, and characterises the waste feedstocks. Handover both ways."],
  ["Year 2 (FY2028)","Japan completes and first-validates the mechanistic model (t+, cation crossover from geometry) and begins DS/IEC + side-chain screening; Indonesia validates the family, establishes side-chain routes, and commissions the MEC rigs. First joint publication."],
  ["Year 3 (FY2029)","Japan finalises the design map and issues blind predictions; Indonesia synthesises the optimum, runs the blind validation, and demonstrates it in real-waste MEC operation; joint IP filing and reporting."],
];
kids.push(table([2000,7936], ["Year","Activities"], rplan));
kids.push(caption("Figure B2. Three-year plan progressing from baselines, through a validated model and design map, to a blind-tested optimum demonstrated on real Indonesian waste."));

// ---- ANNEX C: BUDGET PROPOSAL ----
kids.push(new Paragraph({ children:[], pageBreakBefore:true }));
kids.push(H1("4.  Budget Proposal"));
kids.push(P("Under the no-cross-border-funding rule, the Japan side is funded by JST in Japanese Yen (Part IV) and the Indonesia side by BRIN/LPDP in Indonesian Rupiah (Part V); the two are budgeted and reported separately. The summaries below match Parts IV-V of the form."));
kids.push(H2("Table C1.  Japan-side funding summary (1,000 JPY) - funded by JST"));
const c1 = [
 ["Facilities and equipment","20,000","6,000","4,000","30,000","-"],
 ["Consumables","3,000","3,000","3,000","9,000","-"],
 ["Travel","8,000","9,000","8,000","25,000","20,000"],
 ["Personnel","6,000","7,000","7,000","20,000","4,000"],
 ["Others","2,000","2,000","2,000","6,000","3,000"],
 ["Direct total (a)","39,000","27,000","24,000","90,000","27,000"],
 ["Capacity building (3 ECR)","11,700","11,700","11,700","35,100","-"],
];
kids.push(table([3200,1300,1300,1300,1400,1436], ["Category","FY2027","FY2028","FY2029","TOTAL (a)","Exchange (b)"], c1, { totalRows:[5,6] }));
kids.push(caption("Table C1. Exchange plan (b) = 27,000 = 30.0% of direct (a). Indirect = 30% of direct, added separately by JST. Capacity building: 3.9M JPY/ECR/year, 11.7M/ECR maximum."));
kids.push(H2("Table C2.  Indonesia-side funding summary (1,000 IDR) - funded by BRIN/LPDP"));
const c2 = [
 ["Direct - Personnel (<=25%)","240,000","240,000","240,000","720,000"],
 ["Direct - Consumables","380,000","400,000","370,000","1,150,000"],
 ["Direct - Travel","90,000","95,000","100,000","285,000"],
 ["Direct - Capital (<=10%)","95,000","60,000","40,000","195,000"],
 ["Direct - Others","145,000","155,000","200,000","500,000"],
 ["Indirect (<=5%)","50,000","50,000","50,000","150,000"],
 ["TOTAL (<= Rp 1,000,000/year)","1,000,000","1,000,000","1,000,000","3,000,000"],
];
kids.push(table([3936,1500,1500,1500,1500], ["Category","FY2027","FY2028","FY2029","TOTAL"], c2, { totalRows:[6] }));
kids.push(caption("Table C2. Checks: personnel 24.0% (<25%); capital 6.5% (<10%); indirect 5.0% (<5%); annual request Rp 1 billion; duration 36 months - all within BRIN limits."));

// ---- ANNEX D: RAB ----
kids.push(new Paragraph({ children:[], pageBreakBefore:true }));
kids.push(H1("5.  Detailed Budget Plan (RAB) - Indonesia Side"));
kids.push(P("The RAB required by BRIN itemises the Indonesia-side (Rupiah) budget only; the Japan-side budget is in Part IV (Yen) and no cross-border funding is used. MD/HPC and all software licences are therefore absent here (Japan-side costs). Amounts are in thousand Rupiah (1,000 IDR); 3-year totals."));
kids.push(runs([ T("Title: ",{bold:true}), T("Proton-Selective Fluorine-Free Composite Membranes for MECs - Converting Organic Waste into High-Purity Biohydrogen") ]));
kids.push(runs([ T("Scheme: ",{bold:true}), T("JST-BRIN NEXUS Joint Call 2026 (Bioenergy) / RIIM Kolaborasi     "), T("Indonesia PI: ",{bold:true}), T("Prof. Nasruddin (Universitas Indonesia)") ]));
kids.push(runs([ T("Co-PI / Team: ",{bold:true}), T("Dr.Eng. Obie Farobie, M.Si. (IPB University); Erdiyanto Munandar (BRIN)") ]));
kids.push(H2("Budget Summary (Recapitulation)"));
const recap = [
 ["Personnel (max 25%)","720,000"],
 ["Consumables","1,150,000"],
 ["Travel","285,000"],
 ["Capital (max 10%)","195,000"],
 ["Other Operational Costs","500,000"],
 ["Indirect / Monitoring & Evaluation (max 5%)","150,000"],
 ["TOTAL (3 years)","3,000,000"],
];
kids.push(table([7000,2936], ["Description","Cost (1,000 IDR)"], recap, { totalRows:[6] }));

kids.push(H2("Detailed Breakdown (3-year totals, 1,000 IDR)"));
function rabBlock(title, rows, subtotal) {
  kids.push(P(title, { bold:true, after:60 }));
  const data = rows.map(r => [r[0], r[1], r[2]]);
  data.push(["SUBTOTAL","", subtotal]);
  kids.push(table([3600,4600,1736], ["Item (name / series / origin)","Justification","3-yr (Rp)"], data, { totalRows:[data.length-1] }));
}
rabBlock("A. Personnel", [
 ["Research team honoraria (PI/Co-PI/researcher/students)","Team member honoraria, held below the 25% ceiling.","540,000"],
 ["Field workers (waste sampling)","Feedstock sampling campaign.","120,000"],
 ["Expert fees","Membrane/bioprocess expertise.","60,000"],
], "720,000");
rabBlock("B. Consumables (experiment)", [
 ["SPAES base resin (Kolon, South Korea) + DS/IEC synthesis reagents","Base resin + reagents for the membrane family.","250,000"],
 ["Side-chain modifiers (hydroxyl; n-BuOH)","Channel-narrowing modification routes (Gap 2).","120,000"],
 ["N,N-dimethylacetamide (DMAc)","Casting solvent for SPAES.","60,000"],
 ["Polyimide precursor + electrospinning consumables (~200 nm)","Electrospun PIN mechanical support.","120,000"],
 ["Nafion-211 (DuPont, USA)","Benchmark control membrane.","60,000"],
 ["Chemicals (H2SO4, NaCl, NaOH, HCl, phenolphthalein, std cations)","Acidification, IEC titration, crossover.","90,000"],
 ["Nitrogen (99.999%) & standard gas","Carrier gas & gas-crossover analysis.","60,000"],
 ["MEC electrode materials: carbon felt anode; Ti/Pt cathode","Anode & cathode + VFA substrate.","120,000"],
 ["SAXS measurement time (contracted)","Channel width & cluster spacing.","120,000"],
 ["TEM tomography time (contracted)","3D tortuosity & connectivity.","100,000"],
 ["PALS measurement time (contracted)","Free-volume / void size.","50,000"],
], "1,150,000");
rabBlock("C. Travel", [
 ["Domestic travel (feedstock sampling, IPB-UI, dissemination)","Coordination & national dissemination (Rp 30M/yr).","90,000"],
 ["International travel - data & joint analysis (Tohoku); <=10% cap","Joint analysis at Tohoku (Rp 60/65/70M, FY27-29).","195,000"],
], "285,000");
rabBlock("D. Capital (instrumentation; name / series / origin) - <=10%", [
 ["Data-acquisition multimeter (Keithley 2700 series, USA)","Proton transport number (t+) measurement.","60,000"],
 ["Ag/AgCl reference electrodes (Microelectrode, USA)","Reference for t+ cell (+0.195 V vs NHE).","20,000"],
 ["Gastight syringes (Hamilton SampleLock #1002, 2.5 mL, USA)","MEC headspace gas sampling.","15,000"],
 ["Electrospinning unit hardware","PIN fabrication (~200 nm).","60,000"],
 ["Two-chamber MEC reactor hardware (180 mL)","Biohydrogen MEC operation (0.5 V).","40,000"],
], "195,000");
rabBlock("E. Other Operational Costs", [
 ["Joint Workshop 2 hosting (Indonesia, FY2028)","Convening of the FY2028 joint workshop.","120,000"],
 ["Article Processing Charges (APC)","Publication in reputable Q1 international journals.","150,000"],
 ["Ethical clearance & foreign research permit processing","Ethics Committee & klirensetik permit.","40,000"],
 ["Data deposition (BRIN RIN)","Primary-data repository deposit.","30,000"],
 ["Meeting expenses","Project & coordination meetings.","60,000"],
 ["Research materials & supplies","Consumable office/research supplies.","100,000"],
], "500,000");
rabBlock("F. Indirect (Monitoring & Evaluation) - <=5%", [
 ["Monitoring & evaluation","M&E by the beneficiary institution (capped at 5%).","150,000"],
], "150,000");
kids.push(P("GRAND TOTAL (Indonesia side, 3 years) = 3,000,000 (1,000 IDR) = Rp 3,000,000,000.", { bold:true, before:80 }));
kids.push(P("Note: FE-SEM (TESCAN-MIRA3 LMU, Czech) with EDX, AFM (tapping mode), Instron 5967 tension tester (Instron Corp., USA; ASTM D882), ion chromatography (ICS5000, DIONEX, USA), and gas chromatograph (Series 580, GawMac Instrument Co., USA; Porapak Q column) are existing/shared facilities used on a service basis and are not new-purchase items.", { size:19, italics:true, color:"595959" }));
kids.push(P("Approved,                                                            Bogor, ____ July 2026", { before:200 }));
kids.push(P("Head of Institution / Dean                                Principal Investigator (Indonesia)"));
kids.push(P("Prof. ____________________                          Prof. Nasruddin, Universitas Indonesia"));

// ---------- document ----------
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1152, bottom: 1152, left: 1152, right: 1152 } } },
    children: kids,
  }],
});
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/home/user/JST-Nexus/Proposal/Nexus/Nexus_Indonesia_Form2026_Annexes.docx", buf);
  console.log("written:", buf.length, "bytes");
});
