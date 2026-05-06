const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak, PageNumber
} = require('docx');
const fs = require('fs');

// ─── helpers ─────────────────────────────────────────────────────────
const BLUE1  = "1A3A6E";
const BLUE2  = "2E5FA3";
const TEAL   = "1A6E6E";
const GOLD   = "7B4F00";
const GRAY   = "EEF3FB";
const WHITE  = "FFFFFF";
const DRED   = "8B0000";
const ORANGE = "7A3700";

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, bold: true, size: 36, font: "Arial", color: BLUE1 })] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 28, font: "Arial", color: BLUE2 })] });
}
function h3(text) {
  return new Paragraph({ spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, bold: true, size: 24, font: "Arial", color: TEAL })] });
}
function bullet(text, bold=false) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text, bold, size: 22, font: "Arial" })] });
}
function subbullet(text) {
  return new Paragraph({ numbering: { reference: "subbullets", level: 0 }, spacing: { after: 40 },
    children: [new TextRun({ text, size: 21, font: "Arial" })] });
}
function important(text) {
  return new Paragraph({ spacing: { after: 80 },
    children: [new TextRun({ text: "★ " + text, bold: true, size: 22, font: "Arial", color: DRED })] });
}
function keyterm(label, definition) {
  return new Paragraph({ spacing: { after: 80 },
    children: [
      new TextRun({ text: label + ": ", bold: true, size: 22, font: "Arial", color: GOLD }),
      new TextRun({ text: definition, size: 22, font: "Arial" })
    ] });
}
function boxPara(text, fill=GRAY) {
  return new Paragraph({ spacing: { before: 60, after: 60 }, shading: { fill, type: ShadingType.CLEAR },
    indent: { left: 360, right: 360 },
    children: [new TextRun({ text, size: 21, font: "Arial", italics: true })] });
}
function divider() {
  return new Paragraph({ spacing: { after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "BBBBBB", space: 1 } },
    children: [] });
}
function pb() { return new Paragraph({ children: [new PageBreak()] }); }
function bold(t) { return new TextRun({ text: t, bold: true, size: 22, font: "Arial" }); }
function norm(t) { return new TextRun({ text: t, size: 22, font: "Arial" }); }
function para(children, opts={}) {
  return new Paragraph({ spacing: { after: 80 }, ...opts,
    children: Array.isArray(children) ? children : [new TextRun({ text: children, size: 22, font: "Arial" })] });
}

function makeTable(headers, rows, colWidths) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: colWidths.reduce((a,b)=>a+b,0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h,i) =>
        new TableCell({ borders, width:{size:colWidths[i],type:WidthType.DXA},
          shading:{fill:BLUE1,type:ShadingType.CLEAR}, margins:{top:80,bottom:80,left:120,right:120},
          children:[new Paragraph({children:[new TextRun({text:h,bold:true,size:20,font:"Arial",color:WHITE})]})] }))
      }),
      ...rows.map((row,ri) => new TableRow({ children: row.map((cell,ci) =>
        new TableCell({ borders, width:{size:colWidths[ci],type:WidthType.DXA},
          shading:{fill:ri%2===0?"EEF3FB":WHITE,type:ShadingType.CLEAR}, margins:{top:60,bottom:60,left:120,right:120},
          children:[new Paragraph({children:[new TextRun({text:cell,size:20,font:"Arial"})]})] }))
      }))
    ]
  });
}

// ─────────────────────────────────────────────────────────────────────
const children = [

// ══════════════════════════════════════════════════════
// COVER PAGE
// ══════════════════════════════════════════════════════
new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 2880, after: 240 },
  children: [new TextRun({ text: "ENZYMES", bold: true, size: 72, font: "Arial", color: BLUE1 })] }),
new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new TextRun({ text: "COMPLETE EXAM NOTES — ALL SLIDES COVERED", size: 32, font: "Arial", color: BLUE2 })] }),
new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 480 },
  children: [new TextRun({ text: "Introduction · Nomenclature · Cofactors · Mechanisms · Kinetics · Inhibition · Allosteric Regulation · Isozymes", size: 22, font: "Arial", color: "555555" })] }),
pb(),

// ══════════════════════════════════════════════════════
// CH1: INTRODUCTION
// ══════════════════════════════════════════════════════
h1("CHAPTER 1 — INTRODUCTION TO ENZYMES"),
divider(),

h2("1.1 What are Enzymes?"),
bullet("Enzymes are biological protein catalysts."),
bullet("They speed up the rate of biological reactions."),
bullet("Reactions that could take hundreds of years to complete in the 'real world' occur in seconds."),

h2("1.2 Example — Carbonic Anhydrase"),
bullet("Carbonic anhydrase transfers CO₂ from the tissues into the blood and then to the alveolar air (air sacs in the lungs)."),
bullet("Carbonic anhydrase is one of the fastest enzymes."),
bullet("Each enzyme molecule can hydrate 10⁶ molecules of CO₂ per second."),
bullet("This catalyzed reaction is 10⁷ times as fast as the uncatalyzed one."),

h2("1.3 Energy of Chemical Reactions"),
keyterm("Activation Energy", "The minimum amount of energy needed for reactants to form products in a chemical reaction."),
keyterm("Exothermic Reaction", "Releases heat energy. Energy of the product is LOWER than energy of the reactants."),
keyterm("Endothermic Reaction", "Absorbs heat energy. Energy of the product is HIGHER than energy of the reactants."),
para([bold("Enzymatic reactions: "), norm("Enzymes speed chemical reactions by LESSENING the activation energy. Enzymes do NOT increase the amount of product formation and do NOT get used up in reactions.")]),

h2("1.4 Effect of Temperature on Enzyme Activity (PPT Slide 11 Graph)"),
important("Graph: Effect of Temperature — ASYMMETRIC BELL-SHAPED CURVE"),
bullet("X-axis: Temperature (°C), 0–90°C. Y-axis: % Activity (0–100%)."),
bullet("Activity rises steeply with increasing temperature up to about 40°C (optimum)."),
bullet("After the optimum, the curve drops sharply as enzyme undergoes DENATURATION."),
bullet("The curve is asymmetric — the rise is gradual, the fall is steep and narrow."),

h2("1.5 Effect of pH on Enzyme Activity (PPT Slide 11 Graph)"),
important("Graph: Effect of pH — shows TWO ENZYMES (Pepsin and Trypsin)"),
bullet("X-axis: pH (0–12). Y-axis: Relative Activity (0–120%)."),
bullet("PEPSIN (orange curve): Optimal pH ≈ 2 (highly acidic). Active in the stomach."),
bullet("TRYPSIN (brown curve): Optimal pH ≈ 8 (slightly alkaline). Active in the small intestine."),
bullet("Bell-shaped curves for both — activity drops sharply on either side of the optimum."),
important("Key point from graph: Different enzymes have different optimal pH values."),

h2("1.6 Specificity of Enzymes"),
bullet("Enzymes are highly specific both in the reactions they catalyze and in their choice of reactants."),
bullet("Trypsin — catalyzes the splitting of peptide bonds only on the carboxyl side of lysine and arginine residues."),
bullet("Thrombin — catalyzes the hydrolysis of Arg–Gly bonds in particular peptide sequences only. Even more specific than trypsin."),

h2("1.7 Active Site"),
keyterm("Substrate", "The molecule that is bound in the active site and acted upon by the enzyme."),
bullet("Enzyme-catalyzed reactions take place within the confines of a pocket on the enzyme called the active site."),
bullet("The surface of the active site is lined with amino acid residues that bind the substrate and catalyze its chemical transformation."),
h3("Active Site Residue Properties:"),
bullet("Charge"), bullet("pKa"), bullet("Hydrophobicity"), bullet("Flexibility"), bullet("Reactivity"), bullet("Catalytic activity"),
keyterm("Turnover Number", "Maximum number of moles of substrate that can be converted to product per mole of catalytic site per second."),
divider(),

// ══════════════════════════════════════════════════════
// CH2: NOMENCLATURE
// ══════════════════════════════════════════════════════
h1("CHAPTER 2 — ENZYME NOMENCLATURE & CLASSIFICATION"),
divider(),

h2("2.1 Common Names vs Systematic Names"),
bullet("Many enzymes have common names that provide little information (e.g., trypsin — a proteolytic enzyme secreted by the pancreas)."),
bullet("Other enzymes are named for their substrates and reactions, with the suffix '-ase' added."),
subbullet("ATPase — enzyme that breaks down ATP."),
subbullet("ATP synthase — enzyme that synthesizes ATP."),
bullet("Sometimes the same enzyme has two or more names (e.g., Glucokinase and Hexokinase in glycolysis)."),
bullet("International Union of Biochemistry established an Enzyme Commission (EC) to develop a nomenclature for enzymes."),

h2("2.2 EC Classification — Seven Major Classes"),
important("Reactions were divided into 7 major groups (1–7). Originally 6; class 7 (Translocases) was recently added."),

makeTable(
  ["EC No.", "Class Name", "Reaction Type", "PPT Slide Examples"],
  [
    ["EC 1","Oxidoreductases","Oxidation reactions — transfer of electrons from one molecule to another","Lipoxidases, Dehydrogenases, Glucose oxidase, Reductases"],
    ["EC 2","Transferases","Catalyze the transfer of groups of atoms (e.g., methyl, phosphate, amino) from one molecule to another","Aminotransferase, Transaminase, Transaldolases, Methyltransferases"],
    ["EC 3","Hydrolases","Hydrolysis reactions — cleavage of substrates by water (C–O, C–N, C–C bonds)","Lactase, Proteases, Trypsin, Glycosylases, Peptidases, Nucleases"],
    ["EC 4","Lyases","Catalyze cleavage of C–C, C–O, C–N bonds by elimination, or addition of groups to double bonds","Pectate lyases, Decarboxylase, Hydratases, Aldolases"],
    ["EC 5","Isomerases","Catalyze geometric or structural changes in one molecule (intramolecular)","Topoisomerase, Glucose isomerase, Racemases, Epimerases"],
    ["EC 6","Ligases","Catalyze the junction/joining of two molecules to form a new bond (uses ATP)","Glutathione synthase, Aminoacyl tRNA synthetase, Synthases, Carboxylases"],
    ["EC 7","Translocases","Catalyze the movement of ions or molecules across membranes or their separation within membranes","Ubiquinone reductase, ATP synthase, Ascorbate ferrireductase, ABC-type transporters"],
  ],
  [720, 1440, 2700, 2700]
),
para(""),

h2("2.3 EC Number System (Hexokinase Example)"),
bullet("Each enzyme is assigned a four-digit number preceded by the letters EC."),
bullet("Systematic name: ATP:glucose phosphotransferase — catalyzes the transfer of a phosphoryl group from ATP to glucose."),
bullet("EC number: 2.7.1.1"),
subbullet("First number (2) — Class: Transferase"),
subbullet("Second number (7) — Subclass: Phosphotransferase"),
subbullet("Third number (1) — Phosphotransferase with a hydroxyl group as acceptor"),
subbullet("Fourth number (1) — D-glucose as the phosphoryl group acceptor"),
divider(),

// ══════════════════════════════════════════════════════
// CH3: COFACTORS
// ══════════════════════════════════════════════════════
h1("CHAPTER 3 — COFACTORS & COENZYMES"),
divider(),

h2("3.1 Definitions"),
keyterm("Cofactor", "Small molecule required for catalytic activity of many enzymes."),
keyterm("Apoenzyme", "Enzyme WITHOUT its cofactor — inactive."),
keyterm("Holoenzyme", "Complete, catalytically active enzyme."),
important("FORMULA: Apoenzyme + Cofactor = Holoenzyme"),
bullet("Cofactors can be subdivided into two groups: metals and small organic molecules."),
keyterm("Coenzymes", "Cofactors that are small organic molecules."),

h2("3.2 Cofactor Classification Tree (PPT Slide 23)"),
important("COFACTOR CLASSIFICATION DIAGRAM:"),
makeTable(
  ["Branch", "Sub-branch", "Binding Type"],
  [
    ["Essential Ions","Activator ions","LOOSELY bound to enzyme"],
    ["Essential Ions","Metal ions (Metalloenzymes)","TIGHTLY bound to enzyme"],
    ["Coenzymes","Cosubstrates","LOOSELY bound — altered during reaction, regenerated by other enzymes"],
    ["Coenzymes","Prosthetic groups","TIGHTLY bound — remains bound to the enzyme during the reaction"],
  ],
  [2200, 2600, 2760]
),
para(""),

h2("3.3 Classes of Coenzymes"),
keyterm("Cosubstrates", "Altered during the reaction and regenerated by other enzymes."),
keyterm("Prosthetic Groups", "Remains TIGHTLY BOUND to the enzyme during the reaction."),
keyterm("Vitamin-derived Coenzymes", "Cannot be synthesized by mammals; must be obtained from nutrients."),

h2("3.4 Vitamin–Cofactor Relationship"),
bullet("Cofactors are formed from vitamins. NOT all vitamins are cofactors."),
bullet("All water-soluble vitamins (EXCEPT Vitamin C) are converted/activated to cofactors."),
bullet("Of fat-soluble vitamins, ONLY Vitamin K is converted to a cofactor."),
bullet("Cofactors may also act as specific functional group providers (methyl and acyl groups)."),
important("Why do we need cofactors? — Enzymes alone CANNOT provide all the functional groups for all biochemical reactions."),

h2("3.5 Complete Coenzyme–Vitamin–Role Table (PPT Slide 27)"),
important("THIS TABLE IS DIRECTLY FROM THE PPT — MEMORIZE IT:"),
makeTable(
  ["Coenzyme", "Vitamin (Source)", "Role"],
  [
    ["NAD(P)","Niacin (Vitamin B₃)","Redox (electron carrier — oxidation-reduction)"],
    ["FAD/FMN","Riboflavin (Vitamin B₂)","Redox (electron carrier)"],
    ["Coenzyme A (CoA)","Pantothenic acid (Vitamin B₃*)","Acyl transfer"],
    ["TPP (Thiamine pyrophosphate)","Thiamine (Vitamin B₁)","Transfer of 2-carbon units"],
    ["PLP (Pyridoxal phosphate)","Pyridoxine (Vitamin B₆)","Amino acid reactions"],
    ["Lipoamide","— (not vitamin-derived)","Acyl transfer"],
    ["Ubiquinone","— (not vitamin-derived)","Electron carrier"],
  ],
  [2520, 2520, 2520]
),
para(""),

h2("3.6 Cofactor–Enzyme Pairing Table (PPT Slide 22)"),
important("THIS TABLE IS DIRECTLY FROM THE PPT:"),
makeTable(
  ["Category", "Cofactor", "Associated Enzyme"],
  [
    ["Coenzyme","Thiamine pyrophosphate","Pyruvate dehydrogenase"],
    ["Coenzyme","Flavin adenine nucleotide (FAD)","Monoamine oxidase"],
    ["Coenzyme","Nicotinamide adenine dinucleotide (NAD)","Lactate dehydrogenase"],
    ["Coenzyme","Pyridoxal phosphate (PLP)","Glycogen phosphorylase"],
    ["Coenzyme","Coenzyme A (CoA)","Acetyl CoA carboxylase"],
    ["Coenzyme","Biotin","Pyruvate carboxylase"],
    ["Coenzyme","5'-Deoxyadenosyl cobalamin","Methylmalonyl mutase"],
    ["Coenzyme","Tetrahydrofolate","Thymidylate synthase"],
    ["Metal","Zn²⁺","Carbonic anhydrase"],
    ["Metal","Zn²⁺","Carboxypeptidase"],
    ["Metal","Mg²⁺","EcoRV"],
    ["Metal","Mg²⁺","Hexokinase"],
    ["Metal","Ni²⁺","Urease"],
    ["Metal","Mo (Molybdenum)","Nitrate reductase"],
    ["Metal","Se (Selenium)","Glutathione peroxidase"],
    ["Metal","Mn²⁺","Superoxide dismutase"],
    ["Metal","K⁺","Propionyl CoA carboxylase"],
  ],
  [1500, 2700, 3360]
),
para(""),

h2("3.7 NAD and NADP (Niacin — Vitamin B₃)"),
bullet("Niacin (Vitamin B₃) exists as two forms:"),
subbullet("Nicotinic acid — has –COOH group on pyridine ring."),
subbullet("Nicotinamide — has –CONH₂ group on pyridine ring."),
bullet("Two cofactor forms of niacin: NAD and NADP."),
bullet("These cofactors are NOT tightly held by the enzyme and may be REUSED for reaction after reaction (cosubstrates — loosely bound)."),
bullet("NAD(P)⁺ undergo reduction to NADH and NADPH, accepting a hydride ion (two electrons and one proton) from an oxidizable substrate."),

h2("3.8 FMN and FAD (Riboflavin — Vitamin B₂)"),
bullet("Flavin mononucleotide (FMN) and flavin adenine dinucleotide (FAD) are derived from Riboflavin (Vitamin B₂)."),
bullet("Functions:"),
subbullet("Involved in metabolism of carbohydrates, fats, and proteins (Flavin dehydrogenase / Flavo proteins)."),
subbullet("Hydrogen carriers in the respiratory chain."),

h2("3.9 FAD Catalyzing Reaction (PPT Slide 32 — FAD CATALYZING REACTION)"),
important("★ FAD REACTION CYCLE (WRITE IN EXAM):"),
makeTable(
  ["Reaction", "Equation / Description"],
  [
    ["FAD catalyzed (Succinate dehydrogenase)","Succinate + FAD → Fumarate + FADH₂   (catalyzed by succinate dehydrogenase)"],
    ["Structure of Succinate","-OOC–CH₂–CH₂–COO⁻ (two –CH₂ groups between two carboxylates)"],
    ["Structure of Fumarate","-OOC–CH=CH–COO⁻ (trans double bond formed)"],
    ["FMN reaction (Respiratory chain)","NADH dehydrogenase (FMN) + 2H⁺ + 2e⁻ → NADH dehydrogenase (FMNH₂)"],
  ],
  [2400, 5160]
),
para(""),

h2("3.10 Thiamine Pyrophosphate — TPP (Vitamin B₁)"),
bullet("Thiamine (Vitamin B₁) + ATP → Thiamine Pyrophosphate (TPP) + AMP   [catalyzed by TPP-synthetase]."),
bullet("TPP structure: Pyrimidine ring + Thiazolium ring + Pyrophosphate group."),
bullet("Key feature: Contains an ACIDIC PROTON on the thiazolium carbon — this is the reactive site."),

h3("Reactions in which TPP is a cofactor (PPT Slide 34):"),
makeTable(
  ["Enzyme", "Reaction / Pathway"],
  [
    ["Pyruvate decarboxylase","Alcohol fermentation — converts pyruvate to acetaldehyde"],
    ["Pyruvate dehydrogenase","Synthesis of Acetyl-CoA"],
    ["Alpha-ketoglutarate dehydrogenase","Citric acid cycle"],
    ["Transketolase reaction","Carbon-fixation reactions of photosynthesis"],
    ["Acetolactase synthetase","Valine and leucine biosynthesis"],
  ],
  [3000, 4560]
),
para(""),

h2("3.11 Pyridoxal Phosphate (PLP) — Vitamin B₆"),
bullet("Derived from Vitamin B₆ family (Pyridoxine). Phosphorylated to form PLP."),
bullet("PLP is a prosthetic group for enzymes catalyzing reactions involving amino acids:"),
subbullet("Isomerizations, Decarboxylations, Side chain eliminations or replacements."),
bullet("PLP functions as an intermediate carrier of amino groups at the active site of aminotransferases."),
bullet("Reversible transformations between:"),
subbullet("Pyridoxal phosphate (aldehyde form) — can ACCEPT an amino group."),
subbullet("Pyridoxamine phosphate (aminated form) — can DONATE its amino group to an α-keto acid."),
bullet("PLP is COVALENTLY BOUND to the enzyme's active site through an aldimine (Schiff base) linkage to the –amino group of a Lys residue."),

h2("3.12 Pantothenic Acid — Coenzyme A (CoA)"),
bullet("In its active form serves as a cofactor for Coenzyme A (CoA) and the acyl carrier protein (ACP)."),
bullet("A hydroxyl group of pantothenic acid is joined to a modified ADP moiety by a phosphate ester bond."),
bullet("Its carboxyl group is attached to β-mercaptoethylamine in amide linkage."),
bullet("The –SH group of the mercaptoethylamine moiety forms a thioester with acetate in Acetyl-CoA."),
bullet("Coenzyme A is used for the synthesis of Acetylcholine — an important neurotransmitter in the autonomic nervous system and brain."),

h2("3.13 Metal Ion Cofactors"),
important("One third of all known enzymes requires metal ions for catalytic activity."),
keyterm("Metalloenzymes", "Contain TIGHTLY BOUND metal ion cofactors."),
bullet("Catalytic role — Transition metal ions"),
bullet("Structural role — Na⁺, K⁺, Ca²⁺"),
bullet("Structural AND Catalytic — Mg²⁺, Zn²⁺"),
h3("Three Ways Metal Ions Participate in Catalysis:"),
bullet("By binding to substrates to orient them properly for reaction."),
bullet("By mediating oxidation-reduction reactions through reversible change in the metal ion oxidation state."),
bullet("By electrostatically stabilizing or shielding negative charges."),
bullet("Metalloproteases use a metal ion to hold water in place so it can be ionized to act as a nucleophile."),
divider(),

// ══════════════════════════════════════════════════════
// CH4: MECHANISMS
// ══════════════════════════════════════════════════════
h1("CHAPTER 4 — MECHANISMS OF ENZYME ACTION"),
divider(),

h2("4.1 Introduction"),
bullet("Enzymes are protein catalysts."),
bullet("Enzymes do two important things: (1) Recognize very specific substrates. (2) Perform specific chemical reactions at fantastic speeds."),

h2("4.2 Enzyme Active Site"),
bullet("The active site is a specialized region of the protein where the enzyme interacts with the substrate."),
bullet("Formed in the three-dimensional structure by a collection of different amino acids (active-site residues) that may or may not be adjacent in the primary sequence."),
bullet("Interactions between the active site and the substrate occur via: Hydrophobic interactions, Electrostatic interactions, Hydrogen bonding, Van der Waals interactions."),
bullet("The specificity of binding depends on the precisely defined arrangement of atoms in an active site."),

h2("4.3 Lock and Key Model"),
para([bold("Concept: "), norm("The LOCK AND KEY model uses complementarity between the enzyme active site (the lock) and the substrate (the key).")]),
bullet("The substrate must fit correctly into the active site: right size and shape; charges in correct place; right hydrogen-bond donors and acceptors; right hydrophobic patches."),

h2("4.4 Induced Fit Model"),
para([bold("Concept: "), norm("The structure of the enzyme is DIFFERENT depending on whether the substrate is bound or not.")]),
bullet("The enzyme changes shape (undergoes a conformational change) on binding the substrate."),
bullet("Conformational change converts the enzyme into a new structure in which the substrate and catalytic groups on the enzyme are properly arranged to accelerate the reaction."),
bullet("'Bad' substrates CANNOT cause this conformational change."),
h3("Hexokinase — Induced Fit Example:"),
bullet("Catalyzes the transfer of phosphate from ATP to the 6-hydroxyl group of glucose."),
bullet("Glucose–OH is very similar in reactivity to water."),
bullet("Hexokinase catalyzes the transfer to glucose about 10⁵ times faster than to water."),
bullet("The other parts of glucose (besides OH) are necessary to induce the enzyme to change its conformation and become an efficient catalyst."),

h2("4.5 Four Catalytic Strategies"),
h3("1. Covalent Catalysis:"),
bullet("The active site contains a reactive group, usually a powerful NUCLEOPHILE, that becomes temporarily covalently modified during catalysis."),
bullet("Example: Chymotrypsin."),
h3("2. General Acid-Base Catalysis:"),
bullet("A molecule OTHER THAN WATER plays the role of a proton donor or acceptor."),
bullet("Example: Lysozyme."),
h3("3. Metal Ion Catalysis:"),
bullet("A metal ion may serve as an electrophilic catalyst, stabilizing a negative charge on a reaction intermediate."),
bullet("The metal ion may generate a nucleophile by increasing the acidity of a nearby molecule."),
bullet("The metal ion may bind to substrate, increasing the number of interactions with the enzyme and thus binding energy."),
h3("4. Catalysis by Approximation:"),
bullet("Many reactions include two distinct substrates. The reaction rate may be considerably enhanced by bringing the two substrates together along a single binding surface on an enzyme."),

h2("4.6 Chymotrypsin — Covalent Catalysis"),
bullet("Chymotrypsin: proteolytic enzyme in the digestive system."),
bullet("Cleaves peptide bonds selectively on the carboxyl terminal side of large hydrophobic amino acids: tryptophan, tyrosine, phenylalanine, methionine."),
bullet("Employs a powerful nucleophile (Serine 195) to attack the unreactive carbonyl group of the substrate."),
important("CHARGE-RELAY NETWORK: Serine 195 — Histidine 57 — Aspartate 102 ('proton shuttle')"),
para(""),
important("★ CHYMOTRYPSIN 5-STEP CATALYTIC CYCLE (WRITE IN EXAM):"),
makeTable(
  ["Step","Event","Key Molecule/State"],
  [
    ["Step 1","Charge-relay system (His57 + Asp102) removes a proton from Ser 195","Ser195 becomes a STRONGER NUCLEOPHILE"],
    ["Step 2","Activated Ser195 attacks the peptide bond","FIRST TETRAHEDRAL INTERMEDIATE formed (transient)"],
    ["Step 3","Amino terminal peptide released; His57 donates a proton to the new amino group","ACYL-SER195 INTERMEDIATE formed"],
    ["Step 4","His57 and Asp102 activate a water molecule; water attacks the Acyl-Ser195","SECOND TETRAHEDRAL INTERMEDIATE formed"],
    ["Step 5","Charge-relay donates proton to Ser195; tetrahedral intermediate breaks down","CARBOXYL-TERMINAL PEPTIDE released; enzyme regenerated"],
  ],
  [720, 4320, 2520]
),
para(""),

h2("4.7 Lysozyme — Acid-Base Catalysis"),
bullet("Discovered in 1922 by Alexander Fleming during a cold. Mucosal secretions and tears inhibited bacterial growth."),
bullet("He named the enzyme 'lysozyme' (bacteria LYSing enZYME)."),
bullet("He believed it might be an antibiotic but found proteins are not rugged enough for this role."),
bullet("Lysozyme cleaves polysaccharides that give structural integrity to bacterial cell walls."),
bullet("Cell wall polysaccharides are composed of two kinds of glucose derivatives connected by β(1→4) linkages:"),
subbullet("NAG: N-acetylglucosamine"),
subbullet("NAM: N-acetylmuramic acid"),
bullet("Chitin (poly β(1→4) NAG — in shells of crustaceans) is also a substrate."),

h3("★ LYSOZYME MECHANISM — TWO KEY CATALYTIC RESIDUES (PPT Slides 14–19):"),
important("TWO KEY RESIDUES: Glutamate 35 (#35-Glu) and Aspartate 52 (#52-Asp)"),
makeTable(
  ["Step","Residue","Event"],
  [
    ["Step 1","Glu35 (acid form –COOH)","Glu35 acts as a GENERAL ACID — donates a proton to the oxygen of the glycosidic bond being cleaved. The C-B-A sugar ring (NAG/NAM) distorts toward an oxocarbenium (carbocation) intermediate."],
    ["Step 2","Asp52 (carboxylate –COO⁻)","Asp52 acts as a NUCLEOPHILE / electrostatic stabilizer — stabilizes the positive charge (oxocarbenium ion) on the carbocation intermediate (C1 of the sugar ring)."],
    ["Step 3","Water molecule + Glu35","Water molecule acts as a nucleophile, assisted by Glu35 (now acting as a GENERAL BASE). Water attacks the carbocation intermediate."],
    ["Step 4","Both residues","Product (cleaved sugar) is released; enzyme returns to original state. Cycle repeats."],
  ],
  [720, 1800, 5040]
),
para(""),
important("Note from diagram: The sugar ring is shown in a half-chair conformation (C-B-A). Glu35 is on the left side (–COOH). Asp52 is on the right side (–COO⁻). HO = hydroxyl on the anomeric carbon."),

h2("4.8 Carbonic Anhydrase — Metal Ion Catalysis"),
bullet("Catalyzes: CO₂ + H₂O → HCO₃⁻ + H⁺."),
bullet("Enhances rate by more than 10⁶. At these rates, limiting factor is how fast molecules diffuse to active site!"),
important("★ CARBONIC ANHYDRASE ZN²⁺ CYCLE (WRITE IN EXAM):"),
makeTable(
  ["Step","Event"],
  [
    ["Step 1","Zn²⁺ cofactor is at the active site. Zn²⁺ binds a water molecule."],
    ["Step 2","Binding of water to zinc reduces the pKa for water from its normal 15.7 DOWN TO 7. This allows formation of the strong hydroxide (HO⁻) nucleophile at neutral pH."],
    ["Step 3","The enzyme positions CO₂ for nucleophilic attack by the hydroxide (HO⁻)."],
    ["Step 4","Nucleophilic attack results in formation of BICARBONATE (HCO₃⁻)."],
    ["Step 5","Water displaces the bicarbonate product, starting the cycle again."],
  ],
  [900, 6660]
),
para(""),

h2("4.9 Catalysis by Approximation — Adenylate Kinase"),
bullet("Substrate binding induces large structural changes in adenylate kinase."),
bullet("Direct interactions with the nucleotide substrate lead to local structural rearrangements (movement of the P-loop) within the enzyme, which in turn allow more extensive changes (closing down of the top domain)."),
bullet("Binding of the second substrate, NMP, induces additional conformational changes."),
bullet("Both sets of changes ensure that a catalytically competent conformation is formed ONLY when BOTH the donor and acceptor are bound."),
bullet("The enzyme holds its two substrates close together and oriented to stabilize the transition state that leads to transfer of a phosphoryl group from ATP to NMP."),
divider(),

// ══════════════════════════════════════════════════════
// CH5: KINETICS
// ══════════════════════════════════════════════════════
h1("CHAPTER 5 — ENZYME KINETICS"),
divider(),

h2("5.1 Introduction"),
keyterm("Enzyme Kinetics","Quantitative measurement of the rates of enzyme-catalyzed reactions AND the systematic study of factors that affect these rates."),
bullet("Began in 1902 when Adrina Brown investigated the rate of hydrolysis of sucrose catalyzed by the yeast enzyme invertase."),
bullet("Brown demonstrated: when sucrose concentration >> enzyme concentration, reaction rate becomes independent of sucrose concentration."),
bullet("Brown's proposal: the overall reaction is composed of two elementary reactions — substrate forms a complex with the enzyme, which subsequently decomposes to products and enzyme."),

h2("5.2 Brown's Two-Step Reaction Model"),
boxPara("         k1          k2"),
boxPara("E + S  ⇌  ES  →  P + E"),
boxPara("        k-1"),
bullet("E = Enzyme, S = Substrate, ES = Enzyme-Substrate Complex, P = Product."),
bullet("When [S] is high enough to entirely convert the enzyme to ES form, the second step (k2) becomes the RATE-LIMITING STEP."),
bullet("The overall rate becomes insensitive to further increases in substrate concentration."),

h2("5.3 Two Assumptions to Derive the Rate Equation"),
makeTable(
  ["Assumption","Approach","Key Idea"],
  [
    ["Rapid Equilibrium Assumption","Michaelis-Menten Approach","Assumes RAPID EQUILIBRIUM between E and S to form [ES] complex."],
    ["Steady-State Assumption (SSA)","Briggs and Haldane Approach","[ES] remains CONSTANT (d[ES]/dt = 0) — its rate of synthesis equals its rate of consumption."],
  ],
  [2700, 2400, 2460]
),
para(""),

h2("5.4 Michaelis-Menten Equation Derivation (SSA)"),
bullet("Based on steady-state assumption: d[ES]/dt = 0"),
boxPara("d[ES]/dt = k1[E][S] – k-1[ES] – k2[ES] = 0"),
bullet("Solving for [ES]:  [ES] = [E][S]k1 / (k-1 + k2)"),
bullet("Define KM (Michaelis constant): KM = (k-1 + k2) / k1"),
bullet("Therefore: KM = [E][S] / [ES]"),
bullet("Using enzyme mass conservation: [E]total = [E] + [ES]"),
important("★ THE MICHAELIS-MENTEN EQUATION:"),
boxPara("v₀ = Vmax[S] / (KM + [S])"),
bullet("v₀ = Initial velocity, Vmax = Maximum velocity, [S] = Substrate concentration, KM = Michaelis constant."),

h2("5.5 Michaelis Constant — KM"),
keyterm("KM","The substrate concentration required to reach HALF-MAXIMAL velocity (Vmax/2)."),
bullet("KM is a measure of a substrate's affinity for the enzyme."),
bullet("SMALL KM → substrate binds TIGHTLY to the enzyme; enzyme is saturated at low concentrations."),
important("WHEN V₀ = Vmax/2, then KM = [S]"),
bullet("KA = equilibrium association constant (units: M⁻¹). KD = equilibrium dissociation constant (units: M)."),
bullet("Tight binding = low dissociation constant (KD) + high association constant (KA)."),

h2("5.6 Vmax and Kcat"),
bullet("Vmax = product of the catalytic rate constant (kcat) and the concentration of enzyme."),
boxPara("V = Kcat [Enzyme] [S] / (KM + [S])"),
keyterm("Kcat (Turnover Number)","Equal to K₂. Measures the number of substrate molecules 'turned over' by enzyme per second. Higher Kcat = more substrates per second."),
bullet("Considering total enzyme concentration, the maximal rate the enzyme can attain is Vmax."),

h2("5.7 Features of Michaelis-Menten Kinetics"),
bullet("Assumes the formation of an Enzyme-Substrate complex."),
bullet("Assumes that the ES complex is in rapid equilibrium with free enzyme."),
bullet("Breakdown of ES to form products is assumed to be SLOWER than: (1) formation of ES, AND (2) breakdown of ES to reform E and S."),

h2("5.8 Progress Curve — Steady State (PPT)"),
bullet("Except the transition phase of the reaction (before the shaded block), [ES] remains CONSTANT until the substrate is nearly exhausted."),
bullet("Hence synthesis of ES must equal its consumption over the course of the reaction — ES maintains STEADY STATE."),

h2("5.9 Lineweaver-Burk Plot (Double Reciprocal Plot)"),
bullet("Direct measurement of Vmax often requires impractically high [S] to achieve saturating conditions."),
bullet("The Michaelis-Menten equation is algebraically transformed:"),
h3("Starting with MM equation: v₀ = Vmax[S] / (KM + [S])"),
h3("Take reciprocal of both sides:"),
boxPara("1/v₀ = (KM/Vmax) × (1/[S]) + 1/Vmax"),
bullet("This is the equation for a straight line: y = ax + b, where y = 1/v₀ and x = 1/[S]."),
bullet("Y-intercept = 1/Vmax"),
bullet("Slope = KM/Vmax"),
bullet("X-intercept = −1/KM"),
important("The Lineweaver-Burk (double-reciprocal) plot allows more accurate determination of Vmax and KM, and is very useful in distinguishing between types of enzymatic reaction mechanisms."),

h2("5.10 Kinetics of Isosteric vs Allosteric Enzymes"),
h3("Isosteric Enzymes (one conformation):"),
bullet("Efficiency of substrate binding declines constantly with increasing [A] because the number of free binding sites is constantly decreasing."),
bullet("Follow HYPERBOLIC (Michaelis-Menten) kinetics."),
h3("Allosteric Enzymes:"),
bullet("Binding efficiency initially RISES with increasing [A] because free enzyme is in a low-affinity conformation (square symbols = T state), which is gradually converted into a higher-affinity form (round symbols = R state)."),
bullet("At high [A] values, a lack of free binding sites becomes noticeable and binding strength decreases again."),
bullet("Show SIGMOIDAL (S-shaped) kinetics — NOT hyperbolic."),

h2("5.11 Factors Affecting Enzyme Activity"),
bullet("Physical quantities: temperature, pressure."),
bullet("Chemical properties: pH value, ionic strength."),
bullet("Concentrations of relevant substrates, cofactors, and inhibitors."),

h3("pH Dependency:"),
bullet("Activity plotted against pH = BELL-SHAPED CURVE."),
bullet("Bell shape results from ionizable groups in the side chains of amino acid residues essential for catalysis."),
bullet("Example of two groups: Basic group B (pKa = 8) — must be PROTONATED to be active. Acidic amino acid AH (pKa = 6) — only active in DISSOCIATED state."),
bullet("At optimum pH of 7 — around 90% of both groups are present in the active form."),
bullet("At higher and lower values, one or the other group passes into the inactive state."),

h3("Temperature Dependency:"),
bullet("Temperature dependency is ASYMMETRIC."),
bullet("With increasing temperature, increased thermal movement initially leads to RATE ACCELERATION."),
bullet("At a certain temperature, the enzyme becomes unstable and its activity is LOST within a narrow temperature difference due to DENATURATION."),

h2("5.12 Bisubstrate Kinetics"),
bullet("Most reactions include two substrates and two products: A + B → P + Q."),
bullet("Bisubstrate reactions: transfer of a functional group (phosphoryl or ammonium), OR electrons are transferred between substrates."),
bullet("Two classes: Sequential Displacement and Double Displacement (Ping-Pong)."),

h3("Sequential Displacement:"),
bullet("All substrates MUST BIND TO THE ENZYME before any product is released."),
bullet("Two types: Ordered (defined binding sequence) and Random."),
bullet("ORDERED EXAMPLE — Lactate Dehydrogenase (LDH): The coenzyme (NADH) ALWAYS binds first; lactate is ALWAYS released first."),
bullet("RANDOM EXAMPLE — Creatine Kinase: Formation of phosphocreatine and ADP from ATP and creatine. Order of addition of substrates and release of products is random."),

h3("★ DOUBLE DISPLACEMENT — PING-PONG MECHANISM (PPT CYCLE):"),
bullet("One or more products are RELEASED BEFORE all substrates bind the enzyme."),
bullet("First substrate A is bound and immediately cleaved. A part remains BOUND to the enzyme (modified enzyme intermediate)."),
bullet("After the first product C is released, part of A is transferred to second substrate B."),
important("PING-PONG EXAMPLE — ASPARTATE AMINOTRANSFERASE CYCLE:"),
makeTable(
  ["Step","Action","Substrate/Product"],
  [
    ["PING — Step 1","Aspartate (Substrate A) BINDS to the enzyme","Aspartate bound"],
    ["PING — Step 2","Enzyme removes aspartate's amino group → forms SUBSTITUTED ENZYME INTERMEDIATE","Modified enzyme contains the amino group"],
    ["PING — Step 3","First product OXALOACETATE departs","Oxaloacetate released (Product C)"],
    ["PONG — Step 4","Second substrate α-KETOGLUTARATE binds to the enzyme","α-Ketoglutarate (Substrate B) bound"],
    ["PONG — Step 5","α-Ketoglutarate accepts the amino group from the modified enzyme","Amino group transferred"],
    ["PONG — Step 6","GLUTAMATE released as the final product","Glutamate = Product P"],
  ],
  [1800, 3600, 2160]
),
para(""),
divider(),

// ══════════════════════════════════════════════════════
// CH6: INHIBITION
// ══════════════════════════════════════════════════════
h1("CHAPTER 6 — ENZYME INHIBITION"),
divider(),

h2("6.1 Introduction"),
bullet("The activity of many enzymes can be inhibited by the binding of specific small molecules and ions."),
bullet("Inhibiting enzyme activity serves as a major control mechanism in biological systems."),

h2("6.2 What are Inhibitors?"),
bullet("Inhibitors are chemicals that REDUCE the rate of enzymatic reactions."),
bullet("They are usually specific and work at low concentrations."),
bullet("They block the enzyme but do NOT usually destroy it."),
bullet("Many drugs and poisons are inhibitors of enzymes."),
bullet("Reversible inhibition: equilibrium exists between enzyme and inhibitor."),
bullet("Irreversible inhibition: inhibition progressively increases with time."),

h2("6.3 Competitive Inhibition"),
bullet("Both substrate S and inhibitor I COMPETE FOR THE SAME ACTIVE SITE."),
bullet("Only the ES complex leads to product formation."),
bullet("Increasing the amount of substrate CAN overcome inhibition → Vmax CAN be attained."),
bullet("The apparent KM (KMapp) is INCREASED: KMapp = KM × (1 + [I]/Ki)."),
important("Competitive Inhibition: Vmax UNCHANGED | KM INCREASED"),
h3("Lineweaver-Burk: INTERSECT AT THE Y-AXIS (Vmax unchanged, slope changes)."),
h3("Example: Malonate inhibits succinate dehydrogenase (malonate is structurally similar to succinate)."),

h2("6.4 Non-Competitive Inhibition"),
bullet("Non-competitive inhibitor binds to enzyme at a site DISTINCT from the substrate binding site."),
bullet("It can bind to BOTH free enzyme AND enzyme–substrate complex."),
bullet("The binding of the inhibitor has NO effect on substrate binding, and vice versa."),
bullet("'I' does NOT interfere with formation of ES — NON-COMPETITIVE INHIBITION CANNOT BE REVERSED BY INCREASING [S]."),
important("Non-competitive Inhibition: Vmax DECREASED | KM UNCHANGED"),
boxPara("Vmax,app < Vmax     |     KM,app = KM"),
h3("Lineweaver-Burk: INTERSECT AT THE X-AXIS (KM unchanged, intercepts change)."),
h3("Example: Fructose 1,6-bisphosphatase inhibited by AMP."),
bullet("High AMP signals that ATP levels are low → gluconeogenesis should be shut down while glycolysis is turned on."),

h2("6.5 Uncompetitive Inhibition"),
bullet("Uncompetitive inhibitor does NOT bind to free enzyme."),
bullet("Binds reversibly to the ENZYME-SUBSTRATE COMPLEX (ES) only → forms inactive ESI complex."),
bullet("The ESI complex does NOT form a product."),
bullet("CANNOT be reversed by increasing [S] — increasing [S] actually FAVORS inhibition."),
important("Uncompetitive Inhibition: BOTH Vmax AND KM DECREASED"),
boxPara("Vmax,app < Vmax     |     KM,app < KM"),
bullet("At low [S], uncompetitive inhibitors have LITTLE EFFECT because lower KMapp offsets the decreased Vmax,app."),
h3("Lineweaver-Burk: TWO PARALLEL LINES."),
h3("Example: Alkaline phosphatase inhibited by phenylalanine."),

h2("6.6 Comparison — All Three Types of Reversible Inhibition (PPT Slide 21 Cartoon Summary)"),
makeTable(
  ["Feature","Competitive","Non-Competitive","Uncompetitive"],
  [
    ["Inhibitor binds to","Free enzyme [E] only — competes for ACTIVE SITE","Free enzyme [E] OR [ES] complex — DIFFERENT SITE","[ES] complex ONLY"],
    ["Mechanism equation","E+S→ES→E+P; +I: E+I→EI (EI cannot bind S)","E+I→EI; ES+I→EIS (both dead ends)","E+S→ES→E+P; ES+I→ESI (dead end)"],
    ["Overcome by ↑[S]?","YES","NO","NO — ↑[S] FAVORS inhibition"],
    ["Effect on Vmax","UNCHANGED","DECREASED","DECREASED"],
    ["Effect on KM","INCREASED","UNCHANGED","DECREASED"],
    ["Lineweaver-Burk","Intersect Y-axis","Intersect X-axis","TWO PARALLEL LINES"],
    ["Direct plot","KM′ shifts right; Vmax same","Vmax′ drops; KM same","Both KM and Vmax decrease"],
  ],
  [1920, 2160, 2160, 1320]
),
para(""),

h2("6.7 Irreversible Inhibition"),
bullet("Inhibitor forms a COVALENT LINKAGE with the enzyme — CANNOT be removed."),
bullet("May act at, near, or remote from the active site."),
bullet("May NOT be displaced by addition of excess substrate."),
bullet("The basic structure of the enzyme is modified to the degree that it ceases to work."),
bullet("Michaelis-Menten kinetics CANNOT be applied to irreversible inhibition."),
bullet("Michaelis-Menten PLOT looks like non-competitive inhibition: Vmax,app < Vmax; KM,app = KM."),
bullet("Irreversible inhibitors COVALENTLY MODIFY the enzyme → permanent inhibition."),
h3("Examples of Irreversible Inhibitors:"),
bullet("CYANIDE — covalently binds mitochondrial cytochrome oxidase; inhibits ALL reactions of electron transport."),
bullet("PENICILLIN — irreversibly inhibits bacterial peptidase."),
divider(),

// ══════════════════════════════════════════════════════
// CH7: ALLOSTERIC REGULATION
// ══════════════════════════════════════════════════════
h1("CHAPTER 7 — ALLOSTERIC ENZYME REGULATION"),
divider(),

h2("7.1 Introduction"),
para([bold("Allosteric "), norm("— from the Greek allos ('other') and stereos ('structure').")]),
bullet("Allosteric regulation = control of enzyme activity through binding of effectors at sites OTHER THAN the active site."),

h2("7.2 ATCase — Classic Example"),
keyterm("ATCase (Aspartate Carbamoyltransferase)","Key enzyme of pyrimidine biosynthesis in E. coli."),
bullet("ATCase catalyzes the transfer of a carbamoyl residue from carbamoyl phosphate to the amino group of L-aspartate."),
bullet("ATCase is INHIBITED by CTP (cytidine triphosphate) — the end product of pyrimidine anabolic metabolism."),
bullet("ATCase is ACTIVATED by ATP — the precursor."),
bullet("The rate is FAST in the absence of high CTP concentrations; DECREASES as CTP concentration increases."),
bullet("This exemplifies FEEDBACK (end-product) INHIBITION."),
important("CTP is structurally QUITE DIFFERENT from the substrates. CTP binds to a site DISTINCT from the active site → called ALLOSTERIC or REGULATORY SITE."),
keyterm("Allosteric Inhibitor","CTP is an example."),

h2("7.3 ATCase Structure"),
bullet("Allosteric enzymes are almost always OLIGOMERS with 2–12 subunits."),
bullet("ATCase consists of: Six CATALYTIC subunits (blue) + Six REGULATORY subunits (yellow)."),
important("In ATCase, the catalytic sites and regulatory sites are on SEPARATE POLYPEPTIDE CHAINS."),

h2("7.4 T and R States"),
keyterm("T State","(Tense) — the LESS ACTIVE state."),
keyterm("R State","(Relaxed) — the MORE ACTIVE state."),
important("★ T STATE / R STATE EQUILIBRIUM TABLE:"),
makeTable(
  ["Condition","Effect","State Favored"],
  [
    ["Increasing aspartate (substrate) concentration","Equilibrium shifts toward R form","R State (more active)"],
    ["ATP binds to regulatory subunits","Stabilizes R conformation","R State (more active)"],
    ["CTP binds to regulatory subunits (same sites as ATP)","Promotes transition to T state","T State (less active)"],
  ],
  [2700, 3000, 1860]
),
para(""),

h2("7.5 Allosteric Kinetics — Sigmoidal Curves (PPT Slide 10 Figures)"),
important("★ FIGURE 10.11 — ATCase Sigmoidal Kinetics:"),
bullet("A plot of product formation (N-carbamoylaspartate formation) as a function of substrate concentration ([Aspartate], mM) produces a SIGMOIDAL CURVE."),
bullet("The sigmoidal curve occurs because binding of substrate to one active site FAVORS THE CONVERSION OF THE ENTIRE ENZYME INTO THE R STATE, increasing activity at the other active sites."),
bullet("Thus, the active sites show COOPERATIVITY."),
important("★ FIGURE 10.12 — Basis for the Sigmoidal Curve (R-state vs T-state curves):"),
bullet("An allosteric enzyme can be imagined as a MIXTURE OF TWO MICHAELIS-MENTEN ENZYMES:"),
subbullet("One with a HIGH value of KM — corresponds to the T STATE (low affinity)."),
subbullet("One with a LOW value of KM — corresponds to the R STATE (high affinity)."),
bullet("As the concentration of substrate is increased, the equilibrium shifts from the T state to the R state, which results in a STEEP RISE in activity with respect to substrate concentration → this produces the sigmoidal curve."),
bullet("The R-state curve is ABOVE the actual curve at low [S]; the T-state curve is BELOW the actual curve."),
bullet("Allosteric enzymes DO NOT FOLLOW MICHAELIS-MENTEN KINETICS."),
divider(),

// ══════════════════════════════════════════════════════
// CH8: ISOZYMES
// ══════════════════════════════════════════════════════
h1("CHAPTER 8 — ISOZYMES (ISOENZYMES)"),
divider(),

h2("8.1 Definition"),
keyterm("Isozymes / Isoenzymes","Enzymes that differ in amino acid sequence yet catalyze THE SAME REACTION. Encoded by different genes. Have different kinetic and regulatory properties."),
bullet("The same reaction catalyzed by two or more different molecular forms of an enzyme."),
bullet("May occur in the same species, in the same tissue, or even in the same cell."),

h2("8.2 Lactate Dehydrogenase (LDH) — Classic Isozyme"),
bullet("One of the FIRST enzymes found to have isozymes."),
bullet("LDH in vertebrate tissues exists as at least FIVE DIFFERENT ISOZYMES."),
bullet("All LDH isozymes contain FOUR POLYPEPTIDE CHAINS (Tetramer)."),
bullet("Each monomer is formed by a peptide chain of 334 amino acids."),
bullet("A peptide loop (pink) formed by amino acid residues 98–111 is shown. In the absence of substrate and coenzyme, this partial structure is open and allows access to the substrate binding site."),

h2("8.3 H and M Subunits"),
bullet("Two isozymic polypeptide chains for LDH in human beings:"),
subbullet("H isozyme — highly expressed in the HEART."),
subbullet("M isozyme — found in SKELETAL MUSCLE."),
bullet("Differences between M and H subunits are mainly CONSERVATIVE:"),
subbullet("Conservative: Glycine (G) and Alanine (A); Arginine (R) and Lysine (K)."),
subbullet("Non-conservative (less frequent): Lysine (K) for Glutamine (Q); Threonine (T) for Glutamic acid (E)."),
bullet("Overall, the H subunit contains MORE ACIDIC and FEWER BASIC residues than the M form → more strongly NEGATIVE charge."),

h2("8.4 Five LDH Isozymes — PPT Slide 6 Diagram"),
important("★ LDH ISOZYMES TABLE (FROM PPT SLIDE 6):"),
makeTable(
  ["Isozyme","Subunit Composition","PPT Gene/Separation info","Predominant Tissue (from gel electrophoresis)"],
  [
    ["LDH-1","H₄ (four H subunits)","Encoded by H gene","Cardiac muscle (most prominent)"],
    ["LDH-2","M₁H₃","Encoded by both H and M genes","Cardiac muscle, also blood"],
    ["LDH-3","M₂H₂","Encoded by both H and M genes","Brain, Lungs"],
    ["LDH-4","M₃H₁","Encoded by both H and M genes","Kidney, Pancreas"],
    ["LDH-5","M₄ (four M subunits)","Encoded by M gene","Skeletal muscle, Liver (most prominent)"],
  ],
  [1080, 1800, 2400, 3280]
),
para(""),
bullet("Gel electrophoresis separates all 5 isozymes. LDH5 migrates slowest (most toward anode based on less negative charge), LDH1 migrates fastest."),
bullet("Sequence alignment (PPT): LDH M chain — RYLMGERLGVHPLSCHGWVLGEHGDSSVPVWSGMNVAGCSLKTLHPDLGTD..."),
bullet("Sequence alignment (PPT): LDH H chain — RYLMAEKLGIHPSSCHGWILGEHGDSSVAVWSGVNVAGVSLQELNPEMGTD..."),
bullet("Blue bars = conservative substitutions, Pink bars = non-conservative substitutions."),

h2("8.5 Functional Differences"),
bullet("LDH-4 properties — favour RAPID REDUCTION of very low concentrations of pyruvate to LACTATE in skeletal muscle."),
bullet("LDH-1 properties — favour RAPID OXIDATION of lactate to PYRUVATE in the HEART."),

h2("8.6 Clinical Significance — Myocardial Infarction Diagnosis"),
important("★ LDH ISOZYME PATTERN IN HEART ATTACK (WRITE IN EXAM):"),
makeTable(
  ["Time After Heart Attack","LDH Pattern in Blood","Clinical Interpretation"],
  [
    ["Shortly after attack","Total LDH increases; MORE LDH-2 than LDH-1","Early sign of cardiac damage"],
    ["After 12 hours","Amounts of LDH-1 and LDH-2 are VERY SIMILAR (equal)","Transition period"],
    ["After 24 hours","MORE LDH-1 than LDH-2 (LDH-1 > LDH-2)","Confirms cardiac muscle damage"],
  ],
  [2100, 2700, 2760]
),
para(""),
bullet("Damage to heart tissue results in the RELEASE OF HEART LDH into the blood."),
bullet("Differences in the isozyme content of tissues can be used to assess the TIMING and EXTENT of heart damage due to myocardial infarction."),
divider(),

// ══════════════════════════════════════════════════════
// CH9: QUICK REVISION
// ══════════════════════════════════════════════════════
h1("CHAPTER 9 — QUICK REVISION SUMMARY"),
divider(),

h2("Key Formulas — Memorize"),
makeTable(
  ["Formula","Meaning / Context"],
  [
    ["Apoenzyme + Cofactor = Holoenzyme","Complete active enzyme"],
    ["v₀ = Vmax[S] / (KM + [S])","Michaelis-Menten equation"],
    ["1/v₀ = (KM/Vmax)(1/[S]) + 1/Vmax","Lineweaver-Burk (double reciprocal)"],
    ["KM = [S] when V₀ = Vmax/2","Definition of KM"],
    ["KM = (k-1 + k2) / k1","KM from SSA derivation"],
    ["KMapp = KM(1 + [I]/Ki)","Apparent KM in competitive inhibition"],
    ["d[ES]/dt = 0","Steady-state assumption (Briggs-Haldane)"],
    ["V = Kcat[Enzyme][S] / (KM + [S])","Michaelis-Menten with Kcat"],
  ],
  [4500, 3060]
),
para(""),

h2("Inhibition Quick Reference"),
makeTable(
  ["Type","Binds to","Vmax","KM","Lineweaver-Burk Lines"],
  [
    ["Competitive","Free E (active site)","Unchanged","Increased","Intersect Y-axis"],
    ["Non-competitive","Free E or ES (different site)","Decreased","Unchanged","Intersect X-axis"],
    ["Uncompetitive","ES complex only","Decreased","Decreased","Parallel lines"],
    ["Irreversible","Covalent bond to E (permanent)","Decreased (permanent)","Unchanged","Looks like non-competitive"],
  ],
  [1800, 2100, 1200, 1200, 2260]
),
para(""),

h2("Coenzyme Reference (Complete)"),
makeTable(
  ["Coenzyme","Vitamin","Role","Key Enzyme Example"],
  [
    ["NAD(P)","Niacin (B₃)","Redox — electron carrier","Lactate dehydrogenase"],
    ["FAD/FMN","Riboflavin (B₂)","Redox — electron carrier","Monoamine oxidase, Succinate dehydrogenase"],
    ["Coenzyme A","Pantothenic acid (B₃*)","Acyl transfer","Acetyl CoA carboxylase, Pyruvate dehydrogenase"],
    ["TPP","Thiamine (B₁)","Transfer of 2-carbon units","Pyruvate decarboxylase, α-KG dehydrogenase"],
    ["PLP","Pyridoxine (B₆)","Amino acid reactions","Glycogen phosphorylase, Aminotransferases"],
    ["Lipoamide","Not from vitamin","Acyl transfer","—"],
    ["Ubiquinone","Not from vitamin","Electron carrier","Respiratory chain"],
    ["5'-Deoxyadenosyl cobalamin","Vitamin B₁₂","Rearrangements","Methylmalonyl mutase"],
    ["Biotin","Biotin (Vitamin H)","CO₂ transfer","Pyruvate carboxylase"],
    ["Tetrahydrofolate","Folate","1-carbon transfer","Thymidylate synthase"],
  ],
  [1500, 1500, 2200, 2360]
),
para(""),

h2("All Catalytic Cycles to Know for Exam"),
makeTable(
  ["Cycle / Mechanism","Key Steps Summary","Key Residues/Molecules"],
  [
    ["Chymotrypsin (Covalent catalysis, 5 steps)","1. Ser195 activated by His57+Asp102. 2. Ser195 attacks peptide bond → 1st tetrahedral intermediate. 3. Amino terminal peptide released → Acyl-Ser195. 4. Water activated → attacks Acyl-Ser195 → 2nd tetrahedral intermediate. 5. Carboxyl terminal peptide released.","Ser 195, His 57, Asp 102"],
    ["Lysozyme (Acid-base catalysis, 4 steps)","1. Glu35 (acid) donates H⁺ to glycosidic oxygen; sugar ring forms oxocarbenium ion (C-B-A half-chair). 2. Asp52 stabilizes the carbocation. 3. Water (base-activated by Glu35) attacks the carbocation. 4. Product released; enzyme regenerated.","Glu 35, Asp 52, oxocarbenium ion"],
    ["Carbonic Anhydrase (Metal ion, 5 steps)","1. Zn²⁺ at active site. 2. Water binds Zn²⁺; pKa drops from 15.7 → 7. 3. HO⁻ nucleophile formed. 4. CO₂ positioned for attack by HO⁻ → HCO₃⁻. 5. Water displaces bicarbonate; cycle restarts.","Zn²⁺ cofactor"],
    ["Aspartate Aminotransferase Ping-Pong (6 steps)","PING: 1. Asp binds. 2. Amino group removed → modified enzyme (substituted intermediate). 3. Oxaloacetate released. PONG: 4. α-Ketoglutarate binds. 5. Accepts amino group. 6. Glutamate released.","PLP as prosthetic group"],
    ["ATCase T↔R State Equilibrium","CTP binds regulatory subunit → T state (less active). ATP binds regulatory subunit → R state (more active). Increasing aspartate → R state. The equilibrium between T and R gives the sigmoidal kinetic curve (cooperativity).","CTP (inhibitor), ATP (activator), aspartate (substrate)"],
    ["LDH Isozymes in Myocardial Infarction","Shortly after: LDH-2 > LDH-1. At 12h: LDH-1 = LDH-2. After 24h: LDH-1 > LDH-2 (diagnostic for heart damage).","LDH-1 (H₄), LDH-2 (M₁H₃)"],
  ],
  [2400, 3960, 1200]
),
para(""),

h2("Master Enzyme Examples Table"),
makeTable(
  ["Enzyme","Category / Topic","Key Fact / Number"],
  [
    ["Carbonic Anhydrase","Metal ion catalysis (Zn²⁺)","10⁶ CO₂/sec; 10⁷× faster than uncatalyzed; pKa water drops 15.7→7"],
    ["Trypsin","Enzyme specificity","Cleaves carboxyl side of Lys and Arg only"],
    ["Thrombin","Enzyme specificity (more specific than trypsin)","Cleaves Arg–Gly bonds in specific peptide sequences only"],
    ["Chymotrypsin","Covalent catalysis","Cleaves carboxyl side of Trp, Tyr, Phe, Met; Ser195+His57+Asp102"],
    ["Lysozyme","Acid-base catalysis","Cleaves β(1→4) NAG-NAM bonds; Glu35+Asp52; discovered 1922 by Fleming"],
    ["Hexokinase","Induced fit model","Transfers phosphate to glucose 10⁵× faster than to water"],
    ["Adenylate Kinase","Catalysis by approximation","P-loop movement; closes top domain when both substrates bind"],
    ["ATCase","Allosteric regulation","CTP inhibits; ATP activates; T state ↔ R state; 6 catalytic + 6 regulatory subunits"],
    ["LDH-1 (H₄)","Isozyme","Heart muscle; favours lactate → pyruvate"],
    ["LDH-5 (M₄)","Isozyme","Skeletal muscle; favours pyruvate → lactate"],
    ["Succinate dehydrogenase","Competitive inhibition example","Malonate competitively inhibits; FAD is cofactor"],
    ["Fructose 1,6-bisphosphatase","Non-competitive inhibition example","AMP inhibits; key in gluconeogenesis"],
    ["Alkaline phosphatase","Uncompetitive inhibition example","Phenylalanine is uncompetitive inhibitor; cleaves phosphate esters at alkaline pH"],
    ["Aspartate aminotransferase","Ping-Pong kinetics","Asp → oxaloacetate + α-KG → glutamate; PLP prosthetic group"],
    ["Lactate dehydrogenase (kinetics)","Sequential ordered kinetics","NADH binds first; lactate released first"],
    ["Creatine kinase","Sequential random kinetics","Phosphocreatine + ADP from ATP + creatine"],
  ],
  [2300, 2400, 2860]
),

new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 480, after: 120 },
  children: [new TextRun({ text: "— END OF NOTES — ALL SLIDES COVERED —", bold: true, size: 24, font: "Arial", color: BLUE1 })] }),
];

// ─── BUILD ────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      { reference: "subbullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 270 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: BLUE1 },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: BLUE2 },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/Enzymes_COMPLETE_Notes_v2.docx', buf);
  console.log('Done!');
});
