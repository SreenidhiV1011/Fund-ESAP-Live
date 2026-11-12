/* ════════════════════════════════════════════════════════════════════════
   esg-ic.js — ESG IC Form add-on for the Fund ESAP Tracker

   Adds: a four-tile landing page, an in-app ESG IC Form (replaces the Google
   Form), a sector-wise ESG questionnaire (.xlsx) and the post-submit pop-up.
   Nothing in index.html is rewritten: add ONE line before </body>:

       <script src="esg-ic.js"></script>

   Hidden from everyone unless the URL has ?esgic=1 (or FEATURE_ON is true).
   Responses are written to the Google Sheet through the Apps Script web app
   in Code.gs. Paste its URL into ENDPOINT below.
https://script.google.com/a/macros/blume.vc/s/AKfycbwc27E_i0cbondlrCKT_mMdytEPY9IAIKjdsJolUttSBQDfe7s_Ev5cd4o_QCWcJ4MZ/exec
   ════════════════════════════════════════════════════════════════════════ */
(function () {
'use strict';

/* ── CONFIG ─────────────────────────────────────────────────────────── */
var ENDPOINT    = '';      // Apps Script web app URL ("…/exec")
var FEATURE_ON  = false;   // false = only visible with ?esgic=1 in the URL
var REVIEW_MAIL = 'sreenidhi@blume.vc';

var enabled = FEATURE_ON || /[?&]esgic=1(&|$)/.test(location.search);
if (!enabled) return;

/* ── QUESTION BANK — copied unchanged from the ESG tool ─────────────── */
const S = { sector: '', geos: [] };
const FI=new Set(["fintech","nbfc","insurtech","wealthtech"]);
const SECTOR_NAMES={fintech:"Fintech / Payments",nbfc:"NBFC / Lending / Credit",insurtech:"Insurtech",wealthtech:"Wealthtech",saas:"SaaS / B2B software",ai:"AI / ML",consumer:"Consumer internet",healthtech:"Healthtech / Medtech",edtech:"Edtech",agritech:"Agritech",climate:"Climate / Clean energy",manufacturing:"Manufacturing",ecommerce:"E-commerce / Retail",media:"Media / Entertainment",deeptech:"Deep tech / Hardware",other:"Other"};
const SECTOR_TAGS={fintech:["fintech","nbfc"],nbfc:["nbfc","fintech"],insurtech:["insurtech","fintech"],wealthtech:["wealthtech","fintech"],saas:["saas"],ai:["ai","saas"],consumer:["consumer","ecommerce"],healthtech:["healthtech","hardware"],edtech:["edtech","consumer"],agritech:["agritech"],climate:["climate","energy"],manufacturing:["manufacturing","hardware"],ecommerce:["ecommerce","consumer"],media:["media","consumer"],deeptech:["deeptech","hardware","manufacturing"],other:[]};

function tags(){return new Set(SECTOR_TAGS[S.sector]||[]);}
function geos(){return S.geos;}

function inc(q){
  const qt=q.tags||[];const qg=q.geo||[];
  return (qt.includes("all")||qt.some(t=>tags().has(t)))&&(!qg.length||qg.some(g=>geos().includes(g)));
}

function buildSheetRows(qs){
  const rows=[["Questions","Responses","Score"]];
  qs.forEach(q=>{
    if(!inc(q)) return;
    if(q.section){rows.push([q.section,null,null]);return;}
    if(q.geo&&q.geo.length) rows.push([`⚑ Applies to: ${q.geo.join(", ")} operations`,null,null]);
    rows.push([q.q,null,null]);
  });
  return rows;
}


const ENV_XLSX=[
  {section:"IR Checks",tags:["all"]},
  {q:"Will there be land acquisition while setting up an office or facility?",tags:["all"]},
  {q:"If yes — is the ownership status and current usage of the land to be acquired known?",tags:["all"]},
  {q:"Is there an existing Right of Way (ROW) to the company's office(s)?",tags:["all"]},
  {q:"Will the project require permanent and/or temporary involuntary land acquisition?",tags:["all"]},
  {q:"Will there be loss of crops, trees, or fixed assets due to land acquisition?",tags:["all"]},
  {q:"Will there be physical and/or economic displacement of any persons?",tags:["all"]},
  {q:"Will there be loss of income sources or livelihoods due to land acquisition?",tags:["all"]},
  {section:"IP Checks",tags:["all"]},
  {q:"Are there socio-cultural groups in/near the project area who may be considered indigenous communities?",tags:["all"]},
  {q:"Are there national/local laws or studies that classify these groups as indigenous communities?",tags:["all"]},
  {q:"Do such groups self-identify as a distinct social and cultural group?",tags:["all"]},
  {q:"Will the project directly/indirectly benefit or target indigenous peoples?",tags:["all"]},
  {q:"Will the project affect indigenous peoples' traditional socio-cultural practices?",tags:["all"]},
  {q:"Will the project involve physical displacement from traditional/customary lands?",tags:["all"]},
  {q:"Did the firm's establishment involve acquisition of lands traditionally owned by indigenous peoples?",tags:["all"]},
  {section:"Resource Efficiency and Pollution Prevention",tags:["all"]},
  {q:"Does the company generate any toxic air emissions?",tags:["all"]},
  {q:"Does the company generate any wastewater?",tags:["all"]},
  {q:"Does the company generate any solid waste?",tags:["all"]},
  {q:"Does the company generate any hazardous waste?",tags:["all"]},
  {q:"Does the company generate e-waste (electronic/electrical equipment)?",tags:["all"]},
  {q:"Does the company use hazardous chemicals, including pesticides?",tags:["manufacturing","agritech","hardware","deeptech"]},
  {q:"Does the company generate significant noise and vibration?",tags:["manufacturing","hardware","deeptech"]},
  {q:"Does the company affect water availability in the area of operations?",tags:["manufacturing","agritech"]},
  {q:"Does the company have mechanisms to monitor energy, water, raw material efficiency?",tags:["all"]},
  {q:"Do you encourage use of recycled or eco-certified materials in products or packaging?",tags:["manufacturing","hardware","ecommerce","deeptech"]},
  {section:"Environment Policies, Roles and Responsibilities",tags:["all"]},
  {q:"Does the Company have an Environment and Social (E&S) Policy?",tags:["all"]},
  {q:"Does the company have a designated person responsible for ESG / E&S mitigation post investment?",tags:["all"]},
  {q:"Does the company monitor its electricity consumption?",tags:["all"]},
  {q:"Does the company measure its Scope 1 and Scope 2 GHG emissions?",tags:["all"]},
  {q:"Does the company have a waste management policy or practice waste segregation?",tags:["all"]},
  {q:"Is the company's operation in an ecologically sensitive zone?",tags:["all"]},
  {q:"Is the company willing to update and send the ESG MIS every quarter?",tags:["all"]},
  {section:"Climate and Clean Energy (Sector-Specific)",tags:["climate","energy"]},
  {q:"Does the company track renewable energy generated or enabled (kWh/MW)?",tags:["climate","energy"]},
  {q:"Does the company track GHG emissions avoided or reduced (tCO2e/year)?",tags:["climate","energy"]},
  {section:"Agriculture and Land Use (Sector-Specific)",tags:["agritech"]},
  {q:"Does the company track pesticide/agrochemical usage by type and volume?",tags:["agritech"]},
  {q:"Does the company monitor land use change in its supply chain?",tags:["agritech"]},
];
const SOC_XLSX=[
  {section:"Diversity, Equity and Inclusion",tags:["all"]},
  {q:"Does the company adhere to anti-discrimination and equal opportunity practices?",tags:["all"]},
  {q:"Does the company have a formal anti-discrimination and equal opportunity policy?",tags:["all"]},
  {q:"If no — is the company willing to adopt anti-discrimination practices?",tags:["all"]},
  {q:"Does the company provide equal opportunity in recruitment, promotion, and development?",tags:["all"]},
  {q:"Does the company have a formal DEI (Diversity, Equity and Inclusion) policy?",tags:["all"]},
  {q:"What is the current gender breakdown of the workforce? (M / F / Other — provide numbers)",tags:["all"]},
  {q:"What percentage of senior management roles are held by women?",tags:["all"]},
  {q:"What percentage of the Board comprises women?",tags:["all"]},
  {section:"Community Relations and Stakeholder Engagement",tags:["all"]},
  {q:"Does the company have a grievance mechanism for external stakeholders?",tags:["all"]},
  {q:"Does the company have a GR mechanism/hotline with definitive timelines for grievance resolution?",tags:["all"]},
  {q:"Does the company engage with local communities or indigenous peoples for outreach activities?",tags:["all"]},
  {q:"Does the company work with external vendors / contract manufacturers?",tags:["all"]},
  {q:"Does the company have signed agreements with all external vendors?",tags:["all"]},
  {section:"Supply Chain and Vendor Management (Sector-Specific)",tags:["manufacturing","ecommerce","hardware","deeptech"]},
  {q:"Does the company have a diverse supplier base to mitigate supply chain risk?",tags:["manufacturing","ecommerce","hardware","deeptech"]},
  {q:"Does the company have a supplier code of conduct requiring statutory compliance?",tags:["manufacturing","ecommerce","hardware","deeptech"]},
  {q:"Do vendors/manufacturers have OHS policies and training programs in place?",tags:["manufacturing","ecommerce","hardware","deeptech"]},
  {q:"Do vendors have anti-child labour and anti-discrimination policies?",tags:["manufacturing","ecommerce","hardware","deeptech"]},
  {section:"Data Privacy and Cyber Security",tags:["all"]},
  {q:"Does the company have a formal cyber security policy?",tags:["all"]},
  {q:"Does the company have a Data Privacy Policy to protect customer data?",tags:["all"]},
  {q:"What types of customer data does the company store? (Personal / Contact / Financial / Transaction / Children's)",tags:["all"]},
  {q:"Which data protection laws apply? (DPDP Act 2023 / UK GDPR / CCPA / FTC Act — tick all applicable)",tags:["all"]},
  {q:"Does the company share customer data with third parties or vendors?",tags:["all"]},
  {q:"Does the company have a data retention and deletion policy?",tags:["all"]},
  {q:"Are there automated processes for data deletion?",tags:["all"]},
  {q:"Does the company provide a clear privacy notice to data subjects before collecting personal data?",tags:["all"]},
  {q:"Is all personal data encrypted in transit and at rest?",tags:["all"]},
  {q:"Are vendors and TPAs contractually required to maintain equivalent data protection standards?",tags:["all"]},
  {q:"Is there a system to track consent withdrawal and ensure processing stops immediately?",tags:["all"]},
  {q:"Is there a defined personal data breach response plan?",tags:["all"]},
  {q:"Are third-party vendors assessed for data security risks?",tags:["all"]},
  {section:"Children's Data and Safeguarding (Sector-Specific)",tags:["edtech","consumer","healthtech","media"]},
  {q:"Does the company collect or process data from users under 18 years of age?",tags:["edtech","consumer","healthtech","media"]},
  {q:"Does the company have a specific Child Data Protection policy?",tags:["edtech","consumer","healthtech","media"]},
  {q:"Does the company have a Child Protection Policy (covering abuse, exploitation, safeguarding)?",tags:["edtech","consumer","healthtech","media"]},
  {section:"Healthcare and Patient Safety (Sector-Specific)",tags:["healthtech"]},
  {q:"Are the company's diagnostic or therapeutic tools clinically validated?",tags:["healthtech"]},
  {q:"Are users informed that the tool does not replace professional medical advice?",tags:["healthtech"]},
  {q:"Does the company collect mental health or sensitive diagnostic information from users?",tags:["healthtech"]},
  {q:"Does the company have a protocol for escalation of high-risk users (e.g. suicidal ideation)?",tags:["healthtech"]},
  {section:"AI Ethics and Responsible Technology (Sector-Specific)",tags:["ai","saas","fintech","nbfc"]},
  {q:"Has the company conducted an AI risk assessment (e.g. EU AI Act risk classification)?",tags:["ai","saas","fintech","nbfc"]},
  {q:"Does the company's AI system use automated decision-making that materially affects individuals?",tags:["ai","fintech","nbfc","saas"]},
  {q:"Does the company have an AI ethics or responsible AI policy?",tags:["ai","saas"]},
  {q:"Are AI model outputs audited for bias (gender, caste, race, socioeconomic status)?",tags:["ai","fintech","nbfc"]},
  {section:"Financial Inclusion and Customer Protection (Sector-Specific)",tags:["fintech","nbfc","insurtech","wealthtech"]},
  {q:"Does the company have a customer protection and fair practices code?",tags:["fintech","nbfc","insurtech","wealthtech"]},
  {q:"Does the company have policies to prevent over-indebtedness of borrowers?",tags:["nbfc","fintech"]},
  {q:"Does the company have transparent pricing and fee disclosure to customers?",tags:["fintech","nbfc","insurtech","wealthtech"]},
  {q:"Does the company have a grievance mechanism for borrower/customer complaints per RBI guidelines?",tags:["fintech","nbfc"],geo:["India"]},
  {q:"Does the company comply with FCA consumer duty requirements?",tags:["fintech","nbfc","insurtech","wealthtech"],geo:["UK"]},
  {section:"Health, Safety and Accessibility",tags:["all"]},
  {q:"Does the company have a Health and Safety policy and training programme?",tags:["all"]},
  {q:"Does the property have accessibility amenities for differently-abled persons?",tags:["all"]},
  {q:"Did any employees face work-related injuries in the last year?",tags:["all"]},
  {q:"Does the facility have a first-aid kit on the premises?",tags:["all"]},
  {section:"Internal Grievance Redressal",tags:["all"]},
  {q:"Does the company have an internal grievance redressal (GR) mechanism for employees?",tags:["all"]},
  {q:"Are the contact details of GR committee members made available to all employees?",tags:["all"]},
  {section:"Employee Benefits and Well-Being",tags:["all"]},
  {q:"Has the company issued appointment letters with transparent working terms to all employees?",tags:["all"]},
  {q:"Does the company provide minimum wage to all employees per applicable law?",tags:["all"]},
  {q:"Does the company have a leave policy (annual, sick, parental)?",tags:["all"]},
  {q:"Are parental, sick and paid leaves granted to employees?",tags:["all"]},
  {q:"Does the company provide employees with PF/ESI as applicable under Indian law?",tags:["all"],geo:["India"]},
  {q:"Does the company comply with US payroll benefit requirements (Social Security, Medicare, health insurance)?",tags:["all"],geo:["US"]},
  {q:"Does the company comply with UK statutory employee benefit requirements (pension auto-enrolment, SSP, parental leave)?",tags:["all"],geo:["UK"]},
  {q:"Does the company provide health insurance to employees?",tags:["all"]},
  {q:"Does the company provide employees with skill development and training?",tags:["all"]},
  {q:"Are performance incentives, bonuses, or profit-sharing schemes available?",tags:["all"]},
  {q:"Are employee welfare metrics reported to leadership or the Board?",tags:["all"]},
  {q:"Does the company have a fixed salary payment date?",tags:["all"]},
  {q:"How many employees are above the age of 40?",tags:["all"]},
  {q:"Is the company planning to conduct free annual health check-ups per applicable labour law?",tags:["all"]},
  {section:"Gig and Contract Workers (Sector-Specific)",tags:["consumer","ecommerce","agritech","media"]},
  {q:"Does the company have gig workers or contractual field staff?",tags:["consumer","ecommerce","agritech","media"]},
  {q:"Does the company extend insurance coverage and safety training to gig workers?",tags:["consumer","ecommerce","agritech","media"]},
];
const GOV_XLSX=[
  {section:"Corporate Governance Structure",tags:["all"]},
  {q:"Does the company have a Board of Directors?",tags:["all"]},
  {q:"If no — does the company plan to constitute a Board with independent directors as it scales?",tags:["all"]},
  {q:"How many Board members are there?",tags:["all"]},
  {q:"Does the Board have oversight on ESG issues?",tags:["all"]},
  {q:"Has the Board conducted at least one meeting per quarter since inception?",tags:["all"]},
  {q:"Does the company have women on the Board?",tags:["all"]},
  {q:"If yes — what percentage of the Board comprises women?",tags:["all"]},
  {q:"Does the company have women in senior management?",tags:["all"]},
  {q:"If yes — what percentage of senior management comprises women?",tags:["all"]},
  {q:"Is the ownership structure simple and transparent?",tags:["all"]},
  {q:"Can the company clearly explain the source of funds for existing investors?",tags:["all"]},
  {section:"Policies and Compliance",tags:["all"]},
  {q:"Does the company have a Code of Conduct?",tags:["all"]},
  {q:"Does the company have an Anti-Bribery and Corruption (ABC) policy?",tags:["all"]},
  {q:"Does the company have an Anti-Money Laundering (AML) policy?",tags:["all"]},
  {q:"Does the company have a whistleblower policy and reporting hotline?",tags:["all"]},
  {q:"Does the company have an Equal Opportunity Policy?",tags:["all"]},
  {q:"Does the company have a POSH (Prevention of Sexual Harassment) policy?",tags:["all"],geo:["India"]},
  {q:"Has the company constituted an Internal Complaints Committee (ICC) as part of its POSH policy?",tags:["all"],geo:["India"]},
  {q:"Does the company have a formal sexual harassment / anti-harassment policy equivalent to POSH?",tags:["all"],geo:["US","UK"]},
  {q:"Does the company have a risk management policy?",tags:["all"]},
  {q:"Is the company registered under the Shops and Establishment Act?",tags:["all"],geo:["India"]},
  {q:"Does the company have a rental/lease agreement for its office space?",tags:["all"]},
  {q:"Have there been any regulatory notices, warnings, or penalties received?",tags:["all"]},
  {q:"Does the company maintain basic financial records (bank statements, invoices, contracts)?",tags:["all"]},
  {q:"Has the company experienced any internal fraud, theft, or misuse of funds?",tags:["all"]},
  {q:"Does the company have an internal audit system?",tags:["all"]},
  {q:"Are gifts, hospitality, or commissions used in sales or partnerships?",tags:["all"]},
  {section:"AML, KYC and Financial Integrity (Sector-Specific)",tags:["fintech","nbfc","insurtech","wealthtech"]},
  {q:"Does the company conduct KYC checks for all investors and customers?",tags:["fintech","nbfc","insurtech","wealthtech"]},
  {q:"Does the company have a dedicated AML compliance officer?",tags:["fintech","nbfc","insurtech","wealthtech"]},
  {q:"Does the company have a Suspicious Transaction Reporting (STR) mechanism?",tags:["fintech","nbfc","insurtech","wealthtech"]},
  {q:"Is the company compliant with FCA AML/KYC requirements?",tags:["fintech","nbfc","insurtech","wealthtech"],geo:["UK"]},
  {q:"Is the company compliant with FinCEN/BSA AML requirements?",tags:["fintech","nbfc","insurtech","wealthtech"],geo:["US"]},
  {section:"UK Regulatory Compliance (Sector-Specific)",tags:["all"],geo:["UK"]},
  {q:"Does the company comply with the UK Modern Slavery Act?",tags:["all"],geo:["UK"]},
  {q:"Does the company comply with UK GDPR and have a registered DPO if required?",tags:["all"],geo:["UK"]},
  {q:"Is the company registered with the ICO (Information Commissioner's Office)?",tags:["all"],geo:["UK"]},
  {section:"US Regulatory Compliance (Sector-Specific)",tags:["all"],geo:["US"]},
  {q:"Does the company comply with applicable US state employment laws (WARN Act, OSHA, FLSA)?",tags:["all"],geo:["US"]},
  {q:"Does the company have a documented CCPA / state data privacy compliance programme?",tags:["saas","consumer","ai","fintech","healthtech","edtech","media"],geo:["US"]},
  {section:"Supplier and Product Integrity (Sector-Specific)",tags:["manufacturing","ecommerce","hardware","deeptech"]},
  {q:"Does the company have a supplier code of conduct requiring statutory compliance from all suppliers?",tags:["manufacturing","ecommerce","hardware","deeptech"]},
  {q:"Do suppliers adhere to labour standards, environmental certifications, and safe chemical use?",tags:["manufacturing","ecommerce","hardware","deeptech"]},
  {q:"Has the company had any product recalls, complaints, or safety incidents?",tags:["manufacturing","hardware","deeptech","healthtech"]},
];

/* ── SMALL HELPERS ──────────────────────────────────────────────────── */
function $(s, r) { return (r || document).querySelector(s); }
function $$(s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); }
function escH(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function wholeNum(v) {
  if (v === '' || v == null) return null;
  var n = Number(v);
  return (isFinite(n) && n >= 0 && Math.floor(n) === n) ? n : null;
}

/* ── FORM CONSTANTS ─────────────────────────────────────────────────── */
var COUNTRIES = ['United States', 'United Kingdom', 'Singapore', 'United Arab Emirates',
  'Kenya', 'Nigeria', 'Canada', 'Germany', 'Australia', 'Indonesia'];
var CONFIRMS = [
  ['prohibition', 'Company does not fall under Blume’s Prohibition List (Annexure 1)'],
  ['debarred',    'Company is not on the Debarred Firms & Individuals list'],
  ['myanmar',     'Company has no operations or exposure in Myanmar or Afghanistan']
];
var MANDATORY = ['fund operations in a Developing Country',
  'create developmental benefits for a Developing Country'];
var SECONDARY = ['≥50% of employees are based in a Developing Country',
  '≥50% of income is taxable in a Developing Country',
  '≥50% of R&D activities are conducted in a Developing Country',
  '≥50% of business activities create developmental benefits (e.g. jobs) in a Developing Country'];

/* ── QUESTIONNAIRE BUILDER ──────────────────────────────────────────── */
function geoKeys(p) {
  var g = [];
  if (p.operatesIndia) g.push('India');
  var all = (p.exposureCountries || []).join(' , ');
  if (/\b(united states|usa|u\.s\.a?\.?|us)\b/i.test(all)) g.push('US');
  if (/\b(united kingdom|uk|u\.k\.|england|britain)\b/i.test(all)) g.push('UK');
  return g;
}
function geoText(p) {
  var parts = [];
  if (p.operatesIndia) parts.push('India');
  if (p.operatesOutside) parts.push('Outside India (' + (p.exposureCountries || []).join(', ') + ')');
  return parts.join('; ');
}
function buildRows(p) {
  S.sector = p.sector;
  S.geos = geoKeys(p);
  var gi = [
    ['Name of the Company', p.company],
    ['Company Description', null],
    ['Sector', SECTOR_NAMES[p.sector] || p.sector],
    ['Operating Geographies', geoText(p)],
    ['Is this sector on the Blume Exclusion List?', 'No'],
    ['Registered Address', null],
    ['Size of the organization', null],
    ['No. of Permanent employees', null],
    ['No. of Contractual employees', null],
    ['No. of Part-time employees', null],
    ['No. of Founders (Male / Female)', null],
    ['Inception Date of the Company', null],
    ['General Instructions', null],
    ['Please answer all questions. Do not leave any question blank.', null],
    ['For policies/documents — attach a copy or share the link where it is stored.', null],
    ['If something is not applicable, write NA and briefly explain why.', null],
    ['Information will be kept confidential — used only for ESG assessment.', null]
  ];
  var assess = [[null, null, 'Score', 'Categorisation'], [null, 'IR Categorisation', null, null],
    [null, 'IP Categorisation', null, null], [null, 'Company ESG Score', null, null],
    [null, 'Overall Risk Category', null, null]];
  return { gi: gi, env: buildSheetRows(ENV_XLSX), soc: buildSheetRows(SOC_XLSX),
    gov: buildSheetRows(GOV_XLSX), assess: assess, geos: S.geos.slice() };
}
function downloadXlsx(p) {
  if (!window.XLSX) { toast('Excel library did not load. Refresh and try again.', true); return; }
  var r = buildRows(p);
  var wb = XLSX.utils.book_new();
  var cols = [{ wch: 68 }, { wch: 38 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(r.gi), 'General Information');
  [['Environment', r.env], ['Social', r.soc], ['Governance', r.gov]].forEach(function (s) {
    var ws = XLSX.utils.aoa_to_sheet(s[1]); ws['!cols'] = cols;
    XLSX.utils.book_append_sheet(wb, ws, s[0]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(r.assess), 'Assessment');
  XLSX.writeFile(wb, p.company.replace(/[^\w\-]+/g, '_') + '_ESG_Questionnaire.xlsx');
}
function toast(msg, isErr) {
  if (typeof window.toast === 'function') { window.toast(msg, isErr); } else { alert(msg); }
}
window.EsgIC = { buildRows: buildRows };

/* ── STYLES ─────────────────────────────────────────────────────────── */
var CSS = `
.ic-tiles{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));grid-auto-rows:1fr;gap:20px;margin-top:8px}
.ic-tile{display:flex;flex-direction:column;text-align:left;font-family:var(--font-sans);color:var(--txt);background:var(--sur);border:1px solid var(--bor);border-radius:14px;padding:28px 30px;min-height:200px;cursor:pointer;transition:border-color .15s,box-shadow .15s}
.ic-tile:hover,.ic-tile:focus-visible{border-color:var(--accm);box-shadow:var(--card-hover-shadow);outline:none}
.ic-tile-top{display:flex;justify-content:space-between;align-items:flex-start}
.ic-tile svg{width:30px;height:30px;stroke:var(--acc);fill:none;stroke-width:1.6;stroke-linecap:round;stroke-linejoin:round}
.ic-tile .ic-arrow{width:20px;height:20px;stroke:var(--txtf)}
.ic-tile h2{font-size:1.3rem;font-weight:700;letter-spacing:-.02em;margin-top:18px}
.ic-tile p{font-size:.92rem;line-height:1.55;color:var(--txtm);margin-top:8px;max-width:440px}
.ic-tile .ic-foot{margin-top:auto;padding-top:14px;font-family:var(--font-mono);font-size:.68rem;color:var(--txtf)}
.ic-tile.hero{background:var(--acc2);border-color:var(--acc2);color:#fff}
.ic-tile.hero svg{stroke:#fff}
.ic-tile.hero p{color:rgba(255,255,255,.82)}
.ic-tile.hero .ic-foot{color:rgba(255,255,255,.7)}
.ic-tile.hero .ic-arrow{stroke:rgba(255,255,255,.8)}
@media(max-width:700px){.ic-tiles{grid-template-columns:1fr}}

.ic-form{max-width:780px}
.ic-card{background:var(--sur);border:1px solid var(--bor);border-radius:12px;padding:22px 24px;margin-bottom:16px}
.ic-card h3{font-size:1rem;font-weight:700;margin-bottom:4px}
.ic-hint{font-size:.78rem;color:var(--txtm);margin-bottom:16px;line-height:1.55}
.ic-f{margin-bottom:16px}
.ic-f:last-child{margin-bottom:0}
.ic-lbl{display:block;font-family:var(--font-mono);font-size:.66rem;color:var(--txtm);text-transform:uppercase;letter-spacing:.08em;font-weight:600;margin-bottom:7px}
.ic-req{color:var(--red)}
.ic-f input[type=text],.ic-f input[type=number],.ic-f select,.ic-f textarea{width:100%;padding:10px 14px;background:var(--sur2);border:1px solid var(--bor);border-radius:8px;color:var(--txt);font-family:var(--font-mono);font-size:.8rem}
.ic-f textarea{font-family:var(--font-sans);min-height:96px;resize:vertical;line-height:1.6}
.ic-f input:focus,.ic-f select:focus,.ic-f textarea:focus{outline:none;border-color:var(--acc);box-shadow:0 0 0 3px var(--accd)}
.ic-bad{border-color:var(--red)!important}
.ic-r2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.ic-r3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
@media(max-width:600px){.ic-r2,.ic-r3{grid-template-columns:1fr}}
.ic-check{display:flex;gap:10px;align-items:flex-start;padding:8px 0;font-size:.88rem;line-height:1.5;cursor:pointer}
.ic-check input{margin-top:3px;accent-color:var(--acc);width:16px;height:16px;flex-shrink:0;cursor:pointer}
.ic-sub{margin:4px 0 0 0;padding:14px 16px;background:var(--sur2);border:1px solid var(--bor);border-radius:10px}
.ic-sub .ic-lbl{margin-top:12px}
.ic-sub .ic-lbl:first-child{margin-top:0}
.ic-err{display:none;background:var(--redd);border:1px solid var(--red);color:var(--red);border-radius:10px;padding:14px 18px;font-size:.84rem;margin-bottom:16px;line-height:1.6}
.ic-err ul{margin:6px 0 0 18px}
.ic-note{background:var(--warnd);border:1px solid var(--warn);color:var(--warn);border-radius:8px;padding:10px 14px;font-size:.8rem;line-height:1.55;margin-top:10px}
.ic-ok{background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.3);color:var(--green);border-radius:8px;padding:10px 14px;font-size:.82rem;line-height:1.55}
.ic-warn{background:var(--warnd);border:1px solid var(--warn);color:var(--txt);border-radius:8px;padding:12px 14px;font-size:.82rem;line-height:1.6}
.ic-warn .btn{margin-top:8px}
.ic-foot-bar{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;padding-top:4px}
.ic-who{font-family:var(--font-mono);font-size:.7rem;color:var(--txtm)}
.ic-meta{margin:14px 0;padding:12px 14px;background:var(--sur2);border:1px solid var(--bor);border-radius:10px;font-size:.82rem;line-height:1.7;color:var(--txtm)}
.ic-meta b{color:var(--txt);font-weight:600}
`;

/* ── LANDING PAGE ───────────────────────────────────────────────────── */
var ICONS = {
  esap: '<path d="M9.615 20h-2.615a2 2 0 0 1 -2 -2v-12a2 2 0 0 1 2 -2h8a2 2 0 0 1 2 2v8"/><path d="M14 19l2 2l4 -4"/><path d="M9 8h4"/><path d="M9 12h2"/>',
  tools: '<path d="M7 10h3v-3l-3.5 -3.5a6 6 0 0 1 8 8l6 6a2 2 0 0 1 -3 3l-6 -6a6 6 0 0 1 -8 -8l3.5 3.5"/>',
  esg:   '<path d="M5 21c.5 -4.5 2.5 -8 7 -10"/><path d="M9 18c6.218 0 10.5 -3.288 11 -12v-2h-4.014c-9 0 -11.986 4 -12 9c0 1 0 3 2 5h3z"/>',
  form:  '<path d="M9 5h-2a2 2 0 0 0 -2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2 -2v-12a2 2 0 0 0 -2 -2h-2"/><path d="M9 3m0 2a2 2 0 0 1 2 -2h2a2 2 0 0 1 2 2v0a2 2 0 0 1 -2 2h-2a2 2 0 0 1 -2 -2z"/><path d="M9 14l2 2l4 -4"/>',
  arrow: '<path d="M5 12l14 0"/><path d="M13 18l6 -6"/><path d="M13 6l6 6"/>'
};
var TILES = [
  { go: 'fund4',  icon: 'esap',  title: 'ESAP Status', text: 'Follow action items and progress company by company.', foot: 'Fund IV and Fund V' },
  { go: 'analytics', icon: 'tools', title: 'Tools', text: 'Analytics, filters and exports across both funds.', foot: 'Analytics' },
  { go: 'esg',    icon: 'esg',   title: 'ESG', text: 'IC requirements, annexures, team discussions and templates.', foot: 'ESG for Blume, Discussions, Templates' },
  { go: 'icform', icon: 'form',  title: 'ESG IC Form', text: 'Submit the pre-IC checklist and get the ESG questionnaire for the promoter.', foot: 'Complete before IC 1', hero: true }
];
function svg(name) { return '<svg viewBox="0 0 24 24" aria-hidden="true">' + ICONS[name] + '</svg>'; }
function homeHtml() {
  return '<div class="page-header"><div><div class="page-title">What would you like to open?</div>' +
    '<div class="page-sub">Pick a section to get started</div></div></div><div class="ic-tiles">' +
    TILES.map(function (t) {
      return '<button type="button" class="ic-tile' + (t.hero ? ' hero' : '') + '" data-go="' + t.go + '">' +
        '<div class="ic-tile-top">' + svg(t.icon) + '<svg class="ic-arrow" viewBox="0 0 24 24" aria-hidden="true">' + ICONS.arrow + '</svg></div>' +
        '<h2>' + escH(t.title) + '</h2><p>' + escH(t.text) + '</p><div class="ic-foot">' + escH(t.foot) + '</div></button>';
    }).join('') + '</div>';
}

/* ── ESG IC FORM ────────────────────────────────────────────────────── */
function checkRow(cls, label, attrs) {
  return '<label class="ic-check"><input type="checkbox" class="' + cls + '" ' + (attrs || '') + '><span>' + escH(label) + '</span></label>';
}
function formHtml() {
  var sectorOpts = '<option value="">Select sector</option>' + Object.keys(SECTOR_NAMES).map(function (k) {
    return '<option value="' + k + '">' + escH(SECTOR_NAMES[k]) + '</option>';
  }).join('');
  return '<div class="page-header"><div><div class="page-title">ESG IC Form</div>' +
    '<div class="page-sub">Complete before the Company goes to IC 1</div></div></div>' +
    '<div class="ic-form" id="icForm">' +
    '<div class="ic-err" id="icErr" role="alert"></div>' +

    '<div class="ic-card"><h3>Company</h3><div class="ic-hint">The sector and geography decide which questions go into the ESG questionnaire.</div>' +
    '<div class="ic-f"><label class="ic-lbl" for="icName">Company name <span class="ic-req">*</span></label><input type="text" id="icName" placeholder="e.g. Acme Technologies Pvt Ltd" autocomplete="off"></div>' +
    '<div class="ic-f"><label class="ic-lbl" for="icSector">Sector <span class="ic-req">*</span></label><select id="icSector">' + sectorOpts + '</select></div>' +
    '<div class="ic-f"><label class="ic-lbl" for="icDesc">Brief business description <span class="ic-req">*</span></label><textarea id="icDesc" placeholder="What the Company does and who it serves"></textarea></div>' +
    '<div class="ic-f"><label class="ic-lbl" for="icSite">Website</label><input type="text" id="icSite" placeholder="https://" autocomplete="off"></div>' +
    '<div class="ic-r2"><div class="ic-f"><label class="ic-lbl" for="icMale">Male founders <span class="ic-req">*</span></label><input type="number" min="0" step="1" id="icMale"></div>' +
    '<div class="ic-f"><label class="ic-lbl" for="icFemale">Female founders <span class="ic-req">*</span></label><input type="number" min="0" step="1" id="icFemale"></div></div></div>' +

    '<div class="ic-card"><h3>Where does the Company operate? <span class="ic-req">*</span></h3><div class="ic-hint">Tick both if it operates in India and elsewhere.</div>' +
    checkRow('ic-geo', 'India', 'data-g="india"') + checkRow('ic-geo', 'Outside India', 'data-g="outside"') +
    '<div class="ic-sub" id="icOutside" style="display:none"><span class="ic-lbl">Exposure countries <span class="ic-req">*</span></span>' +
    '<div class="chip-grid" id="icCountries">' + COUNTRIES.map(function (c) {
      return '<div class="chip-option" role="button" tabindex="0" data-c="' + escH(c) + '">' + escH(c) + '</div>';
    }).join('') + '</div>' +
    '<span class="ic-lbl">Other countries</span><div class="ic-f"><input type="text" id="icOther" placeholder="Comma-separated, e.g. Ghana, Vietnam" autocomplete="off"></div></div></div>' +

    '<div class="ic-card"><h3>Eligibility confirmations <span class="ic-req">*</span></h3><div class="ic-hint">All three are required. Annexure 1 is under ESG for Blume.</div>' +
    CONFIRMS.map(function (c) { return checkRow('ic-conf', c[1], 'data-k="' + c[0] + '"'); }).join('') + '</div>' +

    '<div class="ic-card"><h3>Developing Country requirement</h3><div class="ic-hint">Is the Company operating in a Developing Country (Annexure 2)? <span class="ic-req">*</span></div>' +
    '<label class="ic-check"><input type="radio" name="icDev" value="Yes"><span>Yes</span></label>' +
    '<label class="ic-check"><input type="radio" name="icDev" value="No"><span>No</span></label>' +
    '<div class="ic-note" id="icDevNo" style="display:none">Blume invests only in Companies operating in a Developing Country. If you submit, this response is flagged for review.</div>' +
    '<div class="ic-sub" id="icDevYes" style="display:none">' +
    '<span class="ic-lbl">Mandatory condition: ≥50% of investment proceeds from Blume will be used to (at least one) <span class="ic-req">*</span></span>' +
    MANDATORY.map(function (m) { return checkRow('ic-mand', m, 'data-l="' + escH(m) + '"'); }).join('') +
    '<span class="ic-lbl">Secondary condition: at least one must be true <span class="ic-req">*</span></span>' +
    SECONDARY.map(function (m) { return checkRow('ic-sec', m, 'data-l="' + escH(m) + '"'); }).join('') + '</div></div>' +

    '<div class="ic-card"><h3>SDG alignment</h3><div class="ic-hint">One paragraph on how this investment aligns with the SDGs. A sample is in Annexure 3 under ESG for Blume.</div>' +
    '<div class="ic-f"><label class="ic-lbl" for="icSdg">SDG alignment paragraph <span class="ic-req">*</span></label><textarea id="icSdg" style="min-height:140px"></textarea></div></div>' +

    '<div class="ic-card"><h3>Employees</h3><div class="ic-hint">Enter 0 where there are none.</div><div class="ic-r3">' +
    '<div class="ic-f"><label class="ic-lbl" for="icPerm">Permanent <span class="ic-req">*</span></label><input type="number" min="0" step="1" id="icPerm"></div>' +
    '<div class="ic-f"><label class="ic-lbl" for="icPart">Part-time <span class="ic-req">*</span></label><input type="number" min="0" step="1" id="icPart"></div>' +
    '<div class="ic-f"><label class="ic-lbl" for="icCont">Contractual <span class="ic-req">*</span></label><input type="number" min="0" step="1" id="icCont"></div></div></div>' +

    '<div class="ic-foot-bar"><div class="ic-who" id="icWho"></div>' +
    '<button type="button" class="btn primary" id="icSubmit">Submit and get questionnaire</button></div></div>';
}

function collect() {
  var f = $('#icForm');
  var chips = $$('#icCountries .chip-option.selected', f).map(function (c) { return c.getAttribute('data-c'); });
  var other = ($('#icOther', f).value || '').split(/[,;\n]+/).map(function (s) { return s.trim(); }).filter(Boolean);
  var dev = $('input[name=icDev]:checked', f);
  var conf = {}; $$('.ic-conf', f).forEach(function (c) { conf[c.getAttribute('data-k')] = c.checked; });
  var inIndia = $('.ic-geo[data-g=india]', f).checked, outside = $('.ic-geo[data-g=outside]', f).checked;
  var sel = $('#icSector', f).value;
  return {
    company: $('#icName', f).value.trim(), sector: sel,
    sectorLabel: SECTOR_NAMES[sel] || '', description: $('#icDesc', f).value.trim(),
    website: $('#icSite', f).value.trim(),
    maleFounders: wholeNum($('#icMale', f).value), femaleFounders: wholeNum($('#icFemale', f).value),
    operatesIndia: inIndia, operatesOutside: outside,
    exposureCountries: outside ? chips.concat(other) : [],
    confirmations: conf, developingCountry: dev ? dev.value : '',
    mandatory: $$('.ic-mand', f).filter(function (c) { return c.checked; }).map(function (c) { return c.getAttribute('data-l'); }),
    secondary: $$('.ic-sec', f).filter(function (c) { return c.checked; }).map(function (c) { return c.getAttribute('data-l'); }),
    sdg: $('#icSdg', f).value.trim(),
    permanent: wholeNum($('#icPerm', f).value), partTime: wholeNum($('#icPart', f).value),
    contractual: wholeNum($('#icCont', f).value)
  };
}
function validate(p) {
  var msgs = [], bad = [];
  function need(cond, msg, id) { if (!cond) { msgs.push(msg); if (id) bad.push(id); } }
  need(p.company, 'Enter the Company name.', 'icName');
  need(p.sector, 'Pick a sector.', 'icSector');
  need(p.description, 'Add a brief business description.', 'icDesc');
  need(p.maleFounders !== null && p.femaleFounders !== null, 'Enter the number of male and female founders (0 is fine).', null);
  if (p.maleFounders === null) bad.push('icMale'); if (p.femaleFounders === null) bad.push('icFemale');
  need(p.operatesIndia || p.operatesOutside, 'Tick where the Company operates: India, Outside India, or both.');
  if (p.operatesOutside) need(p.exposureCountries.length > 0, 'Pick at least one exposure country, or add one under Other countries.');
  need(CONFIRMS.every(function (c) { return p.confirmations[c[0]]; }),
    'Tick all three eligibility confirmations. If one doesn’t apply, contact ' + REVIEW_MAIL + '.');
  need(p.developingCountry, 'Say whether the Company operates in a Developing Country.');
  if (p.developingCountry === 'Yes') {
    need(p.mandatory.length > 0, 'Pick at least one mandatory condition.');
    need(p.secondary.length > 0, 'Pick at least one secondary condition.');
  }
  need(p.sdg, 'Add the SDG alignment paragraph.', 'icSdg');
  need(p.permanent !== null && p.partTime !== null && p.contractual !== null, 'Enter permanent, part-time and contractual headcount (0 is fine).');
  if (p.permanent === null) bad.push('icPerm'); if (p.partTime === null) bad.push('icPart'); if (p.contractual === null) bad.push('icCont');
  return { msgs: msgs, bad: bad };
}
function showErrors(v) {
  $$('#icForm .ic-bad').forEach(function (e) { e.classList.remove('ic-bad'); });
  var box = $('#icErr');
  if (!v.msgs.length) { box.style.display = 'none'; box.innerHTML = ''; return; }
  v.bad.forEach(function (id) { var el = document.getElementById(id); if (el) el.classList.add('ic-bad'); });
  box.innerHTML = '<strong>Fix these before submitting</strong><ul>' + v.msgs.map(function (m) { return '<li>' + escH(m) + '</li>'; }).join('') + '</ul>';
  box.style.display = 'block';
  box.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

/* ── SAVING TO THE SHEET ────────────────────────────────────────────── */
function saveResponse(p) {
  if (!ENDPOINT) return Promise.resolve({ ok: false, reason: 'not-configured' });
  var user = window.firebase && firebase.auth().currentUser;
  if (!user) return Promise.resolve({ ok: false, reason: 'not-signed-in' });
  return user.getIdToken().then(function (tok) {
    return fetch(ENDPOINT, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ idToken: tok, data: p }) });
  }).then(function (r) { return r.json(); }).then(function (j) {
    return (j && j.ok) ? { ok: true, row: j.row, repeat: !!j.repeat } : { ok: false, reason: (j && j.error) || 'unknown error' };
  }).catch(function (e) { return { ok: false, reason: (e && e.message) || 'network error' }; });
}

/* ── POP-UP ─────────────────────────────────────────────────────────── */
var lastPayload = null, lastSave = null;
function doneHtml() {
  return '<div class="modal" style="max-width:520px"><div class="modal-hd"><div class="modal-title">ESG questionnaire ready</div>' +
    '<div class="modal-close" id="icDoneX" role="button" aria-label="Close">×</div></div>' +
    '<div class="modal-body"><p style="font-size:.95rem;line-height:1.7">If this deal goes to IC 2, ask the promoter to fill this questionnaire.</p>' +
    '<div class="ic-meta" id="icDoneMeta"></div><div id="icSaveState"></div></div>' +
    '<div class="modal-footer"><button type="button" class="btn secondary" id="icDoneClose">Close</button>' +
    '<button type="button" class="btn primary" id="icDoneDl">Download questionnaire (.xlsx)</button></div></div>';
}
function reasonText(r) {
  if (r === 'not-configured') return 'The sheet connection isn’t set up yet.';
  return 'Couldn’t save to the sheet (' + r + ').';
}
function renderSaveState() {
  var el = $('#icSaveState');
  if (lastSave && lastSave.ok) {
    el.innerHTML = '<div class="ic-ok">Response saved to the ESG Checklist sheet.' +
      (lastSave.repeat ? ' This Company already had a response, so this one was added as a new row.' : '') + '</div>';
    return;
  }
  el.innerHTML = '<div class="ic-warn"><strong>Response not saved.</strong> ' + escH(reasonText(lastSave ? lastSave.reason : '')) +
    ' You can still download the questionnaire. Retry, and tell ' + escH(REVIEW_MAIL) + ' if it keeps failing.' +
    '<br><button type="button" class="btn" id="icRetry">Retry saving</button></div>';
  $('#icRetry').onclick = function () {
    this.disabled = true; this.textContent = 'Saving…';
    saveResponse(lastPayload).then(function (r) { lastSave = r; renderSaveState(); if (r.ok) resetForm(); });
  };
}
function openDone() {
  var p = lastPayload, r = buildRows(p);
  var inc = r.geos.length ? 'Includes questions specific to ' + r.geos.join(' and ') + '.' : 'Core questions only.';
  $('#icDoneMeta').innerHTML = '<b>' + escH(p.company) + '</b><br>Sector: ' + escH(p.sectorLabel) +
    '<br>Geography: ' + escH(geoText(p)) + '<br>' + escH(inc);
  renderSaveState();
  $('#icDone').classList.add('on');
}
function closeDone() {
  $('#icDone').classList.remove('on');
  if (lastSave && lastSave.ok) window.switchView('home');
}

/* ── FORM LIFECYCLE ─────────────────────────────────────────────────── */
function who() {
  var u = window.firebase && firebase.auth().currentUser;
  return u && u.email ? 'Submitting as ' + u.email : '';
}
function resetForm() {
  $('#vIcForm').innerHTML = formHtml();
  wireForm();
}
function wireForm() {
  var f = $('#icForm');
  $('#icWho').textContent = who();
  f.addEventListener('change', function (e) {
    var t = e.target;
    if (t.classList.contains('ic-geo')) $('#icOutside').style.display = $('.ic-geo[data-g=outside]', f).checked ? 'block' : 'none';
    if (t.name === 'icDev') {
      $('#icDevYes').style.display = t.value === 'Yes' ? 'block' : 'none';
      $('#icDevNo').style.display = t.value === 'No' ? 'block' : 'none';
    }
  });
  function toggleChip(c) { c.classList.toggle('selected'); }
  f.addEventListener('click', function (e) { var c = e.target.closest('.chip-option'); if (c) toggleChip(c); });
  f.addEventListener('keydown', function (e) {
    var c = e.target.closest && e.target.closest('.chip-option');
    if (c && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); toggleChip(c); }
  });
  $('#icSubmit').addEventListener('click', function () {
    var p = collect(), v = validate(p);
    showErrors(v);
    if (v.msgs.length) return;
    var btn = this; btn.disabled = true; btn.textContent = 'Submitting…';
    saveResponse(p).then(function (r) {
      lastPayload = p; lastSave = r;
      btn.disabled = false; btn.textContent = 'Submit and get questionnaire';
      openDone();
      if (r.ok) resetForm();
    });
  });
}

/* ── NAV, VIEWS AND WIRING INTO THE TRACKER ─────────────────────────── */
var TITLES = { home: 'Home', icform: 'ESG IC Form' };
var NAVS = { home: 'navHome', icform: 'navIcForm' };
var VIEWS = { home: 'vHome', icform: 'vIcForm' };

function showMine(view) {
  window.currentView = view;
  $$('.nav-item').forEach(function (n) { n.classList.remove('active'); });
  var nav = document.getElementById(NAVS[view]); if (nav) nav.classList.add('active');
  var t = document.getElementById('topbarTitle'); if (t) t.textContent = TITLES[view];
  $$('.view').forEach(function (v) { v.classList.remove('on'); });
  document.getElementById(VIEWS[view]).classList.add('on');
  if (view === 'icform') $('#icWho').textContent = who();
  window.scrollTo(0, 0);
}
function injectNav() {
  var nav = $('.sidebar-nav'); if (!nav) return;
  $$('.nav-section-label', nav).forEach(function (l) {
    if (l.textContent.trim() === 'Funds') l.textContent = 'ESAP Status';
  });
  var home = document.createElement('div');
  home.className = 'nav-item'; home.id = 'navHome';
  home.innerHTML = '<span class="nav-icon">🏠</span><span>Home</span>';
  home.addEventListener('click', function () { window.switchView('home'); });
  nav.insertBefore(home, nav.firstChild);
  var form = document.createElement('div');
  form.className = 'nav-item'; form.id = 'navIcForm';
  form.innerHTML = '<span class="nav-icon">📝</span><span>ESG IC Form</span>';
  form.addEventListener('click', function () { window.switchView('icform'); });
  var tpl = document.getElementById('navTemplates');
  if (tpl && tpl.parentNode) tpl.parentNode.insertBefore(form, tpl.nextSibling); else nav.appendChild(form);
}
function injectViews() {
  var host = document.getElementById('vFund4');
  host = host ? host.parentNode : $('.main');
  var home = document.createElement('div'); home.className = 'view'; home.id = 'vHome'; home.innerHTML = homeHtml();
  var form = document.createElement('div'); form.className = 'view'; form.id = 'vIcForm'; form.innerHTML = formHtml();
  host.appendChild(home); host.appendChild(form);
  $$('.ic-tile', home).forEach(function (b) {
    b.addEventListener('click', function () { window.switchView(b.getAttribute('data-go')); });
  });
  var done = document.createElement('div'); done.className = 'modal-overlay'; done.id = 'icDone'; done.innerHTML = doneHtml();
  document.body.appendChild(done);
  $('#icDoneX').onclick = closeDone; $('#icDoneClose').onclick = closeDone;
  $('#icDoneDl').onclick = function () { if (lastPayload) downloadXlsx(lastPayload); };
  wireForm();
}
function wrapSwitch() {
  var orig = window.switchView; if (typeof orig !== 'function') return false;
  window.switchView = function (view) {
    if (view === 'home' || view === 'icform') return showMine(view);
    return orig.apply(this, arguments);
  };
  return true;
}
function patchChecklistLink() {
  $$('#esgTabIC a').forEach(function (a) {
    if ((a.getAttribute('href') || '').indexOf('docs.google.com/forms') === -1) return;
    a.setAttribute('href', '#'); a.removeAttribute('target');
    a.addEventListener('click', function (e) { e.preventDefault(); window.switchView('icform'); });
  });
}
function init() {
  var st = document.createElement('style'); st.id = 'icStyles'; st.textContent = CSS; document.head.appendChild(st);
  if (!wrapSwitch()) { console.warn('esg-ic: switchView not found, add-on not started'); return; }
  injectNav(); injectViews(); patchChecklistLink();
  showMine('home');
  try {
    firebase.auth().onAuthStateChanged(function (u) {
      if (u && /@blume\.vc$/i.test(u.email || '')) showMine('home');
    });
  } catch (e) { /* tracker handles sign-in itself */ }
}
if (document.readyState === 'loading') window.addEventListener('DOMContentLoaded', init); else init();

})();
