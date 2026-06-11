const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageBreak, LevelFormat,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ── Palette ──────────────────────────────────────────────────
const GREEN  = "1B6B3A"; const LGREEN = "D6EAD9";
const GOLD   = "B7770D"; const LGOLD  = "FEF9E7";
const BLUE   = "1A4E7C"; const LBLUE  = "D6EAF8";
const RED    = "C0392B"; const LRED   = "FADBD8";
const PURPLE = "5B2C6F"; const LPURP  = "E8DAEF";
const TEAL   = "0E6655"; const LTEAL  = "D1F2EB";
const ORANGE = "A04000"; const LORANG = "FDEBD0";
const GREY   = "F4F6F7"; const DGREY  = "444444"; const WHITE = "FFFFFF";
const FULL_W = 9360;

// ── Borders ──────────────────────────────────────────────────
const bdr  = (c="BBBBBB",s=4) => ({style:BorderStyle.SINGLE,size:s,color:c});
const allB = (c="CCCCCC") => ({top:bdr(c),bottom:bdr(c),left:bdr(c),right:bdr(c)});

// ── Para helpers ─────────────────────────────────────────────
const h1 = (text,bg=GREEN) => new Paragraph({
  heading:HeadingLevel.HEADING_1,
  children:[new TextRun({text,font:"Arial",size:34,bold:true,color:WHITE})],
  shading:{fill:bg,type:ShadingType.CLEAR},
  spacing:{before:440,after:200},indent:{left:280,right:280},
});
const h2 = (text,c=GREEN) => new Paragraph({
  heading:HeadingLevel.HEADING_2,
  children:[new TextRun({text,font:"Arial",size:27,bold:true,color:c})],
  spacing:{before:300,after:100},
  border:{bottom:{style:BorderStyle.SINGLE,size:6,color:c,space:1}},
});
const h3 = (text,c=GOLD) => new Paragraph({
  heading:HeadingLevel.HEADING_3,
  children:[new TextRun({text,font:"Arial",size:23,bold:true,color:c})],
  spacing:{before:200,after:80},
});
const para = (text,bold=false,color=DGREY) => new Paragraph({
  children:[new TextRun({text,font:"Arial",size:22,color,bold})],
  spacing:{before:60,after:80},
});
const bullet = (text,bold=false) => new Paragraph({
  numbering:{reference:"bullets",level:0},
  children:[new TextRun({text,font:"Arial",size:22,color:DGREY,bold})],
  spacing:{before:40,after:40},
});
const spacer = (n=1) => new Paragraph({children:[new TextRun("")],spacing:{before:0,after:n*80}});
const pbr = () => new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}});

const callout = (text,fill=LGOLD,bc=GOLD,bold=false) => new Table({
  width:{size:FULL_W,type:WidthType.DXA},columnWidths:[FULL_W],
  rows:[new TableRow({children:[new TableCell({
    width:{size:FULL_W,type:WidthType.DXA},
    shading:{fill,type:ShadingType.CLEAR},
    borders:{top:bdr(bc,8),bottom:bdr(bc,8),left:bdr(bc,16),right:bdr(bc,4)},
    margins:{top:120,bottom:120,left:200,right:200},
    children:[new Paragraph({children:[new TextRun({text,font:"Arial",size:22,color:DGREY,bold})]})],
  })]})],
});

const hC = (text,w,fill=GREEN,color=WHITE) => new TableCell({
  width:{size:w,type:WidthType.DXA},shading:{fill,type:ShadingType.CLEAR},
  borders:allB("999999"),margins:{top:100,bottom:100,left:140,right:140},
  verticalAlign:VerticalAlign.CENTER,
  children:[new Paragraph({alignment:AlignmentType.CENTER,
    children:[new TextRun({text,font:"Arial",size:20,bold:true,color})]})],
});
const dC = (text,w,fill=WHITE,align=AlignmentType.LEFT,bold=false,color=DGREY) => new TableCell({
  width:{size:w,type:WidthType.DXA},shading:{fill,type:ShadingType.CLEAR},
  borders:allB("CCCCCC"),margins:{top:80,bottom:80,left:140,right:140},
  verticalAlign:VerticalAlign.CENTER,
  children:[new Paragraph({alignment:align,
    children:[new TextRun({text,font:"Arial",size:20,color,bold})]})],
});

// ═══════════════════════════════════════════════════
// COVER
// ═══════════════════════════════════════════════════
function cover(){ return [
  spacer(3),
  new Paragraph({alignment:AlignmentType.CENTER,
    children:[new TextRun({text:"🐔",font:"Segoe UI Emoji",size:96})],
    spacing:{before:0,after:120}}),
  new Paragraph({alignment:AlignmentType.CENTER,
    children:[new TextRun({text:"LAYER POULTRY FARMING",font:"Arial",size:54,bold:true,color:GREEN})],
    spacing:{before:0,after:60}}),
  new Paragraph({alignment:AlignmentType.CENTER,
    children:[new TextRun({text:"Comprehensive Project Report",font:"Arial",size:38,bold:true,color:GOLD})],
    spacing:{before:0,after:80}}),
  new Paragraph({alignment:AlignmentType.CENTER,
    children:[new TextRun({text:"Expanded Edition | 500 Birds | Kerala Conditions",font:"Arial",size:26,color:DGREY,italics:true})],
    spacing:{before:0,after:320}}),
  new Table({width:{size:FULL_W,type:WidthType.DXA},columnWidths:[FULL_W],
    rows:[new TableRow({children:[new TableCell({
      width:{size:FULL_W,type:WidthType.DXA},
      shading:{fill:LGREEN,type:ShadingType.CLEAR},
      borders:{top:bdr(GREEN,10),bottom:bdr(GREEN,10),left:bdr(GREEN,10),right:bdr(GREEN,10)},
      margins:{top:200,bottom:200,left:360,right:360},
      children:[
        new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"19 SECTIONS | FULL BUSINESS PLAN + TECHNICAL GUIDE",font:"Arial",size:24,bold:true,color:GREEN})],spacing:{before:0,after:120}}),
        new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Site Selection  •  Infrastructure  •  Breed & Flock Management  •  Feed & Nutrition",font:"Arial",size:22,color:DGREY})],spacing:{before:0,after:60}}),
        new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Water & Waste  •  Disease Control  •  Financial Projections  •  Month-by-Month Cash Flow",font:"Arial",size:22,color:DGREY})],spacing:{before:0,after:60}}),
        new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Marketing  •  NABARD Application  •  Supplier Directory  •  Risk Management",font:"Arial",size:22,color:DGREY})],spacing:{before:0,after:0}}),
      ],
    })]})]}),
  spacer(4),
  new Paragraph({alignment:AlignmentType.CENTER,
    children:[new TextRun({text:`${new Date().toLocaleDateString('en-IN',{year:'numeric',month:'long'})}  |  Kerala Layer Farm Series — Expanded Edition`,font:"Arial",size:21,color:DGREY})]}),
  pbr(),
]}

// ═══════════════════════════════════════════════════
// SEC 1: EXECUTIVE SUMMARY
// ═══════════════════════════════════════════════════
function execSummary(){ return [
  h1("1. EXECUTIVE SUMMARY"),
  spacer(1),
  para("This report presents a detailed feasibility study and implementation plan for establishing a 500-bird layer poultry farm in Kerala. The project targets a net profit of ₹9,000–25,000 per month in Year 1, growing to ₹30,000–45,000 per month by Year 2 as buyer networks mature and channel mix improves. The project is NABARD-fundable with 25–33% capital subsidy available to eligible applicants."),
  spacer(1),
  h2("Project KPI Snapshot"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2400,2400,2400,2160],
    rows:[
      new TableRow({children:[hC("Parameter",2400),hC("Value",2400),hC("Unit",2400),hC("Remarks",2160)]}),
      ...([
        ["Flock Size","500","BV-380 layer hens","Brown egg commercial breed"],
        ["Bird Age at Housing","18 weeks","Pullets (ready-to-lay)","Avoids 18-week rearing cost"],
        ["Production Cycle","18 months","Weeks 18–90","72 weeks of laying"],
        ["Peak Lay Rate","90–92%","Weeks 26–50","460 eggs/day at peak"],
        ["Average Lay Rate","76–80%","Full cycle average","Conservative estimate"],
        ["Daily Egg Output (avg)","380–400","Eggs/day","At 76–80% lay rate"],
        ["Monthly Eggs (avg)","11,400–12,000","Eggs/month",""],
        ["Total Project Cost","₹8,85,000","One-time capital","See Section 14"],
        ["NABARD Subsidy","₹1,47,500–2,21,250","25–33% back-ended","Applied via bank"],
        ["Net Capital After Subsidy","₹6,63,750–7,37,500","Effective investment",""],
        ["Monthly Feed Cost","₹51,100","At ₹28/kg feed","Largest operating cost"],
        ["Total Monthly Cost","₹70,517","All-in operating","See Section 14"],
        ["Monthly Revenue (Yr 1)","₹78,125–99,450","₹5.50–7.00/egg avg","Depends on channel"],
        ["Net Profit (Yr 1)","₹7,608–28,933","Per month","After all costs"],
        ["Net Profit (Yr 2+)","₹30,000–45,000","Per month","With premium channels"],
        ["Investment Payback","18–24 months","Conservative",""],
        ["Annual ROI (Yr 3)","35–55%","On capital invested",""],
      ]).map(([p,v,u,r],i)=>new TableRow({children:[
        dC(p,2400,i%2===0?WHITE:GREY,AlignmentType.LEFT,false,DGREY),
        dC(v,2400,i%2===0?WHITE:GREY,AlignmentType.CENTER,true,GREEN),
        dC(u,2400,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(r,2160,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  callout("💡  A 500-bird farm is the ideal entry scale for Kerala. Small enough to manage alone (or with family), large enough to generate meaningful income, and perfectly sized to learn operations before scaling to 1,000–2,200 birds.", LGREEN, GREEN),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 2: KERALA MARKET CONTEXT
// ═══════════════════════════════════════════════════
function marketContextSection(){ return [
  h1("2. KERALA POULTRY MARKET — WHY NOW?", BLUE),
  spacer(1),
  h2("Kerala Egg Market Overview"),
  para("Kerala is one of India's most significant egg-deficit states. Despite high consumer demand driven by a predominantly non-vegetarian population, local production meets only 40–45% of the state's daily requirement of approximately 2.2 crore eggs. The remaining 55–60% — over 1.2 crore eggs daily — is imported from Andhra Pradesh and Tamil Nadu."),
  spacer(1),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[3600,2880,2880],
    rows:[
      new TableRow({children:[hC("Market Indicator",3600),hC("Kerala Data",2880),hC("Implication for Farmer",2880)]}),
      ...([
        ["Daily egg consumption","~2.2 crore eggs","Massive sustained demand"],
        ["Local production","~1.0 crore eggs (45%)","Significant supply gap"],
        ["Daily import from AP/TN","~1.2 crore eggs","Transport cost = price premium for local"],
        ["Egg consumption per capita","~180 eggs/year","Above national average of 89"],
        ["Annual consumption growth","5–7%","Growing demand every year"],
        ["Farm-gate NECC price","₹5.00–5.75/egg","Kerala premium over national avg"],
        ["Retail price (local branded)","₹7.00–9.00/egg","35–60% premium over NECC"],
        ["Nattu Mutta (country egg)","₹12–18/egg","Premium niche, high margin"],
        ["Bakery egg demand (Kerala)","2nd highest in India","Kerala has 14,000+ bakeries"],
      ]).map(([ind,data,impl],i)=>new TableRow({children:[
        dC(ind, 3600,i%2===0?WHITE:GREY,AlignmentType.LEFT,false,DGREY),
        dC(data,2880,i%2===0?WHITE:GREY,AlignmentType.CENTER,true,BLUE),
        dC(impl,2880,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Why 500 Birds is the Right Starting Scale"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[3600,5760],
    rows:[
      new TableRow({children:[hC("Factor",3600),hC("Why 500 Birds is Ideal",5760)]}),
      ...([
        ["Capital risk","Total investment ₹8.85L — loss-limited, learnable. A mistake on 500 birds costs ₹50K max. On 5,000 birds it costs ₹5L+."],
        ["Management bandwidth","One person can manage 500 birds solo with part-time family help. No full-time hired labour needed initially."],
        ["Learning curve","All the critical skills — egg collection, grading, delivery, buyer management, disease spotting — are learned at low risk."],
        ["Cash generation","₹7,000–28,000/month net in Year 1. Enough to build buyer relationships and reinvest in expansion."],
        ["Scale-up path","Use Year 1 profits + NABARD subsidy to add 500 more birds in Month 16–18. Phase 2 with 1,000 birds."],
        ["NABARD eligibility","Small-scale farms (₹5–10L project) qualify for the most accessible NABARD schemes with least documentation."],
      ]).map(([f,why],i)=>new TableRow({children:[
        dC(f,  3600,i%2===0?LBLUE:GREY,AlignmentType.LEFT,true,BLUE),
        dC(why,5760,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 3: SITE SELECTION
// ═══════════════════════════════════════════════════
function siteSelectionSection(){ return [
  h1("3. SITE SELECTION & LAND REQUIREMENTS", TEAL),
  spacer(1),
  callout("📍  Choosing the right site is irreversible. A wrong site — poor drainage, too close to residential area, no road access — creates problems for the entire 18-year life of the farm. Spend time on this before construction.", LTEAL, TEAL),
  spacer(1),
  h2("Site Selection Criteria"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2400,2400,1560,2800],
    rows:[
      new TableRow({children:[hC("Criterion",2400,TEAL),hC("Ideal Specification",2400,TEAL),hC("Minimum",1560,TEAL),hC("Why It Matters",2800,TEAL)]}),
      ...([
        ["Distance from nearest house","300 metres+","100 metres","Odour, noise, disease complaints from neighbours"],
        ["Distance from other poultry farms","500 metres+","300 metres","Cross-contamination, disease transmission risk"],
        ["Road accessibility","All-weather road to gate","Motorable in monsoon","Feed delivery, egg collection vehicle access daily"],
        ["Elevation / drainage","Elevated, naturally sloped","Not low-lying","Kerala floods regularly — low sites flood in monsoon"],
        ["Water source","Borewell within 50m OR municipal supply","Reliable source nearby","500 birds need 150+ litres/day; no water = disaster"],
        ["Electricity connection","3-phase within 200m","Single phase OK","3-phase for larger pumps; single phase adequate for 500 birds"],
        ["Land area required","0.25 acres (1,000 sq.m)","0.15 acres","Shed + manure pit + buffer zone + movement space"],
        ["Soil type","Laterite or red soil","Any well-drained soil","Black cotton soil expands in rain, damages foundation"],
        ["Wind direction","Prevailing wind from south-west","Away from houses","Odour management — shed should not face residences downwind"],
        ["Pollution board distance","500m from water bodies","300m","Kerala's water body density is very high"],
      ]).map(([c,ideal,min,why],i)=>new TableRow({children:[
        dC(c,   2400,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,TEAL),
        dC(ideal,2400,i%2===0?WHITE:GREY),
        dC(min,  1560,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(why,  2800,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Kerala-Specific Site Evaluation Checklist"),
  h3("Flood Risk Assessment"),
  bullet("Visit site during July–August monsoon before buying or constructing — see actual drainage"),
  bullet("Check with Panchayat: is the plot in a flood-prone zone? Revenue records note flood history"),
  bullet("Elevation must be at least 0.5 metres above surrounding fields"),
  bullet("Calculate plinth height needed: minimum 2 feet above highest monsoon water level seen historically"),
  h3("Soil & Foundation"),
  bullet("Kerala laterite soil is ideal — stable, well-draining, strong foundation"),
  bullet("Black cotton soil (found in Palakkad, Thrissur borders): avoid or do deep foundation (extra cost)"),
  bullet("Coconut or rubber plantation land: check for root systems, old tree stumps below surface"),
  bullet("Get soil test from PWD lab or private lab (₹500–800) before finalising"),
  h3("Utility Connections"),
  bullet("Borewell feasibility: check groundwater map at Kerala Water Authority nearest office"),
  bullet("Kerala groundwater depth: 20–80 feet in most districts; Palakkad can be 100–150 feet"),
  bullet("Power line distance: Karnataka Electricity Board / KSEB — check transformer capacity for your area"),
  bullet("Road condition: a 300-metre mud road becomes impassable in monsoon — budget for gravel/concrete approach"),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 4: BREED SELECTION
// ═══════════════════════════════════════════════════
function breedSection(){ return [
  h1("4. BREED SELECTION — DEEP DIVE"),
  spacer(1),
  h2("Why Breed Choice Determines 30% of Your Success"),
  para("In a uniform housing and feeding system, breed is the single biggest variable that determines lay rate, egg weight, shell quality, heat tolerance, and feed efficiency. Getting the breed wrong on 500 birds costs ₹1–2 Lakhs in lost production over 18 months."),
  spacer(1),
  h2("Commercial Layer Breed Comparison"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1400,1120,1120,1120,1120,1120,2360],
    rows:[
      new TableRow({children:[
        hC("Breed",1400),hC("Eggs/Year",1120),hC("Egg Weight",1120),
        hC("Heat Tol.",1120),hC("Feed Conv.",1120),hC("Rating",1120),hC("Best For",2360),
      ]}),
      ...([
        ["BV-380 (Bovans)","340–350","62–65g","Excellent","1.9–2.0","⭐⭐⭐⭐⭐","Kerala's #1 choice — tropical climate adapted"],
        ["Hy-Line Brown","330–345","62–65g","Excellent","1.9–2.1","⭐⭐⭐⭐⭐","High lay rate, hardy, widely available"],
        ["Lohmann Brown","320–335","60–63g","Good","2.0–2.1","⭐⭐⭐⭐","Economical, disease tolerant"],
        ["Babcock B-300","310–330","60–63g","Good","2.0–2.2","⭐⭐⭐⭐","Budget option, widely available"],
        ["H&N Nick Brown","330–340","62–65g","Very Good","2.0–2.1","⭐⭐⭐⭐","Emerging option, good shell quality"],
        ["Giriraja / Desi","180–220","45–50g","Excellent","3.0–3.5","⭐⭐⭐","Nattu mutta premium market only"],
      ]).map(([b,e,ew,ht,fc,r,bf],i)=>new TableRow({children:[
        dC(b, 1400,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.LEFT,i===0,i===0?GREEN:DGREY),
        dC(e, 1120,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER,false,i===0?GREEN:DGREY),
        dC(ew,1120,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(ht,1120,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(fc,1120,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(r, 1120,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(bf,2360,i===0?LGREEN:i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Pullet vs Day-Old Chick — Which to Buy?"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2400,3480,3480],
    rows:[
      new TableRow({children:[hC("Option",2400),hC("Advantages",3480),hC("Disadvantages",3480)]}),
      new TableRow({children:[
        dC("Day-Old Chicks (DOC)\n₹35–50/chick",2400,LBLUE,AlignmentType.LEFT,true,BLUE),
        dC("• Lower upfront cost\n• You control rearing quality\n• Better disease immunity built under your conditions\n• Good if you have rearing infrastructure",3480,WHITE),
        dC("• 18 weeks of rearing before any income\n• Requires brooder, heater, chick starter feed\n• Higher skill level needed\n• Higher mortality risk in first 4 weeks",3480,WHITE),
      ]}),
      new TableRow({children:[
        dC("18-Week Pullets (Ready-to-Lay)\n₹250–320/pullet",2400,LGREEN,AlignmentType.LEFT,true,GREEN),
        dC("• Eggs within 3–4 weeks of arrival\n• Vaccinated, sexed, ready\n• Zero brooding cost or skill needed\n• RECOMMENDED for first-time farmers",3480,GREY),
        dC("• Higher upfront cost (₹1.5L for 500 birds)\n• Quality depends on hatchery reliability\n• Less control over early vaccination history",3480,GREY),
      ]}),
    ],
  }),
  spacer(1),
  callout("✅  RECOMMENDATION for 500-bird starter: Buy 18-week pullets. The ₹1,00,000–1,25,000 extra cost compared to DOC saves you 18 weeks of risk, skill, and infrastructure. Start generating revenue immediately.", LGREEN, GREEN),
  spacer(1),
  h2("Kerala Pullet Suppliers"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2400,2400,2400,2160],
    rows:[
      new TableRow({children:[hC("Supplier",2400),hC("Location",2400),hC("Breeds Available",2400),hC("Contact / Notes",2160)]}),
      ...([
        ["Suguna Hatcheries","Pan-Kerala (nearest depot)","BV-380, Hy-Line Brown","Most widely available; book 3 weeks advance"],
        ["Venky's India","Thrissur / Ernakulam","Hy-Line Brown, Lohmann","Quality consistent; slightly premium price"],
        ["KFC (Kerala Feeds Corp)","Multiple depots, all districts","BV-380","State govt supported; check stock availability"],
        ["Kegg Farms","Pune origin, delivered to Kerala","BV-380 Bovans","Excellent quality; longer transit; cold-chain needed"],
        ["Sri Padmavathy Hatcheries","Coimbatore (nearest TN)","Hy-Line, Babcock","Budget option; book minimum 2 weeks prior"],
      ]).map(([s,l,b,c],i)=>new TableRow({children:[
        dC(s,2400,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,DGREY),
        dC(l,2400,i%2===0?WHITE:GREY),
        dC(b,2400,i%2===0?WHITE:GREY),
        dC(c,2160,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 5: HOUSING & INFRASTRUCTURE
// ═══════════════════════════════════════════════════
function housingSection(){ return [
  h1("5. HOUSING & INFRASTRUCTURE — DETAILED SPECIFICATIONS"),
  spacer(1),
  h2("Shed Design for Kerala's Tropical Climate"),
  para("Kerala's climate demands a shed design that is radically different from northern India. High rainfall (avg 3,000mm), high humidity (70–90% RH), and heat (28–35°C) must all be managed simultaneously. Open-sided, east-west oriented sheds with deep overhangs are standard."),
  spacer(1),
  h2("Primary Shed Specifications (500 Birds)"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2400,2880,4080],
    rows:[
      new TableRow({children:[hC("Parameter",2400),hC("Specification",2880),hC("Detailed Notes",4080)]}),
      ...([
        ["Floor area","1,000–1,100 sq.ft","@2.0–2.2 sq.ft/bird in battery cage system. 22ft × 48ft or 20ft × 50ft are common."],
        ["Orientation","Length axis East–West","Minimises sun exposure on long walls; maximises south-west monsoon breeze cross-ventilation."],
        ["Wall height (sidewalls)","2.5 ft brick plinth + 5 ft wire mesh","Total 7.5 ft side wall. Brick for strength; mesh for maximum air flow above plinth."],
        ["Roof type","GI sheet (20-gauge) + false ceiling","GI sheet as outer roof. False ceiling of Calicut tile or thermocol reduces internal temp by 5–7°C."],
        ["Roof slope","1:4 pitch minimum","Steep pitch sheds rain quickly; Kerala monsoon can deliver 100mm+ in one hour."],
        ["Roof overhang","4.0 ft on all sides","CRITICAL in Kerala. Deep overhang prevents rain from driving inside shed."],
        ["Ridge ventilation","Open ridge (6-inch gap)","Hot air exits at top. Improves cross-ventilation significantly."],
        ["Floor","Plain cement concrete (PCC) 4 inch","Slope towards manure channels: 1:40 gradient. Easy cleaning, no cracks for rodents."],
        ["Plinth height","2.0–2.5 feet above ground","Protects from flood water during monsoon. Critical in low-lying areas."],
        ["Foundation","CC strip foundation","500mm deep × 600mm wide. Adequate for single-storey poultry shed."],
        ["Shed spacing","Minimum 20 feet if building 2 sheds","Air circulation between sheds; fire safety."],
      ]).map(([p,s,d],i)=>new TableRow({children:[
        dC(p,2400,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,DGREY),
        dC(s,2880,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(d,4080,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Battery Cage System — Detailed Specifications"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2800,2400,4160],
    rows:[
      new TableRow({children:[hC("Component",2800),hC("Specification",2400),hC("Notes",4160)]}),
      ...([
        ["Cage type","3-tier battery cage","Galvanised iron wire (GI). Do not buy mild steel — rusts in Kerala humidity."],
        ["Cage dimensions","45cm × 45cm per unit","Houses 3 birds comfortably (150 sq.cm/bird). 4 birds possible but reduces production."],
        ["Number of cages (500 birds)","167–170 cages","500 birds ÷ 3 birds/cage = 167 cages. Order 175 to allow for vacancies."],
        ["Tier arrangement","3 tiers high","Bottom tier: 18 inches from ground. Between tiers: manure tray or board."],
        ["Egg roll-out slope","7–8 degrees","Eggs roll forward to collection trough automatically."],
        ["Feeder trough","V-type, 40cm × 4.5cm opening","Runs length of cage front. 10cm per bird minimum feeding space."],
        ["Manure tray","GI tray or plastic board per tier","Slides out for daily cleaning. Prevents lower-tier bird contamination."],
        ["Cage material","GI wire 2.5mm gauge","Minimum 2mm; 2.5mm recommended for Kerala humidity resistance."],
        ["Galvanising quality","Hot-dip galvanised","Electro-galvanised rusts within 2 years in Kerala coastal humidity."],
        ["Cage supplier","Indiamart / local fabricator","Get quotes from 3 fabricators. Local Kerala fabricators often cheaper than pan-India."],
      ]).map(([c,s,n],i)=>new TableRow({children:[
        dC(c,2800,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,DGREY),
        dC(s,2400,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(n,4160,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Supporting Infrastructure Details"),
  h3("Nipple Drinker System (Strongly Recommended)"),
  bullet("Nipple drinkers reduce water wastage by 70% vs open troughs — critical in Kerala's hard water conditions"),
  bullet("System: overhead PVC pipe → pressure regulator → lateral pipes → nipple drinkers (1 per 5 birds = 100 nipples for 500 birds)"),
  bullet("Pressure regulator: set to 5–7 PSI for proper nipple flow"),
  bullet("Flush lines weekly to prevent mineral buildup — Kerala water is often hard (high TDS)"),
  bullet("Cost: ₹18,000–25,000 complete system for 500 birds"),
  h3("Automatic Feeder System"),
  bullet("V-trough manual feeding adequate for 500 birds — one person fills twice daily"),
  bullet("Automatic chain feeder (₹45,000): saves 30 min/day; recommended only if labour is a constraint"),
  bullet("Feed should be available ad-lib (always available) during daytime hours, restricted 2 hours before lights off"),
  h3("Lighting System"),
  bullet("LED lights: 40W per 1,000 sq.ft — warm white (2700K) preferred"),
  bullet("Lighting programme: 16 hours light per day (6 AM – 10 PM) during lay phase"),
  bullet("Timer switch (₹500): automatically turns lights on/off — removes human error"),
  bullet("Backup lighting: hurricane lantern or battery LED strip for power cuts"),
  bullet("Never suddenly reduce light hours during lay phase — production drops sharply"),
  h3("Fans & Ventilation"),
  bullet("2 exhaust fans (18-inch, 1,400 RPM) at opposite ends of shed"),
  bullet("Fan positioning: one near feed storage end, one at manure disposal end"),
  bullet("Run fans continuously during April–September (peak heat months in Kerala)"),
  bullet("Do not position fans blowing directly on birds — creates draughts, stress"),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 6: WATER MANAGEMENT
// ═══════════════════════════════════════════════════
function waterSection(){ return [
  h1("6. WATER MANAGEMENT — THE SILENT PRODUCTION KILLER", TEAL),
  spacer(1),
  callout("💧  Water quality issues cause 20–30% production drops without any obvious signs. A healthy layer drinks 200–300ml per day. 500 birds need 100–150 litres of clean water daily. Any shortage or quality problem shows up as a lay rate crash 5–7 days later.", LTEAL, TEAL, true),
  spacer(1),
  h2("Daily Water Requirement"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[3600,2880,2880],
    rows:[
      new TableRow({children:[hC("Parameter",3600),hC("Normal Day",2880),hC("Hot Day (>35°C)",2880)]}),
      ...([
        ["Water per bird per day","200–250 ml","300–400 ml (50–60% more)"],
        ["Total for 500 birds","100–125 litres","150–200 litres"],
        ["Overhead tank size required","3,000 litres minimum","5,000 litres in summer months"],
        ["Borewell pump runtime","1.5–2 hours/day","3–4 hours/day"],
        ["Water change frequency","Fresh water 2x daily","Fresh water 3x daily minimum"],
      ]).map(([p,n,h],i)=>new TableRow({children:[
        dC(p,3600,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,DGREY),
        dC(n,2880,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(h,2880,i%2===0?WHITE:GREY,AlignmentType.CENTER,false,RED),
      ]})),
    ],
  }),
  spacer(1),
  h2("Water Quality Standards for Layer Farms"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2400,2000,2000,2960],
    rows:[
      new TableRow({children:[hC("Parameter",2400),hC("Acceptable",2000),hC("Problem Level",2000),hC("Impact if Ignored",2960)]}),
      ...([
        ["pH","6.8–7.5","<6.5 or >8.0","Acidic water corrodes pipes; alkaline reduces mineral absorption"],
        ["TDS (Total Dissolved Solids)","< 1,000 ppm","1,000–3,000 ppm","Reduced water intake, mineral imbalance, kidney issues"],
        ["Hardness (CaCO₃)","< 250 ppm","> 500 ppm","Scale in pipes; nipple clogging; reduces vitamin absorption"],
        ["Bacteria (Coliform)","< 50 CFU/100ml","Any >50","Respiratory infections, gut disease, diarrhoea — production crash"],
        ["Chlorine (if treated)","0.5–1 ppm","> 3 ppm","Destroys good gut bacteria; affects feed efficiency"],
        ["Salinity (NaCl)","< 500 ppm","> 1,000 ppm","Diarrhoea, wet droppings, wet litter"],
      ]).map(([p,ok,bad,impact],i)=>new TableRow({children:[
        dC(p,    2400,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,DGREY),
        dC(ok,   2000,i%2===0?WHITE:GREY,AlignmentType.CENTER,false,GREEN),
        dC(bad,  2000,i%2===0?WHITE:GREY,AlignmentType.CENTER,false,RED),
        dC(impact,2960,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h3("Water Treatment Options for Kerala"),
  bullet("Chlorination: Add sodium hypochlorite (bleach) — 0.5ml per 100 litres of water (maintains 0.5ppm free chlorine)"),
  bullet("Acidification: Add citric acid (2g/litre stock solution, 1ml per litre drinking water) — lowers pH, kills bacteria"),
  bullet("UV steriliser: ₹8,000–15,000 one-time — best option for bore water; no chemicals"),
  bullet("Get water tested every 6 months at PHED lab (free) or private lab (₹500)"),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 7: FEED & NUTRITION
// ═══════════════════════════════════════════════════
function feedSection(){ return [
  h1("7. FEED & NUTRITION — COMPLETE PROGRAMME"),
  spacer(1),
  h2("Feeding Phases by Age"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1400,1200,1200,1200,1400,3360],
    rows:[
      new TableRow({children:[
        hC("Phase",1400),hC("Age",1200),hC("Protein %",1200),
        hC("ME kcal/kg",1200),hC("Feed/Bird/Day",1400),hC("Key Objective",3360),
      ]}),
      ...([
        ["Chick Starter","0–8 wks","22–23%","2,900","10g→50g","Organ development, immune system foundation. DOC: brooder heat required."],
        ["Grower","8–15 wks","15–16%","2,650","50–75g","Body frame, bone strength. Restrict feed slightly — avoid obese birds."],
        ["Pre-Lay","15–18 wks","17–18%","2,750","75–95g","Reproductive system maturation. Increase calcium to 2% now."],
        ["Layer Phase 1","18–40 wks","17–18%","2,750","110–115g","Peak production phase. Maximum nutrition, ad-lib feeding."],
        ["Layer Phase 2","40–60 wks","16–17%","2,750","115–120g","Maintain production. Increase calcium to 4%. Watch for moulting."],
        ["Layer Phase 3","60–72 wks","16–17%","2,700","115–120g","Late lay. Production declining naturally. Prepare for depletion."],
      ]).map(([ph,age,pr,me,fb,obj],i)=>new TableRow({children:[
        dC(ph, 1400,i>=3?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER,true,i>=3?GREEN:DGREY),
        dC(age,1200,i>=3?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(pr, 1200,i>=3?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(me, 1200,i>=3?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(fb, 1400,i>=3?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER,true),
        dC(obj,3360,i>=3?LGREEN:i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Layer Mash Ingredient Breakdown"),
  para("For 500 birds at 120g/bird/day, you need 60 kg of feed per day. Most Kerala farmers buy commercial pelleted/mash feed. However, understanding ingredients helps you evaluate feed quality and detect adulteration."),
  spacer(1),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2400,1400,1400,4160],
    rows:[
      new TableRow({children:[hC("Ingredient",2400),hC("% in Formula",1400),hC("kg per tonne",1400),hC("Role & Quality Indicator",4160)]}),
      ...([
        ["Maize (yellow corn)","50–55%","500–550 kg","Primary energy. Quality: dry, bright yellow, no mould smell. Moisture < 14%."],
        ["Soyabean meal (SBM)","20–25%","200–250 kg","Primary protein (46–48% protein). Never substitute with cheap protein alternatives."],
        ["Rice bran (deoiled)","5–10%","50–100 kg","Energy. Available abundantly in Kerala. Max 10% to avoid rancidity issues."],
        ["Limestone grit","8–10%","80–100 kg","Calcium for shell formation. Critical: poor calcium = thin/soft shells."],
        ["Dicalcium phosphate","1.5–2%","15–20 kg","Phosphorus. Bone strength, prevents leg weakness."],
        ["Fish meal","2–4%","20–40 kg","High quality protein, amino acids. Kerala: locally available from fishmeal units."],
        ["Vegetable oil","1.5–2%","15–20 kg","Energy density. Also improves pellet binding."],
        ["Common salt (NaCl)","0.3–0.4%","3–4 kg","Sodium, chloride. Excess: diarrhoea. Deficiency: poor laying."],
        ["Vitamin & mineral premix","0.25–0.5%","2.5–5 kg","Essential micronutrients. Buy from reputed manufacturer (Venkys, KW Alt)."],
        ["Methionine (DL-Meth)","0.15–0.20%","1.5–2 kg","Limiting amino acid. Without it, protein utilisation is poor."],
      ]).map(([ing,pct,kg,role],i)=>new TableRow({children:[
        dC(ing, 2400,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,DGREY),
        dC(pct, 1400,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(kg,  1400,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(role,4160,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Kerala-Specific Feed Management"),
  h3("Monsoon Feed Storage (Critical)"),
  bullet("Store feed bags on wooden pallets — never on floor. Kerala humidity causes floor-level condensation."),
  bullet("Maximum stock: 15 days at a time during June–September. Larger stocks develop mould."),
  bullet("Check feed visually daily: any caking, colour change, or musty smell → discard immediately"),
  bullet("Aflatoxin contamination in monsoon-damaged maize is the #1 silent killer of Kerala layers"),
  h3("Cost Reduction Strategies"),
  bullet("Buy in 3-tonne lots from KFC or Suguna feed depots — typically ₹1.50–2.00/kg cheaper than retail bags"),
  bullet("Mix 10% local rice bran into commercial feed — saves ₹0.80–1.00/kg on that fraction"),
  bullet("Limestone grit (calcium) can be sourced locally from stone crushing units in Palakkad / Thrissur at ₹3/kg"),
  bullet("Form buying group with 2–3 neighbouring farms for bulk discounts"),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 8: WEEK-BY-WEEK FLOCK MANAGEMENT
// ═══════════════════════════════════════════════════
function flockMgmtSection(){ return [
  h1("8. WEEK-BY-WEEK FLOCK MANAGEMENT CALENDAR", PURPLE),
  spacer(1),
  para("This is your operational playbook. Every action in the right week makes the difference between peak production and below-target performance."),
  spacer(1),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1200,1600,3480,3080],
    rows:[
      new TableRow({children:[
        hC("Bird Age",1200,PURPLE),hC("Lay Rate %",1600,PURPLE),
        hC("Key Actions",3480,PURPLE),hC("Watch Out For",3080,PURPLE),
      ]}),
      ...([
        ["Wk 18–20\n(Arrival)","0–5%","House pullets. Settle into cages for 48 hrs. Feed and fresh water immediately. Check all nipples flowing. Lights: 14 hours. DO NOT disturb for 3 days.",
          "Stress from transport. Offer electrolyte (Electral) in water for first 2 days."],
        ["Wk 21–22","5–25%","First eggs appearing. Small eggs initially — normal. Increase light to 15 hours. Record first egg date in register. Check for soft-shelled eggs.",
          "Excitement leads to overhandling. Reduce noise/visitor entry."],
        ["Wk 23–25","25–60%","Rapid lay rate climb. Ensure calcium (limestone) in feed is 3.5–4%. Increase feed to 110g/day. Monitor body weight — hens should not be losing weight.",
          "Peak protein demand. Do not compromise on feed quality now."],
        ["Wk 26–40","85–92%","PEAK PRODUCTION. Maintain 16-hour lighting. Feed 115–120g/day. Collect eggs 3x daily. Grade and dispatch within 24 hours. Start building direct buyer base.",
          "Heat stress (April–June). Run fans continuously. Offer cool fresh water 3x/day."],
        ["Wk 41–52","80–88%","Stable high production. First batch of spent hen buyers to be identified now. Monitor for moulting birds — isolate if any.",
          "Mid-cycle calcium deficiency shows as pale yolks, thin shells."],
        ["Wk 53–65","75–82%","Slightly declining. Natural reduction. Ensure lighting still 16 hours. Consider supplementing Vitamin E and Selenium for shell quality.",
          "First major moulting wave possible at ~15 months. Isolate moulting birds."],
        ["Wk 66–72","65–75%","Final lay phase. Production declining. Begin planning depletion — contact spent hen buyers 8 weeks in advance. Start counting remaining birds.",
          "Sudden production drops may indicate Newcastle or IB — verify with vet."],
        ["Wk 73–74\n(Depletion)","N/A","Sell all birds to spent hen market. Clean and disinfect shed completely. 3-week gap before new pullets. Evaluate previous batch performance vs targets.",
          "Get best price by booking spent hen buyers 8 weeks prior."],
      ]).map(([age,lay,action,watch],i)=>new TableRow({children:[
        dC(age,   1200,i===3?LGREEN:i%2===0?LPURP:GREY,AlignmentType.CENTER,true,i===3?GREEN:PURPLE),
        dC(lay,   1600,i===3?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER,i===3,i===3?GREEN:DGREY),
        dC(action,3480,i===3?LGREEN:i%2===0?WHITE:GREY),
        dC(watch, 3080,i===3?LGREEN:i%2===0?WHITE:GREY,AlignmentType.LEFT,false,RED),
      ]})),
    ],
  }),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 9: DISEASE MANAGEMENT
// ═══════════════════════════════════════════════════
function diseaseSection(){ return [
  h1("9. DISEASE MANAGEMENT & VACCINATION"),
  spacer(1),
  h2("Vaccination Schedule (For 18-week Pullets Purchased)"),
  para("If you buy 18-week vaccinated pullets, the hatchery handles weeks 1–18. Confirm vaccination history with supplier. Then continue from week 18 onwards:"),
  spacer(1),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1200,2400,1800,1800,2160],
    rows:[
      new TableRow({children:[hC("Age",1200),hC("Vaccine / Disease",2400),hC("Route",1800),hC("Brand Example",1800),hC("Notes",2160)]}),
      ...([
        ["Wk 18 (on arrival)","Ranikhet (Lasota strain)","Eye drop / Drinking water","RAKSHA-NCD (Biogenics)","Booster — critical on arrival after transport stress"],
        ["Wk 20","Fowl Pox","Wing web stab","RAKSHA-POX","Kerala has year-round mosquito pressure — do not skip"],
        ["Wk 24","Ranikhet (R2B booster)","IM Injection","R2B — Venky's / Hester","Mid-lay booster. Gives 3–4 months protection."],
        ["Wk 36","Egg Drop Syndrome (EDS-76)","IM Injection","EDS-76 (CEVA / Biogenics)","Critical: prevents sudden production drops at 9 months"],
        ["Wk 48","Ranikhet (Lasota)","Drinking water","RAKSHA-NCD","Annual booster. Do not miss."],
        ["Every 3 months","Ranikhet surveillance","Monitor / booster if needed","—","If mortality spikes, vaccinate immediately even off-schedule"],
      ]).map(([age,vac,route,brand,notes],i)=>new TableRow({children:[
        dC(age,  1200,i%2===0?WHITE:GREY,AlignmentType.CENTER,true,DGREY),
        dC(vac,  2400,i%2===0?WHITE:GREY,AlignmentType.LEFT,false,DGREY),
        dC(route,1800,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(brand,1800,i%2===0?WHITE:GREY),
        dC(notes,2160,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Common Diseases — Identification & Response"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1800,2200,2200,3160],
    rows:[
      new TableRow({children:[hC("Disease",1800),hC("Symptoms",2200),hC("Production Impact",2200),hC("Response",3160)]}),
      ...([
        ["Ranikhet (Newcastle)","Respiratory sounds, torticollis (twisted neck), green diarrhoea, sudden death","30–100% production loss. Often fatal to entire flock.","Emergency: isolate birds, call vet immediately, emergency vaccination of unaffected birds"],
        ["Infectious Bronchitis (IB)","Respiratory distress, watery egg white, egg shell deformities","15–35% production loss, egg quality drop","Vet consult; supportive treatment; ensure Vit C in water"],
        ["Egg Drop Syndrome (EDS)","Sudden drop in production, thin-shelled eggs, pale yolks","20–30% production loss — can persist for weeks","Vaccination is only prevention. No treatment once infected."],
        ["Marek's Disease","Paralysis, grey eye, tumours","Early flock loss (young birds)","Vaccination at hatchery. No treatment. Maintain flock hygiene."],
        ["Coccidiosis","Bloody diarrhoea, lethargy","Mortality + production drop","Battery cage system reduces risk significantly. Amprolium in water if confirmed."],
        ["Fowl Pox","Skin lesions on comb/face (dry pox) or diphtheric lesions in mouth (wet pox)","10–20% drop","Vaccination before monsoon. Wet pox: vet treatment needed immediately."],
        ["Colibacillosis (E. coli)","Respiratory, swollen face, peritonitis in hens","Mortality, egg quality drop","Water sanitation, antibiotic treatment per vet prescription only"],
      ]).map(([dis,sym,imp,resp],i)=>new TableRow({children:[
        dC(dis, 1800,i%2===0?LRED:GREY,AlignmentType.LEFT,true,RED),
        dC(sym, 2200,i%2===0?WHITE:GREY),
        dC(imp, 2200,i%2===0?WHITE:GREY,AlignmentType.CENTER,false,RED),
        dC(resp,3160,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 10: WASTE & MANURE MANAGEMENT
// ═══════════════════════════════════════════════════
function wasteSection(){ return [
  h1("10. WASTE & MANURE MANAGEMENT", TEAL),
  spacer(1),
  callout("♻️  500 birds produce approximately 100 kg of wet manure per day (2.5–3 tonnes per month). Managed properly, this becomes an income stream worth ₹5,000–8,000 per month. Managed poorly, it becomes a disease risk, odour complaint, and Pollution Board violation.", LTEAL, TEAL),
  spacer(1),
  h2("Manure Production & Composition"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[3600,2880,2880],
    rows:[
      new TableRow({children:[hC("Parameter",3600),hC("Value (500 birds)",2880),hC("Annual Total",2880)]}),
      ...([
        ["Wet manure produced / day","90–100 kg","3.0 tonnes / month"],
        ["Moisture content (fresh)","70–75%","—"],
        ["Dry matter content","25–30%","—"],
        ["Dried manure yield / month","700–900 kg","8.4–10.8 tonnes/year"],
        ["Nitrogen (N) content","2.5–3.5% dry basis","Excellent organic fertiliser"],
        ["Phosphorus (P) content","2.0–2.8% dry basis","—"],
        ["Potassium (K) content","1.5–2.0% dry basis","—"],
        ["Revenue (bulk @ ₹3/kg wet)","₹8,100–9,000/month","₹97,200–1,08,000/year"],
        ["Revenue (packaged @ ₹8/kg dry)","₹5,600–7,200/month","₹67,200–86,400/year"],
      ]).map(([p,v,a],i)=>new TableRow({children:[
        dC(p,3600,i%2===0?WHITE:GREY,AlignmentType.LEFT,false,DGREY),
        dC(v,2880,i%2===0?WHITE:GREY,AlignmentType.CENTER,true,TEAL),
        dC(a,2880,i%2===0?WHITE:GREY,AlignmentType.CENTER),
      ]})),
    ],
  }),
  spacer(1),
  h2("Manure Management Options"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1800,2400,2400,2760],
    rows:[
      new TableRow({children:[hC("Method",1800,TEAL),hC("Process",2400,TEAL),hC("Income Potential",2400,TEAL),hC("Investment",2760,TEAL)]}),
      ...([
        ["Bulk sale (wet)","Collect daily, store in covered pit, sell to vegetable farmers monthly","₹3,000–4,000/month","Nil — just a collection pit ₹5,000"],
        ["Sun drying + bulk","Spread on concrete pad, dry 3–5 days (if weather permits), sell dry","₹5,000–6,000/month","Concrete drying pad: ₹15,000"],
        ["Solar dryer","Polythene tunnel dryer, dries in 2–3 days regardless of weather","₹6,000–8,000/month","₹25,000–35,000 one-time"],
        ["Vermicomposting","Mix with crop waste, add earthworms, compost for 60 days — premium product","₹8,000–12,000/month","₹15,000 setup; 60-day cycle"],
        ["Biogas plant","Feed into biogas digester — cooking gas for farm + slurry as fertiliser","Gas savings ₹1,500–2,000/month","₹50,000–80,000 capital"],
      ]).map(([m,p,inc,inv],i)=>new TableRow({children:[
        dC(m,  1800,i%2===0?LTEAL:GREY,AlignmentType.LEFT,true,TEAL),
        dC(p,  2400,i%2===0?WHITE:GREY),
        dC(inc,2400,i%2===0?WHITE:GREY,AlignmentType.CENTER,true,GREEN),
        dC(inv,2760,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h3("Kerala Pollution Board Compliance for Manure"),
  bullet("1,000+ birds require KSPCB consent. 500-bird farms are below threshold but still maintain good practices."),
  bullet("Never dump manure in streams, canals, or near water bodies — heavy Kerala rainfall washes it into water sources"),
  bullet("Covered manure pit mandatory: minimum 2m × 3m × 1m deep, cement-lined, covered with GI sheet"),
  bullet("Manure buyers: vegetable farmers in Palakkad, Thrissur, Ernakulam — they often collect free or pay ₹500–1,000/load"),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 11: LABOUR MANAGEMENT
// ═══════════════════════════════════════════════════
function labourSection(){ return [
  h1("11. LABOUR MANAGEMENT", ORANGE),
  spacer(1),
  h2("Labour Requirement by Scale"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2400,2400,4560],
    rows:[
      new TableRow({children:[hC("Scale",2400,ORANGE),hC("Labour Required",2400,ORANGE),hC("Typical Arrangement",4560,ORANGE)]}),
      ...([
        ["500 birds","Owner + part-time family help","Owner manages alone, spouse or family member helps at collection (1.5 hrs/day). No hired labour needed."],
        ["1,000 birds","Owner + 1 part-time hired worker","Hired worker for morning collection, feeding, cleaning. ₹6,000–8,000/month part-time."],
        ["2,000+ birds","Owner + 2 full-time workers","Two full-time workers at ₹12,000/month each. Owner does management, delivery, buying."],
      ]).map(([s,l,a],i)=>new TableRow({children:[
        dC(s,2400,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.LEFT,true,i===0?GREEN:DGREY),
        dC(l,2400,i===0?LGREEN:i%2===0?WHITE:GREY),
        dC(a,4560,i===0?LGREEN:i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Daily Work Schedule — Owner-Operated 500-Bird Farm"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1400,2400,5560],
    rows:[
      new TableRow({children:[hC("Time",1400,ORANGE),hC("Task",2400,ORANGE),hC("Detail",5560,ORANGE)]}),
      ...([
        ["5:30 AM","Switch on lights","Lights on before sunrise ensures hens are active and laying from 6 AM"],
        ["6:00 AM","First egg collection","Collect all eggs laid overnight. Count, check for cracks. Fill water nipples."],
        ["6:30 AM","Morning feed","Measure and add feed to troughs. Check feed stock remaining."],
        ["7:00 AM","Manure inspection","Check under cage trays — remove and move to pit. Look for unusual droppings."],
        ["7:30 AM","Dead bird check","Inspect all cages for dead birds. Record in mortality register if any."],
        ["8:00–10:00 AM","Delivery","Pack and deliver yesterday's eggs + today's morning collection to hotels, shops. WhatsApp updates."],
        ["10:00 AM","Second collection","Collect mid-morning eggs. Grade and pack."],
        ["12:00 PM","Water check","Check all nipples flowing. Flush lines if needed."],
        ["3:00 PM","Third collection","Afternoon egg collection — largest collection of the day."],
        ["4:00 PM","Feed top-up","Top up troughs. Ensure birds have feed until lights off."],
        ["6:00 PM","Health observation","Walk through shed, observe bird behaviour, posture, droppings."],
        ["8:00 PM","Lighting check","Confirm 16-hour lighting target reached. Switch to dim or off."],
        ["Evening","Records","Fill daily egg collection log, mortality register if needed, WhatsApp status update."],
      ]).map(([t,task,detail],i)=>new TableRow({children:[
        dC(t,    1400,i%2===0?LORANG:GREY,AlignmentType.CENTER,true,ORANGE),
        dC(task, 2400,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,DGREY),
        dC(detail,5560,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 12: FINANCIAL PROJECTIONS — EXPANDED
// ═══════════════════════════════════════════════════
function financialSection(){ return [
  h1("12. FINANCIAL PROJECTIONS — MONTH-BY-MONTH DETAIL", GREEN),
  spacer(1),
  h2("A. Total Capital Investment"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[4800,2280,2280],
    rows:[
      new TableRow({children:[hC("Item",4800),hC("Amount (₹)",2280),hC("Notes",2280)]}),
      ...([
        ["Land preparation & levelling","30,000","Owned land assumed"],
        ["Shed construction (1,100 sq.ft @ ₹160/sq.ft)","1,75,000","Labour + materials"],
        ["Battery cages (175 cages, 3-tier GI)","1,40,000","Hot-dip galvanised"],
        ["Nipple drinker system (complete)","22,000","Pipes, nipples, tank fittings"],
        ["Feeders (V-trough, 1,100 ft run)","15,000","GI or PVC troughs"],
        ["Overhead water tank (3,000L + stand)","18,000","HDPE tank + stand"],
        ["Borewell (80 ft depth) + pump","45,000","Including motor & piping"],
        ["Generator (3.5 KVA diesel)","40,000","Mandatory for Kerala grid reliability"],
        ["Electrical wiring, fans, LED lights","35,000","Including timer switch"],
        ["Manure pit (concrete, covered)","8,000","3m × 2m × 1m"],
        ["Farm tools & miscellaneous","12,000","Scale, candler, stamp, trays, etc."],
        ["Contingency (10%)","54,000","Buffer for cost overruns"],
        ["TOTAL INFRASTRUCTURE","5,94,000",""],
        ["","",""],
        ["18-week Pullets (500 × ₹300)","1,50,000","BV-380 from hatchery"],
        ["Feed stock (2 months buffer, 3,600 kg)","1,00,800","At ₹28/kg"],
        ["Medicines & vaccines (initial)","15,000","Arrival + first 3 months"],
        ["Working capital (3 months running)","25,200","Misc monthly expenses"],
        ["TOTAL WORKING CAPITAL","2,91,000",""],
        ["","",""],
        ["GRAND TOTAL INVESTMENT","8,85,000",""],
        ["Less: NABARD Subsidy (25%)","(1,47,500)","25% back-ended subsidy"],
        ["NET EFFECTIVE INVESTMENT","7,37,500","After NABARD subsidy"],
      ]).map(([item,amt,notes],i)=>{
        const isTotal = item.includes("TOTAL")||item.includes("GRAND")||item.includes("NET EFFECTIVE");
        const isEmpty = item==="";
        return new TableRow({children:[
          dC(item,4800,isTotal?LGREEN:isEmpty?WHITE:i%2===0?WHITE:GREY,AlignmentType.LEFT,isTotal,isTotal?GREEN:DGREY),
          dC(amt, 2280,isTotal?LGREEN:isEmpty?WHITE:i%2===0?WHITE:GREY,AlignmentType.RIGHT,isTotal,isTotal?GREEN:DGREY),
          dC(notes,2280,isTotal?LGREEN:isEmpty?WHITE:i%2===0?WHITE:GREY),
        ]});
      }),
    ],
  }),
  spacer(1),
  h2("B. Month-by-Month Cash Flow — Year 1"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[680,880,880,880,880,880,880,880,880,1040,1200],
    rows:[
      new TableRow({children:[
        hC("Month",680,BLUE),hC("Lay\nRate %",880,BLUE),hC("Eggs\n/day",880,BLUE),
        hC("Avg Price\n₹/egg",880,BLUE),hC("Egg\nRevenue",880,BLUE),hC("Feed\nCost",880,BLUE),
        hC("Other\nCosts",880,BLUE),hC("Total\nCost",880,BLUE),hC("Net\nProfit",880,BLUE),
        hC("Cumul.\nProfit",1040,BLUE),hC("Notes",1200,BLUE),
      ]}),
      ...([
        [1,"5%",25,"5.50","4,125","51,100","19,417","70,517","(66,392)","(66,392)","Birds settle; first eggs appear"],
        [2,"20%",100,"5.75","17,250","51,100","19,417","70,517","(53,267)","(1,19,659)","Production climbing fast"],
        [3,"45%",225,"6.00","40,500","51,100","19,417","70,517","(30,017)","(1,49,676)","Building buyer relationships"],
        [4,"70%",350,"6.25","65,625","51,100","19,417","70,517","(4,892)","(1,54,568)","Near break-even on monthly cost"],
        [5,"82%",410,"6.50","79,950","51,100","19,417","70,517","9,433","(1,45,135)","First profitable month"],
        [6,"88%",440,"6.50","85,800","51,100","19,417","70,517","15,283","(1,29,852)","Peak approaching"],
        [7,"90%",450,"6.75","91,125","51,100","19,417","70,517","20,608","(1,09,244)","Peak production"],
        [8,"90%",450,"7.00","94,500","51,100","19,417","70,517","23,983","(85,261)","Premium channel improving"],
        [9,"88%",440,"7.00","92,400","51,100","19,417","70,517","21,883","(63,378)",""],
        [10,"87%",435,"7.00","91,350","51,100","19,417","70,517","20,833","(42,545)",""],
        [11,"86%",430,"7.25","93,525","51,100","19,417","70,517","23,008","(19,537)",""],
        [12,"85%",425,"7.25","92,438","51,100","19,417","70,517","21,921","2,384","Capital recovered by Yr end"],
      ]).map(([mo,lr,eggs,price,rev,feed,other,total,net,cumul,notes],i)=>{
        const isProfit = parseFloat(net.replace(/[(),]/g,''))>0 && !net.includes("(");
        return new TableRow({children:[
          dC(String(mo),680,i%2===0?WHITE:GREY,AlignmentType.CENTER,true,BLUE),
          dC(lr,880,i%2===0?WHITE:GREY,AlignmentType.CENTER),
          dC(eggs,880,i%2===0?WHITE:GREY,AlignmentType.CENTER),
          dC(price,880,i%2===0?WHITE:GREY,AlignmentType.CENTER),
          dC(rev,880,i%2===0?WHITE:GREY,AlignmentType.RIGHT),
          dC(feed,880,i%2===0?WHITE:GREY,AlignmentType.RIGHT),
          dC(other,880,i%2===0?WHITE:GREY,AlignmentType.RIGHT),
          dC(total,880,i%2===0?WHITE:GREY,AlignmentType.RIGHT),
          dC(net,880,isProfit?LGREEN:LRED,AlignmentType.RIGHT,true,isProfit?GREEN:RED),
          dC(cumul,1040,i%2===0?WHITE:GREY,AlignmentType.RIGHT,false,cumul.includes("(")?RED:GREEN),
          dC(notes,1200,i%2===0?WHITE:GREY),
        ]});
      }),
    ],
  }),
  spacer(1),
  h2("C. Year 2 Financial Summary"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[5360,2000,2000],
    rows:[
      new TableRow({children:[hC("Item",5360),hC("Conservative (₹)",2000),hC("Optimistic (₹)",2000)]}),
      ...([
        ["Egg sales (avg 78% lay, 380 eggs/day, 365 days)","10,03,500","11,86,875"],
        ["Spent hen sale (500 × ₹200 = once per 18 months — prorated)","66,667","83,333"],
        ["Manure income (packaged / bulk)","54,000","90,000"],
        ["GROSS ANNUAL REVENUE","11,24,167","13,60,208"],
        ["Feed cost (120g/bird/day × 500 × 365 × ₹28/kg)","6,13,200","6,13,200"],
        ["Labour (part-time family help)","72,000","72,000"],
        ["Electricity + generator fuel","36,000","36,000"],
        ["Medicine, vaccines, supplements","24,000","24,000"],
        ["Delivery fuel, packaging, misc","36,000","36,000"],
        ["TOTAL ANNUAL OPERATING COST","7,81,200","7,81,200"],
        ["NET ANNUAL PROFIT (Year 2)","3,42,967","5,79,008"],
        ["NET MONTHLY PROFIT (Year 2)","28,580","48,251"],
      ]).map(([item,cons,opt],i)=>{
        const isTotal = item.includes("TOTAL")||item.includes("NET")||item.includes("GROSS");
        const isEmpty = item==="";
        const shade = item.includes("NET ANNUAL")||item.includes("NET MONTHLY") ? LGREEN :
                      item.includes("GROSS") ? LBLUE : isTotal ? GREY : i%2===0?WHITE:GREY;
        return new TableRow({children:[
          dC(item,5360,shade,AlignmentType.LEFT,isTotal,isTotal?GREEN:DGREY),
          dC(cons,2000,shade,AlignmentType.RIGHT,isTotal,isTotal?GREEN:DGREY),
          dC(opt, 2000,shade,AlignmentType.RIGHT,isTotal,isTotal?GREEN:DGREY),
        ]});
      }),
    ],
  }),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 13: NABARD LOAN & SUBSIDY GUIDE
// ═══════════════════════════════════════════════════
function nabardSection(){ return [
  h1("13. NABARD LOAN & SUBSIDY APPLICATION GUIDE", PURPLE),
  spacer(1),
  callout("💰  NABARD back-ended subsidy of 25% (general) or 33% (SC/ST/NE) of project cost is available for poultry farms. On ₹8.85L project, this is ₹1.47–2.92 Lakhs back to your account. Do NOT start construction before applying — you will lose subsidy eligibility.", LPURP, PURPLE, true),
  spacer(1),
  h2("NABARD Subsidy Scheme Details"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[3600,2880,2880],
    rows:[
      new TableRow({children:[hC("Scheme Parameter",3600),hC("General Category",2880),hC("SC/ST/NE/Women",2880)]}),
      ...([
        ["Subsidy %","25% of project cost","33.33% of project cost"],
        ["Maximum subsidy","₹25,00,000 (₹25 Lakhs)","₹25,00,000 (₹25 Lakhs)"],
        ["On ₹8.85L project","₹1,47,500 back","₹1,96,625 back"],
        ["Own contribution required","25% of project cost","10% of project cost"],
        ["Bank loan covers","50% of project cost","57% of project cost"],
        ["Type of subsidy","Back-ended (credited to loan account after 2 years)","Back-ended"],
        ["Eligible activity","Layer poultry, broiler, hatchery, feed plant","Same"],
        ["Bank channel","Nationalised / scheduled commercial / cooperative banks","Same"],
      ]).map(([p,g,sc],i)=>new TableRow({children:[
        dC(p, 3600,i%2===0?WHITE:GREY,AlignmentType.LEFT,false,DGREY),
        dC(g, 2880,i%2===0?WHITE:GREY,AlignmentType.CENTER,true,BLUE),
        dC(sc,2880,i%2===0?WHITE:GREY,AlignmentType.CENTER,true,GREEN),
      ]})),
    ],
  }),
  spacer(1),
  h2("Step-by-Step NABARD Application Process"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[800,8560],
    rows:[
      new TableRow({children:[hC("Step",800),hC("Action",8560)]}),
      ...([
        ["1","Meet your bank's Agriculture or SME Loan Manager. Say: 'I want to apply for a poultry loan under NABARD subsidy scheme.' Any nationalised bank (SBI, Canara, Indian Bank) or cooperative bank works."],
        ["2","Collect bank's DPR format. Different banks have slightly different formats but all require same information as in this project report."],
        ["3","Submit this project report as your DPR (add your name, farm location, personal details on cover page)."],
        ["4","Attach documents: Aadhaar + PAN, land ownership/lease proof, last 3 years income tax returns (or Form 16), bank passbook (6 months), 2 passport photos."],
        ["5","Bank conducts site visit and loan committee review. Timeline: 3–8 weeks depending on bank."],
        ["6","Loan sanction letter received. DO NOT start construction until this letter is in hand."],
        ["7","Loan disbursed in 2–3 tranches linked to construction progress. First tranche: ₹2–3L to start construction."],
        ["8","NABARD subsidy is registered simultaneously. After 2 years of regular repayment, ₹1.47–1.96L is credited back to your loan account — reducing outstanding principal."],
        ["9","Continue EMI repayment. The subsidy credit in Year 2 feels like a windfall — use it to prepay loan or expand flock."],
      ]).map(([step,action],i)=>new TableRow({children:[
        dC(step,  800,i%2===0?LPURP:GREY,AlignmentType.CENTER,true,PURPLE),
        dC(action,8560,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h3("Other Schemes for Kerala Poultry Farmers"),
  bullet("KFC (Kerala Feeds Corp) subsidised feed: Register at nearest depot — feed at ₹2–3/kg below market rate for registered farmers"),
  bullet("Kudumbashree Poultry Unit: Women-led self-help groups can access NRLM-backed poultry grants up to ₹75,000"),
  bullet("KVASU extension services: Free veterinary consultation, technical guidance from Kerala Veterinary University — contact KVASU, Mannuthy (Thrissur)"),
  bullet("SFAC small farmer scheme: Small Farmers' Agri-Business Consortium — 25% capital grant up to ₹50,000"),
  bullet("Mudra Loan (Kishor scheme): Collateral-free loan up to ₹10 Lakhs through any bank — no subsidy but no collateral needed"),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 14: MARKETING & EGG SALES
// ═══════════════════════════════════════════════════
function marketingSection(){ return [
  h1("14. MARKETING & EGG SALES STRATEGY"),
  spacer(1),
  h2("Target Channel Mix for 500 Birds"),
  para("At 380 eggs/day average production, here is the recommended channel split to achieve a blended price of ₹6.75–7.25/egg:"),
  spacer(1),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2800,960,960,1000,1400,2240],
    rows:[
      new TableRow({children:[
        hC("Channel",2800),hC("% Share",960),hC("Eggs/Day",960),
        hC("Target Price",1000),hC("Monthly Rev",1400),hC("Priority",2240),
      ]}),
      ...([
        ["Direct home delivery (subscription)","30%",114,"₹8.00","27,360","Build first — highest margin"],
        ["Hotels, bakeries, restaurants","28%",106,"₹7.00","22,260","Approach during construction"],
        ["Institutional (hospital/school canteen)","17%",65,"₹6.50","12,675","Steady, volume, monthly payment"],
        ["Local retail grocery shops","20%",76,"₹5.75","13,110","Easiest to close — approach immediately"],
        ["NECC/trader overflow","5%",19,"₹5.25","3,003","Emergency only — lowest price"],
        ["TOTAL / BLENDED","100%",380,"₹6.95","78,408","Target blended: ₹6.75–7.00"],
      ]).map(([ch,pct,eggs,price,rev,pri],i)=>{
        const isTotal = ch.includes("TOTAL");
        return new TableRow({children:[
          dC(ch,  2800,isTotal?LGREEN:i%2===0?WHITE:GREY,AlignmentType.LEFT,isTotal,isTotal?GREEN:DGREY),
          dC(pct, 960, isTotal?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER,isTotal),
          dC(String(eggs),960,isTotal?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
          dC(price,1000,isTotal?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER,true,isTotal?GREEN:GOLD),
          dC(rev,  1400,isTotal?LGREEN:i%2===0?WHITE:GREY,AlignmentType.RIGHT,isTotal,isTotal?GREEN:DGREY),
          dC(pri,  2240,isTotal?LGREEN:i%2===0?WHITE:GREY),
        ]});
      }),
    ],
  }),
  spacer(1),
  h2("Pre-Production Buyer Network Plan"),
  h3("Month 1–2 (During Construction) — 30 contacts minimum"),
  bullet("Visit every hotel, bakery, and tea shop within 10 km — introduce yourself, collect numbers"),
  bullet("Map 3+ apartment complexes within delivery range — meet the resident association secretary"),
  bullet("Identify nearest hospital/school — get name of canteen contractor (not admin — contractor)"),
  bullet("Create WhatsApp Business account, post first farm photo with 'Coming in 60 days' message"),
  h3("Month 3 (6 Weeks Before Birds Arrive) — First commitments"),
  bullet("Follow up with 5–8 best prospects. Bring sample eggs from a neighbouring farm."),
  bullet("Offer 7-day free trial to 2 hotels — let quality speak"),
  bullet("Launch WhatsApp pre-booking: '30 eggs/week subscription — ₹255/week — advance booking open now'"),
  h3("Month 4 (Birds Arrive) — Lock 70% capacity before first egg"),
  bullet("Minimum 250 eggs/day in committed orders before birds arrive — non-negotiable"),
  bullet("Maintain daily WhatsApp status updates even before production starts"),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 15: GOVERNMENT SCHEMES
// ═══════════════════════════════════════════════════
function schemesSection(){ return [
  h1("15. GOVERNMENT SCHEMES & SUPPORT"),
  spacer(1),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2200,1800,1800,1800,1760],
    rows:[
      new TableRow({children:[hC("Scheme",2200),hC("Benefit",1800),hC("Eligibility",1800),hC("How to Apply",1800),hC("Timeline",1760)]}),
      ...([
        ["NABARD Poultry Scheme","25–33% subsidy on project cost","All farmers; more for SC/ST/Women","Via bank — submit DPR before construction","3–8 weeks"],
        ["SFAC Small Farmer Scheme","25% capital grant, max ₹50,000","Small & marginal farmers","District Agriculture Office","4–6 weeks"],
        ["Kudumbashree (Women)","Grant up to ₹75,000","Women SHG members","Kudumbashree block office","6–8 weeks"],
        ["KFC Subsidised Feed","₹2–3/kg below market","All registered Kerala farmers","Farm registration + KFC depot visit","Immediate after registration"],
        ["KVASU Technical Support","Free vet + technical guidance","All poultry farmers","KVASU Mannuthy, Thrissur; or district campus","Ongoing"],
        ["Mudra Loan (Kishor)","Up to ₹10L, no collateral","Any small business","Any nationalized bank","2–4 weeks"],
        ["Animal Husbandry Dept","Subsidised day-old chicks","BPL / marginal farmers","District AH office","Scheme-specific; varies"],
        ["PM-Kisan (indirect)","₹6,000/year income support","Small farmer land holders","CSC centre / PM-Kisan portal","Ongoing"],
      ]).map(([s,b,e,h,t],i)=>new TableRow({children:[
        dC(s,2200,i%2===0?LPURP:GREY,AlignmentType.LEFT,true,PURPLE),
        dC(b,1800,i%2===0?WHITE:GREY),
        dC(e,1800,i%2===0?WHITE:GREY),
        dC(h,1800,i%2===0?WHITE:GREY),
        dC(t,1760,i%2===0?WHITE:GREY,AlignmentType.CENTER),
      ]})),
    ],
  }),
  spacer(1),
  callout("🔑  Submit this DPR to your nearest bank branch or District Industries Centre (DIC) to initiate the NABARD subsidy application — along with your land documents and identity proof.", LGREEN, GREEN),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 16: EQUIPMENT & SUPPLIER DIRECTORY
// ═══════════════════════════════════════════════════
function supplierSection(){ return [
  h1("16. EQUIPMENT & SUPPLIER DIRECTORY — KERALA", TEAL),
  spacer(1),
  para("Use this directory to source materials locally where possible. Local suppliers reduce freight cost, provide faster replacement, and offer better after-sales support."),
  spacer(1),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1600,2000,2400,3360],
    rows:[
      new TableRow({children:[hC("Item",1600,TEAL),hC("Best Source",2000,TEAL),hC("Expected Cost (500 birds)",2400,TEAL),hC("Notes",3360,TEAL)]}),
      ...([
        ["Battery cages","Ernakulam / Thrissur fabricators OR Indiamart","₹1,30,000–1,50,000","Get 3 quotes; insist on hot-dip galvanised"],
        ["Layer feed (commercial)","Suguna / Venky's / KFC depot","₹27–30/kg (bulk rate)","Buy in 2-tonne lots minimum"],
        ["18-week pullets","Suguna Hatcheries (nearest depot)","₹280–320/pullet","Book 3 weeks advance; 500 birds minimum"],
        ["Nipple drinker system","Thiruvalla / Ernakulam agri-input shops","₹18,000–25,000","Lubing or Roxell brand preferred"],
        ["Vaccines","Licensed veterinary pharmacy — any taluk town","Variable — see vaccine schedule","Cold-chain storage mandatory"],
        ["GI roofing sheets (20g)","Thrissur / Ernakulam steel depots","₹95–115/sq.ft installed","Get price with installation labour"],
        ["Egg collection trays (30-egg)","Wholesale: Ernakulam markets","₹2.50–3.00/tray","Order 500 trays minimum"],
        ["Digital weighing scale","Amazon / local electronics shop","₹600–900","Accuracy to 1g for egg grading"],
        ["Generator (3.5 KVA diesel)","Kochi / Thrissur dealers: Honda, Mahindra","₹38,000–45,000","Inverter type recommended for fuel savings"],
        ["Egg stamp (date)","Indiamart: search 'egg date stamp'","₹800–1,200","Essential for premium branding"],
        ["Disinfectant (Virkon S)","Veterinary supply shops","₹850 per 1 kg pack","Lasts 3–4 months for 500-bird farm"],
        ["Feed storage bins","Local fabricator / plastic drum supplier","₹2,000–4,000 for 500 kg capacity","Rodent-proof, moisture-proof essential"],
      ]).map(([item,source,cost,notes],i)=>new TableRow({children:[
        dC(item,  1600,i%2===0?LTEAL:GREY,AlignmentType.LEFT,true,TEAL),
        dC(source,2000,i%2===0?WHITE:GREY),
        dC(cost,  2400,i%2===0?WHITE:GREY,AlignmentType.CENTER,true,GREEN),
        dC(notes, 3360,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 17: RISK MANAGEMENT
// ═══════════════════════════════════════════════════
function riskSection(){ return [
  h1("17. RISK ANALYSIS & CONTINGENCY PLAN"),
  spacer(1),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[2000,1400,1400,2800,1760],
    rows:[
      new TableRow({children:[hC("Risk",2000,RED),hC("Probability",1400,RED),hC("Financial Impact",1400,RED),hC("Mitigation",2800,RED),hC("Emergency Action",1760,RED)]}),
      ...([
        ["Ranikhet outbreak","Low (with vaccination)","₹2–5L flock loss","Strict vaccination schedule; biosecurity; emergency vet contact empanelled","Immediate vet call; isolate affected birds; emergency vaccination"],
        ["Feed price spike (+20%)","Medium","₹10,000–12,000/month extra","Bulk purchase 3 months stock; rice bran substitution; buying group","Pass 50% to buyers; absorb rest temporarily"],
        ["Egg price crash (-20%)","Medium (seasonal)","₹15,000–20,000/month less","40% direct customers insulate against NECC fluctuation","Shift more eggs to direct; approach new institutional buyers"],
        ["Power failure (extended)","High in Kerala","Heat stress; production drop 15–20%","Generator mandatory (already in plan); diesel stock for 72 hours","Start generator within 30 min of grid failure"],
        ["Key buyer drops out","Medium","₹8,000–15,000/month","No single buyer >25% of output; maintain waiting list","Activate waiting list immediately; temporary NECC sale"],
        ["Monsoon flooding","Low–Medium (site dependent)","Shed damage; bird loss","Elevated plinth 2.5 feet; proper drainage; site selection priority","Emergency evacuation plan for birds; insurance"],
        ["Theft or mortality spike","Low","₹10,000–50,000","CCTV at shed entry (₹3,000); daily bird count; visitor restriction","FIR if theft; vet if mortality; review CCTV footage"],
        ["Family health emergency","Medium (life of farmer)","Variable","Farm management SOP documented; train family member","Hire temporary help (₹500/day) for 2–4 weeks"],
      ]).map(([risk,prob,imp,mit,action],i)=>new TableRow({children:[
        dC(risk,  2000,i%2===0?WHITE:GREY,AlignmentType.LEFT,false,RED),
        dC(prob,  1400,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(imp,   1400,i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(mit,   2800,i%2===0?WHITE:GREY),
        dC(action,1760,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 18: SETUP TIMELINE
// ═══════════════════════════════════════════════════
function timelineSection(){ return [
  h1("18. STEP-BY-STEP SETUP TIMELINE", ORANGE),
  spacer(1),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1400,2000,6160],
    rows:[
      new TableRow({children:[hC("Timeline",1400,ORANGE),hC("Milestone",2000,ORANGE),hC("Detailed Actions",6160,ORANGE)]}),
      ...([
        ["Month 1\nWeek 1–2","Site finalisation","Select site (check drainage, distance, road, water). Get soil test. Negotiate land purchase/lease. Apply for Panchayat NOC."],
        ["Month 1\nWeek 3–4","Financial approvals","Meet bank — submit DPR + documents for NABARD loan. Apply for SFAC/SFAC scheme simultaneously. Get NABARD application number."],
        ["Month 2","Construction begins","Levelling and plinth work. Foundation. Borewell drilling (get water test done). Start buying cage quotes."],
        ["Month 3\nWeek 1–2","Structural work","Brick walls up. Roof structure erected. GI sheets fixed. False ceiling (Calicut tile) installed."],
        ["Month 3\nWeek 3–4","Infrastructure fit-out","Cages installed and aligned. Nipple drinker system plumbed. Electrical wiring, fans, lighting done. Generator installed."],
        ["Month 4\nWeek 1","Final preparations","Water tank filled, pump tested. All nipples checked for flow. Disinfect shed with Formalin (spray entire interior, seal 24 hours)."],
        ["Month 4\nWeek 2","Farm registration","Register with District Animal Husbandry Office. Apply for FSSAI Basic/State license online. Open farm bank account."],
        ["Month 4\nWeek 3","Order pullets","Book 500 BV-380 pullets from hatchery. Arrival date 1 week out. Buy 2 months feed stock. Order vaccines from vet."],
        ["Month 4\nWeek 4","Pullets arrive","Receive and house birds. Feed and fresh water immediately. Administer electrolyte solution for 48 hours. Record arrival count."],
        ["Month 5","Early production","Production starts slowly. Lighting: 16 hours. Vaccination: Ranikhet booster (Day 3–5 after arrival). Build buyer network intensively."],
        ["Month 6–8","Peak ramp-up","Production hits 75–85%. Activate all buyer channels. Launch WhatsApp subscription programme. Collect eggs 3x daily."],
        ["Month 9–18","Stable operations","Maintain records daily. Monthly financial review. NABARD subsidy credited at Month 24. Plan flock expansion at Month 16."],
        ["Month 18–20","Depletion","Contact spent hen buyers 8 weeks prior. Sell all birds at ₹200–250 each. Clean and fumigate shed. 3-week rest before new batch."],
      ]).map(([time,mile,actions],i)=>new TableRow({children:[
        dC(time,   1400,i%2===0?LORANG:GREY,AlignmentType.CENTER,true,ORANGE),
        dC(mile,   2000,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,DGREY),
        dC(actions,6160,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// SEC 19: SUCCESS FACTORS & SCALING PATH
// ═══════════════════════════════════════════════════
function successSection(){ return [
  h1("19. KEY SUCCESS FACTORS & SCALING ROADMAP", GREEN),
  spacer(1),
  h2("The 10 Non-Negotiable Success Factors"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[640,3600,5120],
    rows:[
      new TableRow({children:[hC("#",640),hC("Success Factor",3600),hC("What It Means in Practice",5120)]}),
      ...([
        ["1","Buyer network built BEFORE production","Have 250+ eggs/day committed orders before birds arrive. No exception."],
        ["2","Never miss a delivery","Delivery consistency = loyalty. One missed delivery loses a buyer 30% of the time."],
        ["3","Daily record-keeping","Lay rate, mortality, feed, sales — recorded every day. Review every Sunday."],
        ["4","Strict vaccination schedule","Miss one Ranikhet booster = potential total flock loss. Set calendar reminders 3 months out."],
        ["5","Quality feed — no compromise","Feed is 70% of your cost AND 70% of your output. Never buy from unknown, cheap sources."],
        ["6","Water quality maintained","Test every 6 months. Chlorinate or UV-treat continuously. Flush lines weekly."],
        ["7","16-hour lighting discipline","Never drop below 14 hours once birds start laying. Timer switch prevents human error."],
        ["8","Manure managed as income","Sell or compost manure from Month 1 — it adds ₹5,000–8,000/month."],
        ["9","Direct channel constantly growing","Every month: add 5 new direct subscribers. 60 months = 300 direct customers = price power."],
        ["10","Reserve fund maintained","Keep 1 month operating cost (₹70,000) as emergency reserve at all times."],
      ]).map(([n,factor,practice],i)=>new TableRow({children:[
        dC(n,      640,i%2===0?LGREEN:GREY,AlignmentType.CENTER,true,GREEN),
        dC(factor, 3600,i%2===0?WHITE:GREY,AlignmentType.LEFT,true,DGREY),
        dC(practice,5120,i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  h2("Scaling Roadmap: 500 → 2,200 Birds"),
  new Table({
    width:{size:FULL_W,type:WidthType.DXA},columnWidths:[1400,1400,1600,1560,4400],
    rows:[
      new TableRow({children:[hC("Phase",1400),hC("Timeline",1400),hC("Flock",1600),hC("Net/Month",1560),hC("Key Condition to Proceed",4400)]}),
      ...([
        ["Phase 1","Months 1–18","500 birds","₹7,608–28,580","Complete BEFORE scaling: 80+ WhatsApp subscribers, 5+ B2B accounts, 70%+ lay rate achieved, reserve fund ₹70,000"],
        ["Phase 2","Months 18–30","1,000 birds","₹35,000–55,000","Add 500 more BV-380. New shed bay or extend. Use profits + saved capital. NABARD subsidy credited."],
        ["Phase 3","Months 30–42","2,000 + 200 nattu","₹1,60,000–2,12,000","New shed, nattu mutta unit, manure business, packaged brand. 300+ direct subscribers."],
        ["Phase 4","Month 42+","Franchise / training","₹2,50,000+","Start training other farmers. Become an anchor farm for a cluster. Consulting income."],
      ]).map(([ph,time,flock,net,cond],i)=>new TableRow({children:[
        dC(ph,  1400,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER,true,i===0?GREEN:DGREY),
        dC(time,1400,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(flock,1600,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.CENTER),
        dC(net, 1560,i===0?LGREEN:i%2===0?WHITE:GREY,AlignmentType.Center,true,i===0?GREEN:DGREY),
        dC(cond,4400,i===0?LGREEN:i%2===0?WHITE:GREY),
      ]})),
    ],
  }),
  spacer(1),
  callout("🎯  THE BOTTOM LINE: A well-run 500-bird farm in Kerala is a genuine business, not a side income. Year 1 is the learning year — expect modest profits while buyer network builds. Year 2 is where the real money starts. Year 3+ with 2,000+ birds is where you change your life.", LGREEN, GREEN, true),
  spacer(2),
]}

// ═══════════════════════════════════════════════════
// DOCUMENT ASSEMBLY
// ═══════════════════════════════════════════════════
const doc = new Document({
  numbering:{config:[
    {reference:"bullets",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,
      style:{paragraph:{indent:{left:480,hanging:240}}}}]},
  ]},
  styles:{
    default:{document:{run:{font:"Arial",size:22}}},
    paragraphStyles:[
      {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:34,bold:true,font:"Arial",color:WHITE},
        paragraph:{spacing:{before:440,after:200},outlineLevel:0}},
      {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:27,bold:true,font:"Arial",color:GREEN},
        paragraph:{spacing:{before:300,after:100},outlineLevel:1}},
      {id:"Heading3",name:"Heading 3",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:23,bold:true,font:"Arial",color:GOLD},
        paragraph:{spacing:{before:200,after:80},outlineLevel:2}},
    ],
  },
  sections:[{
    properties:{page:{size:{width:12240,height:15840},margin:{top:1080,right:1080,bottom:1080,left:1080}}},
    headers:{default:new Header({children:[new Paragraph({
      children:[
        new TextRun({text:"Layer Poultry Farming — 500 Birds | Expanded Project Report",font:"Arial",size:18,bold:true,color:GREEN}),
        new TextRun({text:"\t",font:"Arial"}),
        new TextRun({text:"Kerala Conditions | 19 Sections | Full Business Plan",font:"Arial",size:18,color:DGREY,italics:true}),
      ],
      tabStops:[{type:TabStopType.RIGHT,position:TabStopPosition.MAX}],
      border:{bottom:{style:BorderStyle.SINGLE,size:4,color:GREEN,space:1}},
      spacing:{before:0,after:80},
    })]})},
    footers:{default:new Footer({children:[new Paragraph({
      children:[
        new TextRun({text:"© Kerala Layer Farm Series — Expanded Edition",font:"Arial",size:16,color:DGREY}),
        new TextRun({text:"\t",font:"Arial"}),
        new TextRun({children:["Page ",{type:"page"}],font:"Arial",size:16,color:DGREY}),
      ],
      tabStops:[{type:TabStopType.RIGHT,position:TabStopPosition.MAX}],
      border:{top:{style:BorderStyle.SINGLE,size:4,color:GREEN,space:1}},
      spacing:{before:80,after:0},
    })]})},
    children:[
      ...cover(),
      ...execSummary(),        pbr(),
      ...marketContextSection(), pbr(),
      ...siteSelectionSection(), pbr(),
      ...breedSection(),       pbr(),
      ...housingSection(),     pbr(),
      ...waterSection(),       pbr(),
      ...feedSection(),        pbr(),
      ...flockMgmtSection(),   pbr(),
      ...diseaseSection(),     pbr(),
      ...wasteSection(),       pbr(),
      ...labourSection(),      pbr(),
      ...financialSection(),   pbr(),
      ...nabardSection(),      pbr(),
      ...marketingSection(),   pbr(),
      ...schemesSection(),     pbr(),
      ...supplierSection(),    pbr(),
      ...riskSection(),        pbr(),
      ...timelineSection(),    pbr(),
      ...successSection(),
    ],
  }],
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync("/home/claude/Kerala_500Bird_Expanded_Report.docx",buf);
  console.log("Done");
});
