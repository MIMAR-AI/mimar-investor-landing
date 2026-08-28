import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "file:///C:/Users/Ahmad%20AbdulMajeed/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const outputDir = path.resolve("D:/Develop/MIMAR-AI/outputs/mimar-financial-model-2026-08-28");
await fs.mkdir(outputDir, { recursive: true });

const wb = Workbook.create();
const summary = wb.worksheets.add("Executive Summary");
const assumptions = wb.worksheets.add("Assumptions");
const core = wb.worksheets.add("Core Monthly");
const engage = wb.worksheets.add("Engage Monthly");
const scenarios = wb.worksheets.add("Scenarios");
const checks = wb.worksheets.add("Checks");
const sources = wb.worksheets.add("Sources");

const ink = "#F4F7F8";
const muted = "#A2ACB6";
const accent = "#9EE8EF";
const bg = "#0D0F13";
const surface = "#15191F";
const inputFill = "#E8F5FF";
const inputFont = "#0070C0";
const calcFill = "#F3F4F6";
const green = "#DCFCE7";
const red = "#FEE2E2";

function setTitle(sheet, title, subtitle, endCol = "H") {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${endCol}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = { fill: bg, font: { color: ink, bold: true, size: 18 }, rowHeight: 34 };
  sheet.getRange(`A2:${endCol}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format = { fill: bg, font: { color: muted, italic: true, size: 10 }, rowHeight: 26, wrapText: true };
}

function styleHeader(range) {
  range.format = {
    fill: surface,
    font: { color: accent, bold: true, size: 10 },
    rowHeight: 24,
    borders: { bottom: { style: "thin", color: accent } },
    wrapText: true,
  };
}

function currency(range) { range.format.numberFormat = "$#,##0;[Red]($#,##0);-"; }
function currency1(range) { range.format.numberFormat = "$#,##0.0;[Red]($#,##0.0);-"; }
function pct(range) { range.format.numberFormat = "0.0%"; }

// Assumptions
setTitle(assumptions, "MIMAR 3-Year Revenue Model — Assumptions", "Blue cells are editable inputs. Validation status distinguishes published inputs from assumptions that require paid-cohort evidence.", "H");
assumptions.getRange("A4:H4").values = [["Driver", "Y1", "Y2", "Y3", "Unit", "Evidence type", "Validation status", "Source / rationale"]];
styleHeader(assumptions.getRange("A4:H4"));
const assumptionRows = [
  ["CORE — ending individual paid users",120,600,1800,"users","Management target",25,"MIMAR business plan; no observed cohort data yet"],
  ["CORE — blended individual ARPU",18,21,23,"$/month","Published price + mix assumption",60,"MIMAR prices: Plus $10, Pro $25; mix not yet observed"],
  ["CORE — ending team accounts",25,85,230,"accounts","Management target",25,"MIMAR business plan; founder-led sales assumption"],
  ["CORE — average seats per team",8,10,12,"seats/account","Management assumption",30,"Target teams are 5–50 people; seat adoption untested"],
  ["CORE — blended team seat price",22,24,26,"$/seat/month","Published price + mix assumption",65,"MIMAR Team Plus $12 and Pro $30"],
  ["CORE — ending enterprise contracts",1,4,14,"contracts","Management target",20,"No signed enterprise contracts included"],
  ["CORE — enterprise ACV",18000,20000,24000,"$/year","Market-anchored assumption",40,"Custom annual contract; requires security/SSO/SLA proof"],
  ["CORE — ending active API wallets",30,170,630,"wallets","Management target",20,"API wallet is not yet GA; adoption unvalidated"],
  ["CORE — average API spend",35,75,95,"$/wallet/month","Usage assumption",20,"Prepaid token pricing is set; consumption mix is unknown"],
  ["CORE — onboarding fee per new B2B account",1000,900,750,"$/new account","Published range + mix",50,"Business plan range $1.5k–$5k; model discounts as onboarding becomes repeatable"],
  ["CORE — COGS as % revenue",0.30,0.25,0.23,"%","Management target",35,"Includes inference, hosting, storage and direct support"],
  ["",null,null,null,"","",null,""],
  ["ENGAGE — ending Starter accounts",16,45,85,"accounts","Management target",25,"Bottom-up ramp; no observed retention yet"],
  ["ENGAGE — Starter price",99,99,99,"$/month","Published test price",75,"MIMAR Engage plan"],
  ["ENGAGE — ending Growth accounts",16,55,165,"accounts","Management target",25,"Bottom-up ramp; no observed retention yet"],
  ["ENGAGE — Growth price",249,249,249,"$/month","Published test price",75,"MIMAR Engage plan"],
  ["ENGAGE — ending Scale accounts",3,17,43,"accounts","Management target",20,"Higher-support segment; unvalidated"],
  ["ENGAGE — Scale price",599,599,599,"$/month","Published test price",70,"MIMAR Engage plan"],
  ["ENGAGE — ending Enterprise accounts",1,3,7,"accounts","Management target",15,"No signed contracts; procurement/security risk"],
  ["ENGAGE — Enterprise monthly ARPA",1000,1000,1000,"$/month","Published floor",50,"$12k minimum ACV in Engage plan"],
  ["ENGAGE — usage/add-on revenue",0.00,0.05,0.08,"% of subscription","Management assumption",20,"Channel fees excluded; only MIMAR overages/add-ons"],
  ["ENGAGE — standard onboarding fee",600,500,400,"$/new account","Published range + mix",50,"Plan range $500–$1.5k; declines with templates/self-serve"],
  ["ENGAGE — enterprise onboarding fee",3000,4000,5000,"$/new account","Management assumption",35,"Integration/security scope"],
  ["ENGAGE — COGS as % revenue",0.20,0.21,0.22,"%","Management target",40,"Excludes pass-through Meta fees from both revenue and margin"],
];
for (const row of assumptionRows) {
  if (typeof row[6] === "number") {
    row[6] = row[6] >= 70 ? "Published input" : row[6] >= 50 ? "Market anchored" : row[6] >= 30 ? "Model assumption" : "Validate with cohorts";
  }
}
assumptions.getRange(`A5:H${4 + assumptionRows.length}`).values = assumptionRows;
assumptions.getRange("B5:D28").format = { fill: inputFill, font: { color: inputFont } };
currency(assumptions.getRange("B11:D11"));
currency(assumptions.getRange("B24:D24"));
currency(assumptions.getRange("B26:D27"));
pct(assumptions.getRange("B15:D15"));
pct(assumptions.getRange("B25:D25"));
pct(assumptions.getRange("B28:D28"));
assumptions.getRange("A30:H30").values = [["SCENARIO MULTIPLIERS",null,null,null,null,null,null,"Applied to Base revenue; ranges are directional until conversion and retention cohorts exist."]];
styleHeader(assumptions.getRange("A30:H30"));
assumptions.getRange("A31:D34").values = [
  ["Scenario","Y1","Y2","Y3"],
  ["Conservative",0.65,0.55,0.45],
  ["Base",1,1,1],
  ["Upside",1.25,1.45,1.60],
];
styleHeader(assumptions.getRange("A31:D31"));
pct(assumptions.getRange("B32:D34"));
assumptions.getRange("B32:D34").format.fill = inputFill;
assumptions.freezePanes.freezeRows(4);
assumptions.getRange("A:H").format.columnWidth = 14;
assumptions.getRange("A:A").format.columnWidth = 43;
assumptions.getRange("E:E").format.columnWidth = 18;
assumptions.getRange("F:F").format.columnWidth = 25;
assumptions.getRange("G:G").format.columnWidth = 23;
assumptions.getRange("H:H").format.columnWidth = 58;

// Core monthly model
setTitle(core, "MIMAR Core — Monthly Revenue Build", "36-month base case. Customer counts interpolate between annual ending targets; this is a planning model, not cohort evidence.", "S");
core.getRange("A4:S4").values = [["Month","Year","Individual users","Ind. ARPU","Individual rev.","Team accounts","Seats / team","Seat price","Team rev.","Enterprise accounts","Enterprise ACV","Enterprise rev.","API wallets","Spend / wallet","API rev.","New B2B accts","Onboarding fee","Onboarding rev.","CORE TOTAL"]];
styleHeader(core.getRange("A4:S4"));
const coreRows = [];
for (let m=1; m<=36; m++) {
  const y = Math.ceil(m/12);
  const j = ((m-1)%12)+1;
  const r = 4+m;
  const col = ["B","C","D"][y-1];
  const prevCol = y===1 ? null : ["B","C"][y-2];
  const ramp = (assRow) => y===1
    ? `=Assumptions!$${col}$${assRow}*${j}/12`
    : `=Assumptions!$${prevCol}$${assRow}+(Assumptions!$${col}$${assRow}-Assumptions!$${prevCol}$${assRow})*${j}/12`;
  coreRows.push([new Date(2026, 8 + m - 1, 1), y, ramp(5), `=Assumptions!$${col}$6`, `=C${r}*D${r}`, ramp(7), `=Assumptions!$${col}$8`, `=Assumptions!$${col}$9`, `=F${r}*G${r}*H${r}`, ramp(10), `=Assumptions!$${col}$11`, `=J${r}*K${r}/12`, ramp(12), `=Assumptions!$${col}$13`, `=M${r}*N${r}`, m===1?`=F${r}+J${r}`:`=MAX(0,F${r}+J${r}-F${r-1}-J${r-1})`, `=Assumptions!$${col}$14`, `=P${r}*Q${r}`, `=SUM(E${r},I${r},L${r},O${r},R${r})`]);
}
core.getRange("A5:S40").values = coreRows.map(row => row.map(v => (typeof v === "string" && v.startsWith("=")) ? null : v));
core.getRange("C5:S40").formulas = coreRows.map(row => row.slice(2).map(v => (typeof v === "string" && v.startsWith("=")) ? v : null));
core.getRange("A5:A40").format.numberFormat = "mmm-yy";
core.getRange("C5:C40").format.numberFormat = "0.0";
core.getRange("F5:G40").format.numberFormat = "0.0";
core.getRange("J5:J40").format.numberFormat = "0.00";
core.getRange("M5:M40").format.numberFormat = "0.0";
core.getRange("P5:P40").format.numberFormat = "0.00";
currency1(core.getRange("D5:E40")); currency1(core.getRange("H5:I40")); currency1(core.getRange("K5:L40")); currency1(core.getRange("N5:O40")); currency1(core.getRange("Q5:S40"));
core.freezePanes.freezeRows(4); core.freezePanes.freezeColumns(2);
core.getRange("A:S").format.columnWidth = 14; core.getRange("A:A").format.columnWidth = 12; core.getRange("S:S").format.columnWidth = 17;

// Engage monthly model
setTitle(engage, "MIMAR Engage — Monthly Revenue Build", "Subscriptions, MIMAR usage/add-ons and onboarding are separate. Meta/channel fees are pass-through and excluded.", "R");
engage.getRange("A4:R4").values = [["Month","Year","Starter accts","Starter rev.","Growth accts","Growth rev.","Scale accts","Scale rev.","Enterprise accts","Enterprise rev.","Usage %","Usage/add-ons","New standard","Std onboarding fee","New enterprise","Ent onboarding fee","Onboarding rev.","ENGAGE TOTAL"]];
styleHeader(engage.getRange("A4:R4"));
const engageRows = [];
for (let m=1; m<=36; m++) {
  const y=Math.ceil(m/12), j=((m-1)%12)+1, r=4+m;
  const col=["B","C","D"][y-1], prevCol=y===1?null:["B","C"][y-2];
  const ramp=(assRow)=>y===1?`=Assumptions!$${col}$${assRow}*${j}/12`:`=Assumptions!$${prevCol}$${assRow}+(Assumptions!$${col}$${assRow}-Assumptions!$${prevCol}$${assRow})*${j}/12`;
  const prev=(c)=>m===1?"0":`${c}${r-1}`;
  engageRows.push([new Date(2026,8 + m - 1,1),y,ramp(17),`=C${r}*Assumptions!$${col}$18`,ramp(19),`=E${r}*Assumptions!$${col}$20`,ramp(21),`=G${r}*Assumptions!$${col}$22`,ramp(23),`=I${r}*Assumptions!$${col}$24`,`=Assumptions!$${col}$25`,`=SUM(D${r},F${r},H${r},J${r})*K${r}`,`=MAX(0,C${r}+E${r}+G${r}-${prev("C")}-${prev("E")}-${prev("G")})`,`=Assumptions!$${col}$26`,`=MAX(0,I${r}-${prev("I")})`,`=Assumptions!$${col}$27`,`=M${r}*N${r}+O${r}*P${r}`,`=SUM(D${r},F${r},H${r},J${r},L${r},Q${r})`]);
}
engage.getRange("A5:R40").values = engageRows.map(row=>row.map(v=>(typeof v==="string"&&v.startsWith("="))?null:v));
engage.getRange("C5:R40").formulas = engageRows.map(row=>row.slice(2).map(v=>(typeof v==="string"&&v.startsWith("="))?v:null));
engage.getRange("A5:A40").format.numberFormat="mmm-yy";
engage.getRange("C5:C40").format.numberFormat="0.0"; engage.getRange("E5:E40").format.numberFormat="0.0"; engage.getRange("G5:G40").format.numberFormat="0.0"; engage.getRange("I5:I40").format.numberFormat="0.00";
pct(engage.getRange("K5:K40"));
engage.getRange("M5:M40").format.numberFormat="0.00"; engage.getRange("O5:O40").format.numberFormat="0.00";
currency1(engage.getRange("D5:D40")); currency1(engage.getRange("F5:F40")); currency1(engage.getRange("H5:H40")); currency1(engage.getRange("J5:J40")); currency1(engage.getRange("L5:L40")); currency1(engage.getRange("N5:N40")); currency1(engage.getRange("P5:R40"));
engage.freezePanes.freezeRows(4); engage.freezePanes.freezeColumns(2);
engage.getRange("A:R").format.columnWidth=14; engage.getRange("A:A").format.columnWidth=12; engage.getRange("R:R").format.columnWidth=17;

// Executive summary
setTitle(summary, "MIMAR Portfolio — Investor Revenue Model", "Base case | 36 months from Sep 2026 | Core and Engage shown separately | Forecasts are not historical results", "L");
summary.getRange("A4:D4").values = [["Revenue stream","Y1","Y2","Y3"]]; styleHeader(summary.getRange("A4:D4"));
const sumBlocks = [
  ["CORE — Individuals", "=SUM('Core Monthly'!E5:E16)", "=SUM('Core Monthly'!E17:E28)", "=SUM('Core Monthly'!E29:E40)"],
  ["CORE — Team seats", "=SUM('Core Monthly'!I5:I16)", "=SUM('Core Monthly'!I17:I28)", "=SUM('Core Monthly'!I29:I40)"],
  ["CORE — Enterprise", "=SUM('Core Monthly'!L5:L16)", "=SUM('Core Monthly'!L17:L28)", "=SUM('Core Monthly'!L29:L40)"],
  ["CORE — API wallet", "=SUM('Core Monthly'!O5:O16)", "=SUM('Core Monthly'!O17:O28)", "=SUM('Core Monthly'!O29:O40)"],
  ["CORE — Onboarding", "=SUM('Core Monthly'!R5:R16)", "=SUM('Core Monthly'!R17:R28)", "=SUM('Core Monthly'!R29:R40)"],
  ["CORE TOTAL", "=SUM(B5:B9)", "=SUM(C5:C9)", "=SUM(D5:D9)"],
  ["",null,null,null],
  ["ENGAGE — Starter", "=SUM('Engage Monthly'!D5:D16)", "=SUM('Engage Monthly'!D17:D28)", "=SUM('Engage Monthly'!D29:D40)"],
  ["ENGAGE — Growth", "=SUM('Engage Monthly'!F5:F16)", "=SUM('Engage Monthly'!F17:F28)", "=SUM('Engage Monthly'!F29:F40)"],
  ["ENGAGE — Scale", "=SUM('Engage Monthly'!H5:H16)", "=SUM('Engage Monthly'!H17:H28)", "=SUM('Engage Monthly'!H29:H40)"],
  ["ENGAGE — Enterprise", "=SUM('Engage Monthly'!J5:J16)", "=SUM('Engage Monthly'!J17:J28)", "=SUM('Engage Monthly'!J29:J40)"],
  ["ENGAGE — Usage/add-ons", "=SUM('Engage Monthly'!L5:L16)", "=SUM('Engage Monthly'!L17:L28)", "=SUM('Engage Monthly'!L29:L40)"],
  ["ENGAGE — Onboarding", "=SUM('Engage Monthly'!Q5:Q16)", "=SUM('Engage Monthly'!Q17:Q28)", "=SUM('Engage Monthly'!Q29:Q40)"],
  ["ENGAGE TOTAL", "=SUM(B12:B17)", "=SUM(C12:C17)", "=SUM(D12:D17)"],
  ["",null,null,null],
  ["PORTFOLIO REVENUE", "=B10+B18", "=C10+C18", "=D10+D18"],
  ["Ending MRR (recurring only)", "='Core Monthly'!S16-'Core Monthly'!R16+'Engage Monthly'!R16-'Engage Monthly'!Q16", "='Core Monthly'!S28-'Core Monthly'!R28+'Engage Monthly'!R28-'Engage Monthly'!Q28", "='Core Monthly'!S40-'Core Monthly'!R40+'Engage Monthly'!R40-'Engage Monthly'!Q40"],
  ["Exit ARR", "=B21*12", "=C21*12", "=D21*12"],
  ["Recurring revenue mix", "=(B20-B9-B17)/B20", "=(C20-C9-C17)/C20", "=(D20-D9-D17)/D20"],
  ["Core share", "=B10/B20", "=C10/C20", "=D10/D20"],
  ["Engage share", "=B18/B20", "=C18/C20", "=D18/D20"],
  ["Gross profit", "=B10*(1-Assumptions!B15)+B18*(1-Assumptions!B28)", "=C10*(1-Assumptions!C15)+C18*(1-Assumptions!C28)", "=D10*(1-Assumptions!D15)+D18*(1-Assumptions!D28)"],
  ["Blended gross margin", "=B26/B20", "=C26/C20", "=D26/D20"],
  ["Model status", "Planning case", "Planning case", "Planning case"],
];
summary.getRange("A5:D28").values = sumBlocks.map(row=>row.map(v=>(typeof v==="string"&&v.startsWith("="))?null:v));
summary.getRange("B5:D28").formulas = sumBlocks.map(row=>row.slice(1).map(v=>(typeof v==="string"&&v.startsWith("="))?v:null));
currency(summary.getRange("B5:D22")); currency(summary.getRange("B26:D26"));
pct(summary.getRange("B23:D25")); pct(summary.getRange("B27:D27"));
summary.getRange("A10:D10").format = { fill: "#DFF7FA", font: { bold:true, color:"#12323A" }, borders:{ top:{style:"thin",color:accent}, bottom:{style:"thin",color:accent} } };
summary.getRange("A18:D18").format = { fill: "#E8EEFF", font: { bold:true, color:"#172554" }, borders:{ top:{style:"thin",color:"#93C5FD"}, bottom:{style:"thin",color:"#93C5FD"} } };
summary.getRange("A20:D20").format = { fill: bg, font: { bold:true, color:accent, size:12 }, borders:{ top:{style:"medium",color:accent}, bottom:{style:"medium",color:accent} } };
summary.getRange("A30:D30").values = [["Forecast range","Y1","Y2","Y3"]]; styleHeader(summary.getRange("A30:D30"));
summary.getRange("A31:D33").values = [["Conservative",null,null,null],["Base",null,null,null],["Upside",null,null,null]];
summary.getRange("B31:D33").formulas = [
  ["=B20*Assumptions!B32","=C20*Assumptions!C32","=D20*Assumptions!D32"],
  ["=B20","=C20","=D20"],
  ["=B20*Assumptions!B34","=C20*Assumptions!C34","=D20*Assumptions!D34"],
];
currency(summary.getRange("B31:D33"));
summary.getRange("F4:I4").values = [["Year","Core","Engage","Portfolio"]]; styleHeader(summary.getRange("F4:I4"));
summary.getRange("F5:F7").values = [["Y1"],["Y2"],["Y3"]];
summary.getRange("G5:I7").formulas = [["=B10","=B18","=B20"],["=C10","=C18","=C20"],["=D10","=D18","=D20"]];
currency(summary.getRange("G5:I7"));
const chart = summary.charts.add("bar", summary.getRange("F4:H7"));
chart.title = "Revenue mix by product ($)"; chart.hasLegend = true; chart.setPosition("F9","L25"); chart.yAxis = { numberFormatCode: "$0.0,,\"M\"" };
summary.getRange("F27:L30").merge(); summary.getRange("F27").values = [["Forecast note: this is a driver-based planning case, not company guidance. Conservative and upside ranges are shown because paid conversion, retention, seat expansion and usage cohorts are still being established."]];
summary.getRange("F27:L30").format = { fill: "#FFF7ED", font:{ color:"#7C2D12", italic:true }, wrapText:true, verticalAlignment:"center" };
summary.getRange("A:L").format.columnWidth=15; summary.getRange("A:A").format.columnWidth=32; summary.getRange("F:F").format.columnWidth=12;
summary.freezePanes.freezeRows(4);

// Scenarios
setTitle(scenarios, "Scenario View", "The range widens by year because conversion, retention and expansion are not yet observed.", "G");
scenarios.getRange("A4:D4").values=[["Scenario","Y1","Y2","Y3"]]; styleHeader(scenarios.getRange("A4:D4"));
scenarios.getRange("A5:A7").values=[["Conservative"],["Base"],["Upside"]];
scenarios.getRange("B5:D7").formulas=[["='Executive Summary'!B31","='Executive Summary'!C31","='Executive Summary'!D31"],["='Executive Summary'!B32","='Executive Summary'!C32","='Executive Summary'!D32"],["='Executive Summary'!B33","='Executive Summary'!C33","='Executive Summary'!D33"]];
currency(scenarios.getRange("B5:D7"));
scenarios.getRange("A10:G10").values=[["What changes","Conservative","Base","Upside","Evidence needed","Decision gate","Owner"]]; styleHeader(scenarios.getRange("A10:G10"));
scenarios.getRange("A11:G15").values=[
  ["Customer acquisition","Slower pilots and partner ramp","Plan targets","Faster referrals + self-serve","Qualified pipeline, win rate","5–7 paid conversions","Founder"],
  ["Retention","Higher early churn","Target churn after beta","Strong reuse/expansion","3+ monthly cohorts","GRR/NRR measured","Product/CS"],
  ["Team seat expansion","6–8 seats","8–12 seats","12+ seats","Invites and active-seat cohorts","2 workflows + 5 active users","Product"],
  ["API adoption","Delayed GA","Launch after reliability gate","Developer pull accelerates","Wallet activation and spend","20 retained wallets","Engineering"],
  ["Engage usage","Mostly base subscription","5–8% add-ons by Y2–Y3","10%+ add-ons","Inbound volume and overage","10 paid retained accounts","Engage GM"],
];
scenarios.getRange("A:G").format.columnWidth=22; scenarios.getRange("A:A").format.columnWidth=25; scenarios.getRange("E:G").format.columnWidth=28; scenarios.getRange("A10:G15").format.wrapText=true;

// Checks
setTitle(checks, "Model Checks", "PASS means the workbook reconciles mechanically; it does not validate market demand.", "D");
checks.getRange("A4:D4").values=[["Check","Status","Delta","Where to fix"]]; styleHeader(checks.getRange("A4:D4"));
checks.getRange("A5:A11").values=[["Portfolio equals Core + Engage — Y1"],["Portfolio equals Core + Engage — Y2"],["Portfolio equals Core + Engage — Y3"],["Core annual sums tie to monthly model"],["Engage annual sums tie to monthly model"],["Exit ARR exceeds annual revenue during ramp"],["Scenario ordering is logical"]];
checks.getRange("B5:B11").formulas=[["=IF(ABS('Executive Summary'!B20-'Executive Summary'!B10-'Executive Summary'!B18)<0.01,\"PASS\",\"FAIL\")"],["=IF(ABS('Executive Summary'!C20-'Executive Summary'!C10-'Executive Summary'!C18)<0.01,\"PASS\",\"FAIL\")"],["=IF(ABS('Executive Summary'!D20-'Executive Summary'!D10-'Executive Summary'!D18)<0.01,\"PASS\",\"FAIL\")"],["=IF(ABS(SUM('Core Monthly'!S5:S40)-SUM('Executive Summary'!B10:D10))<0.01,\"PASS\",\"FAIL\")"],["=IF(ABS(SUM('Engage Monthly'!R5:R40)-SUM('Executive Summary'!B18:D18))<0.01,\"PASS\",\"FAIL\")"],["=IF(AND('Executive Summary'!B22>'Executive Summary'!B20,'Executive Summary'!C22>'Executive Summary'!C20,'Executive Summary'!D22>'Executive Summary'!D20),\"PASS\",\"REVIEW\")"],["=IF(AND('Executive Summary'!B31<'Executive Summary'!B32,'Executive Summary'!B32<'Executive Summary'!B33,'Executive Summary'!D31<'Executive Summary'!D32,'Executive Summary'!D32<'Executive Summary'!D33),\"PASS\",\"FAIL\")"]];
checks.getRange("C5:C11").formulas=[["='Executive Summary'!B20-'Executive Summary'!B10-'Executive Summary'!B18"],["='Executive Summary'!C20-'Executive Summary'!C10-'Executive Summary'!C18"],["='Executive Summary'!D20-'Executive Summary'!D10-'Executive Summary'!D18"],["=SUM('Core Monthly'!S5:S40)-SUM('Executive Summary'!B10:D10)"],["=SUM('Engage Monthly'!R5:R40)-SUM('Executive Summary'!B18:D18)"],["='Executive Summary'!D22-'Executive Summary'!D20"],["='Executive Summary'!D33-'Executive Summary'!D31"]];
checks.getRange("D5:D11").values=[["Executive Summary / Core Monthly"],["Executive Summary / Core Monthly"],["Executive Summary / Core Monthly"],["Core Monthly"],["Engage Monthly"],["Assumptions: ending accounts and timing"],["Assumptions: scenario multipliers"]];
currency(checks.getRange("C5:C11")); checks.getRange("A:D").format.columnWidth=28; checks.getRange("A:A").format.columnWidth=42;

// Sources
setTitle(sources, "Sources and Benchmark Log", "External benchmarks anchor structure and price bands; they do not prove MIMAR customer demand.", "I");
sources.getRange("A4:I4").values=[["Item","Finding used","As of","Source type","Source","URL","Applies to","Model implication","Caveat"]]; styleHeader(sources.getRange("A4:I4"));
const sourceRows=[
  ["MIMAR Core prices","$10 Plus; $25 Pro; Team $12/$30; API prepaid wallet","2026-08-28","Management decision","MIMAR business plan","D:/Develop/MIMAR-AI/MIMAR_BUSINESS_PLAN_AR_2026.md","Core","Price inputs are relatively high-confidence","Customer mix and retention remain untested"],
  ["MIMAR Engage prices","Starter $99; Growth $249; Scale $599; Enterprise $12k ACV floor","2026-08-28","Management decision","MIMAR Engage plan","D:/Develop/MIMAR-AI/MIMAR_ENGAGE_BUSINESS_PLAN_AR_2026.md","Engage","Tier prices sit within adjacent-market range","Local willingness to pay requires pilots"],
  ["GitHub Copilot","Individual $10/$39/$100; Business $19 and Enterprise $39 per seat","2026-08-28","Official pricing","GitHub Docs","https://docs.github.com/en/copilot/get-started/plans","Core","Supports $10–$30 entry/seat band","Not a direct workflow-management competitor"],
  ["n8n","€20 Starter and €50 Pro, billed annually; execution-based pricing","2026-08-28","Official pricing","n8n","https://n8n.io/pricing/","Core","Workflow value can support subscription + usage","Different buyer and product scope"],
  ["Manychat","$29/$69/$139 annual-billed plans; active-contact limits","2026-08-28","Official pricing","Manychat","https://manychat.com/pricing","Engage","Shows low-end regional/self-serve anchor","Broader creator automation; not trust-first support"],
  ["respond.io","$79/$159/$279; WhatsApp fees excluded","2026-08-28","Official pricing","respond.io","https://respond.io/pricing","Engage","Supports $99/$249 workspace tiers and pass-through channel fees","AI outcome depth differs"],
  ["Gorgias AI Agent","About $0.90–$1 per automated resolution plus helpdesk subscription","2026-08-28","Official pricing","Gorgias","https://www.gorgias.com/pricing","Engage","Supports a usage/overage layer","E-commerce-centric and outcome-priced"],
  ["Intercom Fin","$0.99 resolution/procedure outcome; $9.99 qualification","2026-08-28","Official help center","Intercom","https://www.intercom.com/help/en/articles/8205718-fin-ai-agent-outcomes","Engage","Confirms measurable outcomes can carry usage pricing","MIMAR should not claim outcome pricing before attribution data"],
  ["Private SaaS growth","2026 survey median growth 22% across 1,000+ private B2B SaaS companies","2026-08-28","Industry benchmark","SaaS Capital","https://www.saas-capital.com/research/private-saas-company-growth-rate-benchmarks/","Portfolio","Use as mature-company context, not pre-PMF target","Sample generally more mature than MIMAR"],
  ["AI gross margin","BVP 2025: AI Shooting Stars averaged about 60% gross margin","2026-08-28","Industry benchmark","Bessemer Venture Partners","https://www.bvp.com/atlas/the-state-of-ai-2025","Portfolio","75–80% target is ambitious and must be measured","Selected high-performing AI cohort"],
  ["Pricing practice","Treat early pricing as a living experiment and test willingness to pay","2026-08-28","Growth guidance","Demand Curve","https://www.demandcurve.com/lessons/choose-a-pricing-model","Portfolio","Keep price confidence above volume confidence","Guidance, not market data"],
  ["Pitch clarity","Explain what the company does simply; market sizing can be bottom-up","2026-08-28","Founder guidance","Y Combinator","https://www.ycombinator.com/blog/how-to-pitch-your-company/","Investor page","Show drivers and caveats, not opaque headline totals","General pitch guidance"],
];
sources.getRange(`A5:I${4+sourceRows.length}`).values=sourceRows;
sources.getRange("A:I").format.columnWidth=20; sources.getRange("A:A").format.columnWidth=23; sources.getRange("B:B").format.columnWidth=52; sources.getRange("F:F").format.columnWidth=58; sources.getRange("H:I").format.columnWidth=40; sources.getRange(`A4:I${4+sourceRows.length}`).format.wrapText=true; sources.freezePanes.freezeRows(4);

// General styling
for (const sheet of [summary, assumptions, core, engage, scenarios, checks, sources]) {
  const used = sheet.getUsedRange();
  if (used) used.format.font = { name: "Aptos", size: 10 };
}

const exported = await SpreadsheetFile.exportXlsx(wb);
const outPath = path.join(outputDir, "MIMAR_Investor_Revenue_Model_2026-08-28.xlsx");
await exported.save(outPath);

for (const sheetName of ["Executive Summary","Assumptions","Core Monthly","Engage Monthly","Scenarios","Checks","Sources"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName.replaceAll(" ","_")}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const checkInspect = await wb.inspect({ kind: "region", sheetId: "Checks", range: "A4:D11", maxChars: 6000 });
const errorInspect = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, maxChars: 6000 });
console.log(JSON.stringify({ outPath, checks: checkInspect.ndjson, errors: errorInspect.ndjson }));
