const pptxgen = require('pptxgenjs');
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Hailemariam Mersha et al.';
pptx.subject = 'CS-UH 2219E Computational Social Science — Milestone 2';
pptx.title = 'GenAI Exposure and Full-Time to Involuntary Part-Time Transitions';
pptx.company = 'NYU Abu Dhabi';
pptx.lang = 'en-US';
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'en-US'
};
pptx.defineSlideMaster({
  title: 'MASTER',
  background: { color: 'FFFFFF' },
  objects: [
    { rect: { x: 0, y: 0, w: 13.333, h: 0.10, fill: { color: '6D28D9' }, line: { color: '6D28D9' } } },
    { text: { text: 'CS-UH 2219E · Computational Social Science · Milestone 2', options: { x: 0.55, y: 7.16, w: 6.3, h: 0.18, fontFace: 'Aptos', fontSize: 8.5, color: '7C3AED', margin: 0 } } },
    { text: { text: 'Source: IPUMS CPS Basic Monthly 2017–2025; team calculations. LM-AIOE merge pending.', options: { x: 6.15, y: 7.16, w: 6.63, h: 0.18, fontFace: 'Aptos', fontSize: 8, color: '9CA3AF', align: 'right', margin: 0 } } },
  ],
  slideNumber: { x: 12.85, y: 7.14, color: '7C3AED', fontSize: 8 }
});

const C = {
  violet: '6D28D9',
  violet2: '7C3AED',
  dark: '2E1065',
  lavender: 'EDE9FE',
  lavender2: 'F5F3FF',
  mid: 'A78BFA',
  gray: '6B7280',
  lightGray: 'E5E7EB',
  paleGray: 'F9FAFB',
  white: 'FFFFFF',
  redSoft: 'FEE2E2',
  red: 'B91C1C',
  greenSoft: 'ECFDF5',
  green: '047857',
};

function addTitle(slide, title, subtitle='') {
  slide.addText(title, { x: 0.65, y: 0.40, w: 12.0, h: 0.48, fontFace: 'Aptos Display', fontSize: 25, bold: true, color: C.dark, margin: 0, breakLine: false });
  if (subtitle) slide.addText(subtitle, { x: 0.68, y: 0.95, w: 11.8, h: 0.34, fontSize: 11.5, color: C.gray, margin: 0 });
}
function addPill(slide, text, x, y, w, fill=C.lavender, color=C.violet) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.36, rectRadius: 0.08, fill: { color: fill }, line: { color: fill } });
  slide.addText(text, { x: x+0.08, y: y+0.06, w: w-0.16, h: 0.22, fontSize: 10, bold: true, color, align: 'center', margin: 0 });
}
function statCard(slide, x, y, w, h, value, label, note='') {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: C.lavender2 }, line: { color: C.lightGray, width: 1 } });
  slide.addText(value, { x: x+0.18, y: y+0.16, w: w-0.36, h: 0.42, fontSize: 24, bold: true, color: C.violet, margin: 0 });
  slide.addText(label, { x: x+0.18, y: y+0.66, w: w-0.36, h: 0.35, fontSize: 11.5, bold: true, color: C.dark, margin: 0 });
  if (note) slide.addText(note, { x: x+0.18, y: y+1.06, w: w-0.36, h: h-1.18, fontSize: 9.5, color: C.gray, margin: 0 });
}
function bullet(slide, text, x, y, w, fs=15, color=C.dark) {
  slide.addText([{text:'• ',options:{bold:true,color:C.violet}},{text,options:{color}}], { x, y, w, h: 0.42, fontSize: fs, margin: 0, breakLine: false });
}
function addSectionLabel(slide, text, x, y, w=2.2) { addPill(slide, text.toUpperCase(), x, y, w, C.lavender, C.violet); }

// Slide 1
{
  const s = pptx.addSlide('MASTER');
  s.addText('Milestone 2', { x: 0.72, y: 1.15, w: 3.0, h: 0.42, fontSize: 18, bold: true, color: C.violet, margin: 0 });
  s.addText('From Full-Time to\nInvoluntary Part-Time?', { x: 0.70, y: 1.72, w: 7.7, h: 1.60, fontFace: 'Aptos Display', fontSize: 34, bold: true, color: C.dark, margin: 0, breakLine: false });
  s.addText('Building the longitudinal CPS dataset and validating the labor-market transition outcome before merging occupation-level language-model exposure.', { x: 0.73, y: 3.55, w: 7.25, h: 0.95, fontSize: 17, color: C.gray, margin: 0 });
  addPill(s, '2017–2025 Basic Monthly CPS', 0.73, 4.80, 2.55);
  addPill(s, '1.06M validated transitions', 3.48, 4.80, 2.55);
  addPill(s, '13,103 FT→PTER events', 6.23, 4.80, 2.30);
  s.addShape(pptx.ShapeType.roundRect, { x: 9.25, y: 1.35, w: 3.05, h: 4.55, rectRadius: 0.1, fill: { color: C.lavender2 }, line: { color: C.lavender } });
  s.addText('Research question', { x: 9.55, y: 1.70, w: 2.45, h: 0.38, fontSize: 15, bold: true, color: C.violet, margin: 0 });
  s.addText('Did workers in U.S. occupations with greater exposure to language-model capabilities become more likely to transition from full-time wage employment to part-time work for economic reasons after 2022?', { x: 9.52, y: 2.25, w: 2.45, h: 2.20, fontSize: 16, bold: true, color: C.dark, margin: 0.02, valign: 'mid' });
  s.addText('Today: validate the panel, outcome, and descriptive structure. Next: merge LM-AIOE and test exposure-specific divergence.', { x: 9.52, y: 4.75, w: 2.45, h: 0.72, fontSize: 10.5, color: C.gray, margin: 0 });
}

// Slide 2
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'What changed since Milestone 1?', 'The project has moved from a proposed design to a working longitudinal labor-market dataset.');
  const cols = [0.75, 4.55, 8.35];
  const titles = ['Then: proposed data', 'Now: working panel', 'Still missing: AI exposure'];
  const bodies = [
    ['IPUMS CPS identified', 'PTER outcome proposed', 'One-year linking strategy planned'],
    ['11.64M raw person-records', '1.06M validated one-year transitions', '13,103 FT→PTER events observed'],
    ['LM-AIOE workbook', 'Occupation crosswalk validation', 'Exposure × post-2022 analysis']
  ];
  for (let i=0;i<3;i++) {
    s.addShape(pptx.ShapeType.roundRect,{x:cols[i],y:1.65,w:3.35,h:4.65,rectRadius:0.08,fill:{color:i===1?C.lavender2:C.white},line:{color:i===1?C.violet:C.lightGray,width:i===1?1.5:1}});
    s.addText(titles[i],{x:cols[i]+0.22,y:1.95,w:2.9,h:0.45,fontSize:17,bold:true,color:i===1?C.violet:C.dark,margin:0});
    bodies[i].forEach((b,j)=>bullet(s,b,cols[i]+0.25,2.65+j*0.88,2.82,14));
  }
  s.addText('Milestone 2 question for the data section: does the panel behave sensibly enough to support the planned causal/associational analysis?', { x: 1.18, y: 6.45, w: 10.9, h: 0.48, fontSize: 15.5, bold: true, color: C.dark, align: 'center', margin: 0 });
}

// Slide 3
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Sample construction', 'A transparent funnel from raw CPS records to the one-year transition panel.');
  const stages = [
    ['11,635,165','Basic Monthly person-records'],
    ['6,849,416','Age 18–64'],
    ['4,344,203','At-work wage/salary'],
    ['3,456,082','Full-time wage/salary'],
    ['1,112,384','Eligible one-year baselines'],
    ['1,062,438','Validated one-year transitions'],
  ];
  const widths = [10.9,9.8,8.75,7.7,6.6,5.85];
  const startX = 1.2;
  stages.forEach((st,i)=>{
    const x = startX + (10.9-widths[i])/2;
    const y = 1.55 + i*0.82;
    s.addShape(pptx.ShapeType.chevron,{x,y,w:widths[i],h:0.62,fill:{color:i===5?C.violet:C.lavender2},line:{color:i===5?C.violet:C.lavender}});
    s.addText(st[0],{x:x+0.25,y:y+0.10,w:1.85,h:0.27,fontSize:17,bold:true,color:i===5?C.white:C.violet,margin:0});
    s.addText(st[1],{x:x+2.15,y:y+0.10,w:widths[i]-2.45,h:0.27,fontSize:13.5,bold:true,color:i===5?C.white:C.dark,margin:0});
  });
  s.addText('95.5% of eligible baselines are successfully observed in the same month one year later.', { x: 3.2, y: 6.53, w: 6.9, h: 0.38, fontSize: 15, bold: true, color: C.violet, align:'center', margin: 0 });
}

// Slide 4
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'How the one-year CPS panel is built', 'We use the validated longitudinal identifier and the CPS 4–8–4 rotation structure.');
  addSectionLabel(s,'Baseline',0.8,1.55,1.35);
  addSectionLabel(s,'Follow-up',8.55,1.55,1.45);
  s.addShape(pptx.ShapeType.roundRect,{x:0.85,y:2.15,w:3.9,h:2.65,rectRadius:0.08,fill:{color:C.lavender2},line:{color:C.lavender}});
  s.addText('Worker at time t',{x:1.1,y:2.45,w:2.8,h:0.38,fontSize:20,bold:true,color:C.dark,margin:0});
  bullet(s,'Age 18–64',1.12,3.05,2.9,13);
  bullet(s,'At work, wage/salary',1.12,3.48,2.9,13);
  bullet(s,'WKSTAT = 11 (full-time)',1.12,3.91,2.9,13);
  bullet(s,'Positive LNKFW1YWT',1.12,4.34,2.9,13);
  s.addShape(pptx.ShapeType.rightArrow,{x:5.2,y:2.85,w:2.7,h:1.15,fill:{color:C.violet},line:{color:C.violet}});
  s.addText('same CPSIDV\nsame month\n+ 1 year',{x:5.47,y:2.99,w:2.1,h:0.78,fontSize:15,bold:true,color:C.white,align:'center',margin:0});
  s.addShape(pptx.ShapeType.roundRect,{x:8.55,y:2.15,w:3.9,h:2.65,rectRadius:0.08,fill:{color:C.white},line:{color:C.violet,width:1.5}});
  s.addText('Worker at time t+1 year',{x:8.82,y:2.45,w:3.0,h:0.38,fontSize:20,bold:true,color:C.dark,margin:0});
  bullet(s,'Full-time',8.85,3.05,2.9,13);
  bullet(s,'PTER / involuntary PT',8.85,3.48,2.9,13);
  bullet(s,'Unemployed',8.85,3.91,2.9,13);
  bullet(s,'Other labor-force status',8.85,4.34,2.9,13);
  statCard(s,1.4,5.35,3.0,1.1,'1,112,384','eligible baselines');
  statCard(s,5.17,5.35,3.0,1.1,'1,062,438','validated matches');
  statCard(s,8.95,5.35,3.0,1.1,'13,103','FT→PTER transitions');
}

// Slide 5
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Transition cohorts are large in every year', 'The panel gives us enough observations to estimate year-specific patterns and later interact them with AI exposure.');
  const cats = ['2017→18','2018→19','2019→20','2020→21','2021→22','2022→23','2023→24','2024→25'];
  const n = [159924,158068,147156,116866,124820,123515,124692,107397];
  s.addChart(pptx.ChartType.bar,[{name:'Validated transitions',labels:cats,values:n}],{
    x:0.9,y:1.55,w:7.8,h:4.9,catAxisLabelFontSize:11,valAxisLabelFontSize:10,
    showLegend:false,showTitle:false,showValue:true,showCatName:false,
    dataLabelPosition:'outEnd',dataLabelColor:C.dark,dataLabelFormatCode:'#,##0',
    chartColors:[C.violet],gridLine:{color:'E5E7EB',width:1},showValue:false,
    valAxisMinVal:0,valAxisMaxVal:180000,valAxisMajorUnit:30000,
    border:{color:'FFFFFF',transparency:100}
  });
  statCard(s,9.15,1.75,2.85,1.45,'159,924','largest cohort','2017→2018');
  statCard(s,9.15,3.45,2.85,1.45,'107,397','smallest cohort','2024→2025');
  s.addShape(pptx.ShapeType.roundRect,{x:9.15,y:5.12,w:2.85,h:1.15,rectRadius:0.05,fill:{color:C.lavender2},line:{color:C.lavender}});
  s.addText('October 2024→2025 is absent because no October 2025 CPS was collected.',{x:9.35,y:5.34,w:2.45,h:0.65,fontSize:10.5,color:C.gray,margin:0});
}

// Slide 6
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'PTER transitions show a clear COVID shock', 'Weighted one-year probability that a baseline full-time worker is part-time for economic reasons one year later.');
  const cats = ['2017→18','2018→19','2019→20','2020→21','2021→22','2022→23','2023→24','2024→25'];
  const vals = [1.128,1.135,2.848,1.168,0.978,1.089,1.142,1.171];
  s.addChart(pptx.ChartType.line,[{name:'Weighted FT→PTER rate',labels:cats,values:vals}],{
    x:0.85,y:1.62,w:8.35,h:4.8,showLegend:false,showTitle:false,
    catAxisLabelFontSize:11,valAxisLabelFontSize:10,
    valAxisMinVal:0,valAxisMaxVal:3.2,valAxisMajorUnit:0.5,
    valAxisTitle:'Percent',showValue:false,
    chartColors:[C.violet],lineSize:3,showMarker:true,markerSize:7,
    gridLine:{color:'E5E7EB',width:1},border:{color:'FFFFFF',transparency:100}
  });
  s.addShape(pptx.ShapeType.roundRect,{x:9.55,y:1.85,w:2.8,h:1.55,rectRadius:0.06,fill:{color:C.redSoft},line:{color:C.redSoft}});
  s.addText('2.85%',{x:9.84,y:2.08,w:2.15,h:0.48,fontSize:29,bold:true,color:C.red,align:'center',margin:0});
  s.addText('2019→2020\nCOVID-era spike',{x:9.85,y:2.62,w:2.15,h:0.55,fontSize:12,bold:true,color:C.red,align:'center',margin:0});
  s.addShape(pptx.ShapeType.roundRect,{x:9.55,y:3.75,w:2.8,h:2.05,rectRadius:0.06,fill:{color:C.lavender2},line:{color:C.lavender}});
  s.addText('Interpretation',{x:9.78,y:4.03,w:2.2,h:0.32,fontSize:14,bold:true,color:C.violet,margin:0});
  s.addText('The pandemic years cannot be treated as ordinary pre/post observations. We will show them explicitly and run robustness checks excluding them.',{x:9.78,y:4.52,w:2.2,h:0.96,fontSize:11.5,color:C.dark,margin:0});
}

// Slide 7
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'What happens one year later?', 'Weighted follow-up status across all 1.06M validated transitions.');
  const outcomes = [
    ['Full-time','82.82%',C.violet],
    ['PT non-economic','7.26%',C.mid],
    ['PTER','1.34%','8B5CF6'],
    ['Unemployed','1.85%','C4B5FD'],
    ['Other / NILF / absent','6.73%','DDD6FE'],
  ];
  let x=0.85;
  outcomes.forEach((o,i)=>{
    const widths=[5.2,2.15,1.15,1.35,2.05];
    const w=widths[i];
    s.addShape(pptx.ShapeType.rect,{x,y:2.0,w,h:0.78,fill:{color:o[2]},line:{color:'FFFFFF',width:1}});
    x+=w;
  });
  x=0.85;
  outcomes.forEach((o,i)=>{
    const widths=[5.2,2.15,1.15,1.35,2.05]; const w=widths[i];
    s.addText(o[1],{x,y:2.18,w,h:0.34,fontSize:i===0?18:13.5,bold:true,color:i<3?C.white:C.dark,align:'center',margin:0});
    x+=w;
  });
  const cardData=[
    ['82.82%','remain full-time'],['7.26%','non-economic PT'],['1.34%','PTER'],['1.85%','unemployed']
  ];
  cardData.forEach((d,i)=>statCard(s,0.9+i*3.05,3.35,2.7,1.55,d[0],d[1]));
  s.addText('The PTER outcome is relatively rare. That is why we keep unemployment as a complementary outcome and report effect sizes with uncertainty—not only p-values.',{x:1.15,y:5.55,w:10.9,h:0.72,fontSize:16,bold:true,color:C.dark,align:'center',margin:0});
}

// Slide 8
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'What does “part-time for economic reasons” look like?', 'Reason codes among the 13,103 FT→PTER transitions.');
  const labels=['Slack work / business conditions','Could only find part-time','Seasonal work','Job started/ended during week'];
  const vals=[78.65,14.74,3.69,2.92];
  s.addChart(pptx.ChartType.bar,[{name:'Share of PTER',labels,values:vals}],{
    x:0.95,y:1.65,w:8.2,h:4.7,showLegend:false,showTitle:false,showValue:true,
    dataLabelPosition:'outEnd',dataLabelFormatCode:'0.0"%"',dataLabelColor:C.dark,
    catAxisLabelFontSize:11,valAxisLabelFontSize:10,valAxisMinVal:0,valAxisMaxVal:90,valAxisMajorUnit:15,
    chartColors:[C.violet],gridLine:{color:'E5E7EB',width:1},border:{color:'FFFFFF',transparency:100}
  });
  s.addShape(pptx.ShapeType.roundRect,{x:9.55,y:1.85,w:2.75,h:3.85,rectRadius:0.06,fill:{color:C.lavender2},line:{color:C.lavender}});
  s.addText('Why this matters',{x:9.82,y:2.12,w:2.2,h:0.38,fontSize:15,bold:true,color:C.violet,margin:0});
  s.addText('Most PTER transitions are tied to slack work or business conditions—not simply workers choosing fewer hours.',{x:9.82,y:2.68,w:2.15,h:1.05,fontSize:14,bold:true,color:C.dark,margin:0});
  s.addText('This makes the outcome a useful measure of labor-market deterioration, while still distinct from unemployment.',{x:9.82,y:4.10,w:2.15,h:1.05,fontSize:11.5,color:C.gray,margin:0});
}

// Slide 9
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Age is already an important confounder', 'Descriptive PTER transition rates differ sharply across age groups.');
  const labels=['18–24','25–34','35–44','45–54','55–64'];
  const vals=[2.646,1.387,1.200,1.136,1.078];
  s.addChart(pptx.ChartType.column,[{name:'Weighted FT→PTER rate',labels,values:vals}],{
    x:0.95,y:1.65,w:7.7,h:4.8,showLegend:false,showTitle:false,showValue:true,
    dataLabelPosition:'outEnd',dataLabelFormatCode:'0.00"%"',dataLabelColor:C.dark,
    catAxisLabelFontSize:12,valAxisLabelFontSize:10,valAxisMinVal:0,valAxisMaxVal:3.0,valAxisMajorUnit:0.5,
    chartColors:[C.violet],gridLine:{color:'E5E7EB',width:1},border:{color:'FFFFFF',transparency:100}
  });
  statCard(s,9.15,1.85,2.85,1.6,'2.65%','age 18–24','More than twice the rate for workers 45+.');
  s.addShape(pptx.ShapeType.roundRect,{x:9.15,y:3.78,w:2.85,h:2.1,rectRadius:0.06,fill:{color:C.lavender2},line:{color:C.lavender}});
  s.addText('Method implication',{x:9.42,y:4.06,w:2.2,h:0.35,fontSize:14,bold:true,color:C.violet,margin:0});
  s.addText('AI exposure is also correlated with occupation composition. We therefore need demographic controls and heterogeneity checks rather than comparing raw rates alone.',{x:9.42,y:4.55,w:2.2,h:1.05,fontSize:11.5,color:C.dark,margin:0});
}

// Slide 10
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Aggregate PTER did not rise after 2022', 'That does not answer our hypothesis—it clarifies what the exposure analysis must test.');
  statCard(s,1.05,1.78,3.7,2.0,'1.131%','Clean pre-period','Pooled weighted rate for 2017–18 baselines.');
  statCard(s,4.82,1.78,3.7,2.0,'2.848%','COVID shock','2019 baseline → 2020 follow-up.');
  statCard(s,8.59,1.78,3.7,2.0,'1.133%','Post-2022','Pooled weighted rate for 2022–24 baselines.');
  s.addText('≈ no aggregate change',{x:4.45,y:4.25,w:4.4,h:0.55,fontSize:25,bold:true,color:C.violet,align:'center',margin:0});
  s.addText('Our hypothesis is not “PTER increased nationally.” It is whether the post-2022 change was different for workers in more language-model-exposed occupations than for workers in less-exposed occupations.',{x:1.35,y:5.02,w:10.65,h:0.95,fontSize:17,bold:true,color:C.dark,align:'center',margin:0});
  s.addText('Therefore, the AIOE merge is the step that turns this descriptive panel into the actual research test.',{x:2.45,y:6.20,w:8.45,h:0.42,fontSize:13.5,color:C.gray,align:'center',margin:0});
}

// Slide 11
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Next: merge language-model exposure and test divergence', 'The CPS side is ready; the remaining task is occupation exposure and the regression design.');
  const steps=[
    ['1','CPS OCC2010','Baseline occupation'],
    ['2','Occupation crosswalk','Validate one-to-one / aggregated mappings'],
    ['3','6-digit SOC','Common occupation key'],
    ['4','LM-AIOE','Continuous language-model exposure'],
    ['5','Merged panel','Worker transition + exposure'],
  ];
  steps.forEach((st,i)=>{
    const x=0.75+i*2.48;
    s.addShape(pptx.ShapeType.ellipse,{x:x+0.78,y:1.62,w:0.62,h:0.62,fill:{color:C.violet},line:{color:C.violet}});
    s.addText(st[0],{x:x+0.78,y:1.76,w:0.62,h:0.22,fontSize:14,bold:true,color:C.white,align:'center',margin:0});
    s.addText(st[1],{x:x+0.15,y:2.45,w:1.9,h:0.42,fontSize:15,bold:true,color:C.dark,align:'center',margin:0});
    s.addText(st[2],{x:x+0.05,y:3.02,w:2.1,h:0.85,fontSize:10.5,color:C.gray,align:'center',margin:0});
    if(i<4) s.addShape(pptx.ShapeType.rightArrow,{x:x+2.05,y:1.78,w:0.55,h:0.25,fill:{color:C.mid},line:{color:C.mid}});
  });
  s.addShape(pptx.ShapeType.roundRect,{x:1.0,y:4.45,w:11.3,h:1.65,rectRadius:0.06,fill:{color:C.lavender2},line:{color:C.lavender}});
  s.addText('Primary test',{x:1.35,y:4.76,w:1.55,h:0.32,fontSize:14,bold:true,color:C.violet,margin:0});
  s.addText('Exposure × post-2022', {x:2.85,y:4.65,w:3.05,h:0.48,fontSize:22,bold:true,color:C.dark,margin:0});
  s.addText('Then: exposure quartiles, year-specific event-study interactions, COVID-exclusion robustness, and complementary unemployment outcome.',{x:6.05,y:4.62,w:5.7,h:0.72,fontSize:13,color:C.dark,margin:0});
}

// Slide 12
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Reproducible analysis pipeline', 'The GitHub package contains the code used to generate these CPS results.');
  const files=[
    ['01_build_cps_transitions.py','Streams the raw fixed-width IPUMS file, constructs the strict baseline sample, links CPSIDV one year forward, and writes transition-level data.'],
    ['02_summarize_cps_transitions.py','Creates year-by-year weighted rates, outcome distributions, PTER-reason tables, age heterogeneity, and reproducible figures.'],
    ['03_merge_aioe_template.py','Reserved for the occupation crosswalk + LM-AIOE merge once the exposure workbook is added.'],
    ['README.md + requirements.txt','Documents definitions, expected counts, commands, and dependency versions.'],
  ];
  files.forEach((f,i)=>{
    const y=1.55+i*1.18;
    s.addShape(pptx.ShapeType.roundRect,{x:0.9,y,w:3.55,h:0.82,rectRadius:0.05,fill:{color:C.lavender2},line:{color:C.lavender}});
    s.addText(f[0],{x:1.15,y:y+0.22,w:3.05,h:0.3,fontFace:'Aptos Mono',fontSize:12.5,bold:true,color:C.violet,margin:0});
    s.addText(f[1],{x:4.75,y:y+0.09,w:7.4,h:0.64,fontSize:12.5,color:C.dark,margin:0});
  });
  s.addShape(pptx.ShapeType.roundRect,{x:1.45,y:6.27,w:10.35,h:0.55,rectRadius:0.05,fill:{color:C.paleGray},line:{color:C.lightGray}});
  s.addText('Do not commit raw IPUMS microdata or row-level derived microdata to GitHub; commit the scripts and aggregate result tables.',{x:1.7,y:6.43,w:9.85,h:0.22,fontSize:11.5,bold:true,color:C.gray,align:'center',margin:0});
}

pptx.writeFile({ fileName: '/mnt/data/milestone2_cps_preliminary_results.pptx' });
