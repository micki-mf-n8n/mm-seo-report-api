'use strict';
const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, ImageRun, PageBreak, LevelFormat, Header, Footer
} = require('docx');
const { mmLogo, mfLogo } = require('./logos');

const app  = express();
app.use(express.json({ limit: '10mb' }));

// ── BRAND ──────────────────────────────────────────────────────────────────────
const BLUE_POS='1155CC',BLUE_MM='17375E',BLUE_COMP='1155CC';
const GREEN_CHG='188038',RED_CHG='B31412',GRAY_DASH='999999';
const ROW_MM_BG='EBF5FB',ROW_ALT='F8FAFB',BORDER_CLR='D8D8D8';
const TITLE_CLR='202124',SUBHEAD_CLR='444444';
const thin={style:BorderStyle.SINGLE,size:4,color:BORDER_CLR};
const none={style:BorderStyle.NONE,size:0,color:'FFFFFF'};
const cellB={top:thin,bottom:thin,left:thin,right:thin};
const bottomB={top:none,bottom:thin,left:none,right:none};

// ── HELPERS ────────────────────────────────────────────────────────────────────
function hdrCell(t){return new TableCell({borders:bottomB,shading:{fill:'FFFFFF',type:ShadingType.CLEAR},margins:{top:60,bottom:80,left:120,right:80},verticalAlign:VerticalAlign.BOTTOM,children:[new Paragraph({children:[new TextRun({text:t,bold:true,size:18,color:TITLE_CLR,font:'Calibri'})]})]})}
function textCell(t,{bg='FFFFFF',bold=false,color='000000',size=18,align=AlignmentType.LEFT,italic=false}={}){return new TableCell({borders:cellB,shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:60,bottom:60,left:120,right:80},children:[new Paragraph({alignment:align,children:[new TextRun({text:t,bold,italic,size,color,font:'Calibri'})]})]})}
function posCell(pos,isMM=false){const bg=isMM?ROW_MM_BG:'FFFFFF';const r=pos!=='-';const n=r?parseInt(pos):999;const color=!r?GRAY_DASH:n<=10?GREEN_CHG:n<=30?BLUE_POS:SUBHEAD_CLR;return new TableCell({borders:cellB,shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:60,bottom:60,left:80,right:80},verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:pos,bold:r&&n<=30,size:18,color,font:'Calibri'})]})]})}
function body(t,{size=22,bold=false,spacing=120}={}){return new Paragraph({children:[new TextRun({text:t,bold,size,font:'Calibri'})],spacing:{after:spacing}})}
function bullet(t){return new Paragraph({numbering:{reference:'bullets',level:0},children:[new TextRun({text:t,size:22,font:'Calibri'})],spacing:{after:100}})}
function spacer(n=160){return new Paragraph({children:[],spacing:{after:n}})}

// ── DOCUMENT BUILDER ───────────────────────────────────────────────────────────
async function buildReport(data) {
  const { month, year, siteAudit, keywords, visibilityTable, narrative, rankedCount } = data;
  const audit = siteAudit || {};
  const vis   = visibilityTable || [];

  // Sort rank table: MM last, others by top10 desc
  const RANK = [...vis]
    .map(v => {
      const mmKws = (keywords||[]).filter(k => k.montross && k.montross !== '-');
      const top10  = mmKws.filter(k => parseInt(k.montross) <= 10).length;
      const top100 = mmKws.filter(k => parseInt(k.montross) <= 100).length;
      if (v.domain?.includes('montrossmiller')) return { domain: v.domain, top10, top100, _isMM: true };
      // estimate from visibility order
      return { domain: v.domain, top10: 0, top100: 0, _isMM: false };
    });

  // Build keywords array for the table
  const KW_ARRAY = (keywords || []).map((k, i) => [
    `${i+1}. ${k.keyword}`,
    String(k.montross  ?? '-'),
    String(k.hensley   ?? '-'),
    String(k.wkw       ?? '-'),
    String(k.forthepeople ?? '-'),
    String(k.nleelaw   ?? '-'),
    String(k.ckflaw    ?? '-'),
  ]);

  // Top10/100 counts from keywords
  const mmPos    = (keywords||[]).map(k => k.montross).filter(v => v && v !== '-').map(Number);
  const mmTop10  = mmPos.filter(n => n <= 10).length;
  const mmTop100 = mmPos.filter(n => n <= 100).length;

  // Narrative bullets — use Claude's narrative or fallback
  const narr = narrative || {};
  const bullets = [
    narr.siteAuditNarrative || `The site holds a ${audit.health ?? 0}% health score and ${audit.aiSearchHealth ?? 0}% AI search health with ${audit.errors ?? 0} errors and ${audit.warnings ?? 0} warnings.`,
    narr.competitiveNarrative || `Montross Miller ranked for ${rankedCount || mmTop100} of 26 tracked keywords this month.`,
    narr.visibilityNarrative  || `Visibility data available in the table below.`,
  ];

  function kwRow(data, idx) {
    const bg=idx%2===1?ROW_ALT:'FFFFFF';
    const[kw,mm,h,wkw,fp,nl,ckf]=data;
    return new TableRow({children:[
      new TableCell({borders:cellB,shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:60,bottom:60,left:120,right:80},children:[new Paragraph({children:[new TextRun({text:kw,size:18,font:'Calibri',color:'333333'})]})]}),
      posCell(mm,true),posCell(h),posCell(wkw),posCell(fp),posCell(nl),posCell(ckf),
    ]});
  }

  function auditTable() {
    const errClr = (audit.errors||0)===0 ? GREEN_CHG : RED_CHG;
    const wrnClr = (audit.warnings||0)===0 ? GREEN_CHG : 'E67E22';
    const rows=[
      ['Site Health Score', `${audit.health??0}%`,   (audit.health??0)>=99?'\u2713 Excellent':'Minor issues present'],
      ['Errors',           String(audit.errors??0),  (audit.errors??0)===0?'\u2713 None':`${audit.errors} to fix`],
      ['Warnings',         String(audit.warnings??0),(audit.warnings??0)===0?'\u2713 None':'Low priority \u2014 to review'],
      ['Notices',          String(audit.notices??0), 'Informational only'],
      ['Pages Crawled',    String(audit.pagesCrawled??0),''],
      ['AI Search Health', `${audit.aiSearchHealth??0}%`,'\u2713 Excellent'],
    ];
    return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3500,1800,4060],rows:[
      new TableRow({children:[hdrCell('Metric'),hdrCell('Value'),hdrCell('Status')]}),
      ...rows.map((r,i)=>{
        const bg=i%2===1?ROW_ALT:'FFFFFF';
        let clr=TITLE_CLR;
        if(r[1]==='0'||r[1].startsWith('10')||parseInt(r[1])>=96) clr=GREEN_CHG;
        else if(i===2&&(audit.warnings||0)>0) clr='E67E22';
        else if(i===1&&(audit.errors||0)>0) clr=RED_CHG;
        return new TableRow({children:[
          textCell(r[0],{bold:true,bg,size:18}),
          textCell(r[1],{bold:true,color:clr,size:18,align:AlignmentType.CENTER}),
          textCell(r[2],{italic:true,color:SUBHEAD_CLR,size:18}),
        ]});
      })
    ]});
  }

  function visTable() {
    const sorted = [...vis].sort((a,b)=>{
      if(a.domain?.includes('montrossmiller')) return 1;
      if(b.domain?.includes('montrossmiller')) return -1;
      return (b.visibility||0)-(a.visibility||0);
    });
    return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[4060,2000,1800,1500],rows:[
      new TableRow({children:[hdrCell('Domain'),hdrCell(`Visibility (${month} end)`),hdrCell('Change'),hdrCell('Trend')]}),
      ...sorted.map((v,i)=>{
        const isMM=v.domain?.includes('montrossmiller');
        const bg=isMM?ROW_MM_BG:i%2===1?ROW_ALT:'FFFFFF';
        const chg=v.change??0;
        const up=chg>=0;
        const visStr=`${(v.visibility??0).toFixed(2)}%`;
        const chgStr=`${up?'+':''}${chg.toFixed(2)}%`;
        return new TableRow({children:[
          new TableCell({borders:cellB,shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:60,bottom:60,left:120,right:80},children:[new Paragraph({children:[new TextRun({text:v.domain||'',bold:isMM,size:18,color:isMM?BLUE_MM:BLUE_COMP,font:'Calibri'})]})]}),
          textCell(visStr,{bold:true,size:18,align:AlignmentType.CENTER,bg}),
          textCell(chgStr,{size:18,color:up?GREEN_CHG:RED_CHG,align:AlignmentType.CENTER,bg}),
          textCell(up?'\u2191':'\u2193',{size:20,color:up?GREEN_CHG:RED_CHG,align:AlignmentType.CENTER,bg}),
        ]});
      })
    ]});
  }

  function rankTable() {
    // Build rank data from keywords
    const domains = [
      {key:'hensley',domain:'hensleylegal.com'},
      {key:'wkw',domain:'wkw.com'},
      {key:'forthepeople',domain:'forthepeople.com'},
      {key:'nleelaw',domain:'nleelaw.com'},
      {key:'ckflaw',domain:'ckflaw.com'},
      {key:'montross',domain:'montrossmillerlaw.com'},
    ];
    const rankData = domains.map(({key,domain})=>{
      const vals=(keywords||[]).map(k=>k[key]||'-').filter(v=>v!=='-').map(Number);
      return {domain,top10:vals.filter(n=>n<=10).length,top100:vals.filter(n=>n<=100).length};
    }).sort((a,b)=>{
      if(a.domain.includes('montrossmiller')) return 1;
      if(b.domain.includes('montrossmiller')) return -1;
      return b.top10-a.top10||b.top100-a.top100;
    });
    return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[5760,1800,1800],rows:[
      new TableRow({children:[hdrCell('Domain'),hdrCell('Top 10 Keywords'),hdrCell('Top 100 Keywords')]}),
      ...rankData.map((r,i)=>{
        const isMM=r.domain.includes('montrossmiller');
        const bg=isMM?ROW_MM_BG:i%2===1?ROW_ALT:'FFFFFF';
        return new TableRow({children:[
          new TableCell({borders:cellB,shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:60,bottom:60,left:120,right:80},children:[new Paragraph({children:[new TextRun({text:r.domain,bold:isMM,size:18,color:isMM?BLUE_MM:BLUE_COMP,font:'Calibri'})]})]}),
          textCell(String(r.top10),{bold:true,size:18,align:AlignmentType.CENTER,bg}),
          textCell(String(r.top100),{bold:true,size:18,align:AlignmentType.CENTER,bg}),
        ]});
      })
    ]});
  }

  const mmVisRow = vis.find(v=>v.domain?.includes('montrossmiller'))||{};
  const mmVisStr = `${(mmVisRow.visibility??0).toFixed(2)}%`;
  const mmChgStr = `${(mmVisRow.change??0)>=0?'+':''}${(mmVisRow.change??0).toFixed(2)}%`;

  return new Document({
    numbering:{config:[{reference:'bullets',levels:[{level:0,format:LevelFormat.BULLET,text:'\u2022',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
    styles:{
      default:{document:{run:{font:'Calibri',size:22}}},
      paragraphStyles:[
        {id:'Title',name:'Title',basedOn:'Normal',quickFormat:true,run:{font:'Calibri',size:44,bold:true},paragraph:{spacing:{before:320,after:160}}},
      ]
    },
    sections:[{
      properties:{page:{size:{width:12240,height:15840},margin:{top:1080,right:1080,bottom:1080,left:1080,header:708,footer:708}}},
      headers:{default:new Header({children:[
        new Paragraph({spacing:{after:80},children:[
          new ImageRun({data:mfLogo,transformation:{width:170,height:27},type:'png'}),
          new TextRun({break:1}),new TextRun({break:1}),
          new TextRun({text:'9114 TECHNOLOGY LN, FISHERS, IN 46038',size:16,color:SUBHEAD_CLR,font:'Calibri'}),
        ]}),
      ]})},
      footers:{default:new Footer({children:[
        new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:`Generated ${month} ${year}`,size:16,color:SUBHEAD_CLR,font:'Calibri'})]}),
        new Paragraph({children:[]}),
        new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({data:mfLogo,transformation:{width:340,height:53},type:'png'})]}),
      ]})},
      children:[
        // COVER
        new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:2880,after:480},children:[new ImageRun({data:mmLogo,transformation:{width:260,height:86},type:'jpg'})]}),
        new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:480,after:160},children:[new TextRun({text:'SEO Audit',bold:true,size:72,font:'Calibri',color:TITLE_CLR})]}),
        new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:0},children:[new TextRun({text:`${month} ${year}`,size:44,font:'Calibri',color:TITLE_CLR})]}),
        new Paragraph({children:[new PageBreak()]}),

        // SITE AUDIT
        new Paragraph({heading:HeadingLevel.TITLE,children:[new TextRun({text:'Site Audit Update',font:'Calibri',bold:true,size:44})]}),
        spacer(40),
        new Paragraph({children:[new TextRun({font:'Calibri',size:32,text:'The site audit report is an assessment of the technical search engine optimization (SEO) and health of a website. A lower score means the website has issues that should be resolved to improve search engine optimization.'})],spacing:{before:160,after:120}}),
        spacer(80),
        body(`The summary below shows the number of errors and warnings on the Montross Miller website. The site audit health score is currently ${audit.health??0}%. There are ${audit.errors??0} errors and ${audit.warnings??0} warnings${(audit.warnings??0)===0?' — only '+(audit.notices??0)+' informational notices that do not impact crawling or rankings.':' that should be reviewed.'}`),
        spacer(80),
        auditTable(),
        new Paragraph({children:[new PageBreak()]}),

        // COMPETITIVE POSITIONING
        new Paragraph({heading:HeadingLevel.TITLE,children:[new TextRun({text:'Competitive Positioning',font:'Calibri',bold:true,size:44})]}),
        spacer(40),
        new Paragraph({children:[
          new TextRun({font:'Calibri',size:32,text:'The table below shows the current position of Montross Miller and the top competitors for the 26 keywords we are tracking. These rankings are a snapshot based on Google\u2019s data for the results in '}),
          new TextRun({font:'Calibri',size:32,italics:true,text:'Indiana'}),
          new TextRun({font:'Calibri',size:32,text:'.'}),
        ],spacing:{before:160,after:120}}),
        spacer(80),
        body(`In ${month} ${year}, Montross Miller ranked for ${rankedCount||mmTop100} of 26 tracked keywords. Notable highlights include:`),
        ...bullets.map(b=>bullet(b)),
        spacer(120),
        new Table({
          width:{size:9360,type:WidthType.DXA},
          columnWidths:[2700,900,900,900,1020,900,1040],
          rows:[
            new TableRow({children:[
              hdrCell('Keyword'),
              new TableCell({borders:bottomB,shading:{fill:ROW_MM_BG,type:ShadingType.CLEAR},margins:{top:60,bottom:80,left:80,right:80},children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'montross\nmillerlaw.com',bold:true,size:16,color:BLUE_MM,font:'Calibri'})]})]}),
              hdrCell('hensley\nlegal.com'),hdrCell('wkw.com'),
              hdrCell('forthe\npeople.com'),hdrCell('nlee\nlaw.com'),hdrCell('ckf\nlaw.com'),
            ]}),
            ...KW_ARRAY.map((kw,i)=>kwRow(kw,i))
          ]
        }),
        spacer(80),
        new Paragraph({spacing:{after:80},children:[
          new TextRun({text:'Position key:  ',bold:true,size:18,font:'Calibri',color:TITLE_CLR}),
          new TextRun({text:'\u25a0 ',size:18,font:'Calibri',color:GREEN_CHG}),new TextRun({text:'Top 10   ',size:18,font:'Calibri',color:SUBHEAD_CLR}),
          new TextRun({text:'\u25a0 ',size:18,font:'Calibri',color:BLUE_POS}),new TextRun({text:'Top 11\u201330   ',size:18,font:'Calibri',color:SUBHEAD_CLR}),
          new TextRun({text:'\u25a0 ',size:18,font:'Calibri',color:SUBHEAD_CLR}),new TextRun({text:'Top 31\u2013100   ',size:18,font:'Calibri',color:SUBHEAD_CLR}),
          new TextRun({text:'\u2014 Not in Top 100',size:18,font:'Calibri',color:GRAY_DASH}),
        ]}),
        new Paragraph({children:[new PageBreak()]}),

        // VISIBILITY
        new Paragraph({heading:HeadingLevel.TITLE,children:[new TextRun({text:'Competitive Positioning Continued',font:'Calibri',bold:true,size:44})]}),
        spacer(60),
        new Paragraph({children:[new TextRun({font:'Calibri',size:32,text:'Google\u2019s visibility metric in SEO refers to how often a website appears in search engine results compared to its competitors. It\u2019s a way to measure how visible or prominent your website is across all tracked keywords. As higher volume keywords improve in ranking, visibility will increase as well.'})],spacing:{before:160,after:120}}),
        spacer(80),
        body(`Below is a chart that outlines the visibility of Montross Miller versus the competitors. Montross Miller\u2019s visibility for the keywords within our tracking campaign was ${mmVisStr} in ${month}, a change of ${mmChgStr} compared to the start of the month.`),
        spacer(80),
        body('Indiana Visibility (26 keywords)',{bold:true,size:24}),
        spacer(80),
        visTable(),
        new Paragraph({children:[new PageBreak()]}),

        // RANKINGS
        new Paragraph({heading:HeadingLevel.TITLE,children:[new TextRun({text:'Competitive Positioning Continued',font:'Calibri',bold:true,size:44})]}),
        spacer(60),
        body('Below are two charts that compare keyword rankings for Montross Miller and competitors. The Rankings in the Top 10 and Top 100 tables show how many times each competitor and Montross Miller appear in the top 10 and top 100 results for the focus keywords.'),
        body(`Of the 26 keywords we\u2019re tracking, Montross Miller currently has ${mmTop10} keyword${mmTop10!==1?'s':''} within the Top 10 and ${mmTop100} keywords within the Top 100.`),
        spacer(120),
        rankTable(),
        new Paragraph({children:[new PageBreak()]}),

        // SUMMARY
        new Paragraph({heading:HeadingLevel.TITLE,children:[new TextRun({text:'Summary',font:'Calibri',bold:true,size:44})]}),
        spacer(80),
        bullet(`The site audit health score is currently ${audit.health??0}%, with ${audit.errors??0} errors, ${audit.warnings??0} warnings, and an AI Search Health score of ${audit.aiSearchHealth??0}%.`),
        bullet(bullets[1]),
        bullet(`Montross Miller\u2019s visibility was ${mmVisStr} at month end, a change of ${mmChgStr} compared to the start of ${month}.`),
        spacer(120),
        new Paragraph({heading:HeadingLevel.TITLE,children:[new TextRun({text:'Next Steps',font:'Calibri',bold:true,size:44})]}),
        spacer(60),
        body(`In ${month} ${year}, our SEO initiatives include:`),
        bullet('[Add next steps before sending to client]'),
      ]
    }]
  });
}

// ── ROUTES ─────────────────────────────────────────────────────────────────────
app.get('/health', (_, res) => res.json({ ok: true }));

app.post('/generate', async (req, res) => {
  try {
    const data = req.body;
    if (!data.month || !data.year) return res.status(400).json({ error: 'month and year required' });

    const doc    = await buildReport(data);
    const buffer = await Packer.toBuffer(doc);
    const base64 = buffer.toString('base64');
    const filename = data.filename || `MontrossMiller_SEO_${data.month}${data.year}.docx`;

    res.json({ success: true, filename, base64, sizeKB: (buffer.length/1024).toFixed(1) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SEO Report API running on port ${PORT}`));
