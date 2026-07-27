/* Foresight V24 deployment and interaction fixes */
(function(){
  'use strict';

  const NS='http://www.w3.org/2000/svg';
  const byId=id=>document.getElementById(id);
  const safeName=(name,fallback='Foresight_Model')=>String(name||fallback).replace(/[^a-z0-9 _-]+/gi,'').trim().replace(/\s+/g,'_')||fallback;
  const xmlEsc=v=>String(v??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[ch]));
  const colName=n=>{let s='';for(let x=n;x>0;x=Math.floor((x-1)/26))s=String.fromCharCode(65+(x-1)%26)+s;return s};
  const cell=(v,s=0)=>({v,s});
  const formula=(f,s=4,v=0)=>({f,s,v});

  const V24MiniXLSX={
    async build(sheets){
      if(typeof JSZip==='undefined')throw new Error('Excel packaging library is unavailable.');
      const zip=new JSZip();
      const now=new Date().toISOString();
      const contentTypes=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">','<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>','<Default Extension="xml" ContentType="application/xml"/>','<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>','<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>','<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>','<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'];
      sheets.forEach((_,i)=>contentTypes.push(`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`));
      contentTypes.push('</Types>');
      zip.file('[Content_Types].xml',contentTypes.join(''));
      zip.folder('_rels').file('.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
      zip.folder('docProps').file('core.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Foresight Generated Forecast Model</dc:title><dc:creator>Foresight</dc:creator><cp:lastModifiedBy>Foresight</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
      zip.folder('docProps').file('app.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Foresight</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheets.length}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="${sheets.length}" baseType="lpstr">${sheets.map(s=>`<vt:lpstr>${xmlEsc(s.name)}</vt:lpstr>`).join('')}</vt:vector></TitlesOfParts><Company>phamax</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion></Properties>`);
      const wbSheets=sheets.map((s,i)=>`<sheet name="${xmlEsc(String(s.name).slice(0,31))}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('');
      zip.folder('xl').file('workbook.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="18000"/></bookViews><sheets>${wbSheets}</sheets><calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>`);
      const rels=sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('');
      zip.folder('xl').folder('_rels').file('workbook.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
      zip.folder('xl').file('styles.xml',this.styles());
      const wsFolder=zip.folder('xl').folder('worksheets');
      sheets.forEach((s,i)=>wsFolder.file(`sheet${i+1}.xml`,this.sheetXml(s)));
      return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',compression:'DEFLATE',compressionOptions:{level:6}});
    },
    styles(){
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="4"><numFmt numFmtId="164" formatCode="#,##0;[Red](#,##0);-"/><numFmt numFmtId="165" formatCode="0.0%;[Red](0.0%);-"/><numFmt numFmtId="166" formatCode="$#,##0;[Red]($#,##0);-"/><numFmt numFmtId="167" formatCode="0.00"/></numFmts><fonts count="6"><font><sz val="10"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Aptos"/></font><font><color rgb="FF0000FF"/><sz val="10"/><name val="Aptos"/></font><font><color rgb="FF008000"/><sz val="10"/><name val="Aptos"/></font><font><color rgb="FF666666"/><sz val="10"/><name val="Aptos"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC8102E"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17191D"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEFF6FF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFECFDF3"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF4E5"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FFD0D5DD"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="11"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="166" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
    },
    sheetXml(sheet){
      const rows=sheet.rows||[],maxCol=Math.max(1,...rows.map(r=>r.length));
      const dim=`A1:${colName(maxCol)}${Math.max(1,rows.length)}`;
      const cols=(sheet.widths||[]).map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${Number(w)||12}" customWidth="1"/>`).join('');
      const pane=sheet.freezeRows?`<sheetViews><sheetView workbookViewId="0"><pane ySplit="${sheet.freezeRows}" topLeftCell="A${sheet.freezeRows+1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`:'<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
      const rowXml=rows.map((row,ri)=>{
        const cells=row.map((raw,ci)=>this.cellXml(raw,ri+1,ci+1)).join('');
        const ht=ri===0?22:18;
        return `<row r="${ri+1}" ht="${ht}" customHeight="1">${cells}</row>`;
      }).join('');
      const merges=(sheet.merges||[]).length?`<mergeCells count="${sheet.merges.length}">${sheet.merges.map(r=>`<mergeCell ref="${r}"/>`).join('')}</mergeCells>`:'';
      const filter=sheet.autoFilter?`<autoFilter ref="${sheet.autoFilter}"/>`:'';
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${pane}<dimension ref="${dim}"/>${cols?`<cols>${cols}</cols>`:''}<sheetData>${rowXml}</sheetData>${filter}${merges}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
    },
    cellXml(raw,row,col){
      if(raw===null||raw===undefined||raw==='')return '';
      const ref=`${colName(col)}${row}`;
      let v=raw,s=0,f=null;
      if(typeof raw==='object'&&!Array.isArray(raw)){v=raw.v;s=raw.s||0;f=raw.f||null}
      if(f!==null)return `<c r="${ref}" s="${s}"><f>${xmlEsc(f)}</f><v>${Number.isFinite(Number(v))?Number(v):0}</v></c>`;
      if(typeof v==='number'&&Number.isFinite(v))return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
      if(typeof v==='boolean')return `<c r="${ref}" s="${s}" t="b"><v>${v?1:0}</v></c>`;
      return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
    }
  };

  function scenarioAssumptions(){
    return [
      [cell('Assumption',1),cell('Unit',1),cell('Base',1),cell('Strong',1),cell('Weak',1),cell('Guidance',1)],
      ['Starting population','patients',cell(1000000,3),cell(1000000,3),cell(1000000,3),'Replace with country population or entry pool'],
      ['Annual population growth','%',cell(0.005,5),cell(0.007,5),cell(0.003,5),'Annual growth rate'],
      ['Prevalence rate','%',cell(0.05,5),cell(0.052,5),cell(0.048,5),'Use only for prevalence-pool models'],
      ['Diagnosis rate — start','%',cell(0.60,5),cell(0.62,5),cell(0.57,5),'Diagnosed / prevalent'],
      ['Diagnosis rate — end','%',cell(0.65,5),cell(0.70,5),cell(0.60,5),'End-of-horizon diagnosis rate'],
      ['Treatment rate','%',cell(0.80,5),cell(0.84,5),cell(0.75,5),'Treated / diagnosed'],
      ['Clinical eligibility','%',cell(0.60,5),cell(0.65,5),cell(0.55,5),'Eligible / treated'],
      ['Launch year','year',cell(2027,3),cell(2027,3),cell(2028,3),'First year with product share'],
      ['Peak product share','%',cell(0.20,5),cell(0.25,5),cell(0.15,5),'Peak share of eligible patients'],
      ['Years to peak','years',cell(5,3),cell(4,3),cell(7,3),'Linear demonstration ramp'],
      ['Annual doses per patient','doses',cell(365,3),cell(365,3),cell(365,3),'Replace with dosing logic'],
      ['Adherence','%',cell(0.75,5),cell(0.80,5),cell(0.68,5),'Applied to product volume'],
      ['List price per dose','currency',cell(100,10),cell(103,10),cell(97,10),'Replace with approved price'],
      ['Annual price growth','%',cell(0.02,5),cell(0.025,5),cell(0.005,5),'Annual price change'],
      ['Gross-to-net discount','%',cell(0.18,5),cell(0.16,5),cell(0.23,5),'Applied to gross revenue']
    ];
  }

  function forecastBlock(title,assumptionCol,startYear,endYear,startRow){
    const yrs=Array.from({length:endYear-startYear+1},(_,i)=>startYear+i);
    const r={title:startRow,header:startRow+1,pop:startRow+2,prevRate:startRow+3,prev:startRow+4,diagRate:startRow+5,diag:startRow+6,txRate:startRow+7,tx:startRow+8,eligRate:startRow+9,elig:startRow+10,share:startRow+11,prod:startRow+12,dose:startRow+13,adh:startRow+14,vol:startRow+15,price:startRow+16,gross:startRow+17,gtn:startRow+18,net:startRow+19};
    const rows=[];
    rows.push([cell(title,2)]);
    rows.push([cell('Metric',1),cell('Unit',1),...yrs.map(y=>cell(y,1))]);
    const add=(label,unit,vals)=>rows.push([label,unit,...vals]);
    const a=n=>`Assumptions!$${assumptionCol}$${n}`;
    add('Population','patients',yrs.map((_,i)=>i===0?formula(a(2),4):formula(`${colName(i+1)}${r.pop}*(1+${a(3)})`,4)));
    add('Prevalence rate','%',yrs.map(()=>formula(a(4),6)));
    add('Prevalent patients','patients',yrs.map((_,i)=>formula(`${colName(i+3)}${r.pop}*${colName(i+3)}${r.prevRate}`,4)));
    add('Diagnosis rate','%',yrs.map((_,i)=>formula(`${a(5)}+(${a(6)}-${a(5)})*${i}/${Math.max(1,yrs.length-1)}`,6)));
    add('Diagnosed patients','patients',yrs.map((_,i)=>formula(`${colName(i+3)}${r.prev}*${colName(i+3)}${r.diagRate}`,4)));
    add('Treatment rate','%',yrs.map(()=>formula(a(7),6)));
    add('Treated patients','patients',yrs.map((_,i)=>formula(`${colName(i+3)}${r.diag}*${colName(i+3)}${r.txRate}`,4)));
    add('Clinical eligibility','%',yrs.map(()=>formula(a(8),6)));
    add('Eligible patients','patients',yrs.map((_,i)=>formula(`${colName(i+3)}${r.tx}*${colName(i+3)}${r.eligRate}`,4)));
    add('Product share','%',yrs.map((yr,i)=>formula(`IF(${yr}<${a(9)},0,MIN(${a(10)},(${yr}-${a(9)}+1)/${a(11)}*${a(10)}))`,6)));
    add('Product patients','patients',yrs.map((_,i)=>formula(`${colName(i+3)}${r.elig}*${colName(i+3)}${r.share}`,4)));
    add('Annual doses per patient','doses',yrs.map(()=>formula(a(12),4)));
    add('Adherence','%',yrs.map(()=>formula(a(13),6)));
    add('Product volume','doses',yrs.map((_,i)=>formula(`${colName(i+3)}${r.prod}*${colName(i+3)}${r.dose}*${colName(i+3)}${r.adh}`,4)));
    add('List price per dose','currency',yrs.map((_,i)=>formula(`${a(14)}*(1+${a(15)})^${i}`,7)));
    add('Gross revenue','currency',yrs.map((_,i)=>formula(`${colName(i+3)}${r.vol}*${colName(i+3)}${r.price}`,7)));
    add('Gross-to-net discount','%',yrs.map(()=>formula(a(16),6)));
    add('Net revenue','currency',yrs.map((_,i)=>formula(`${colName(i+3)}${r.gross}*(1-${colName(i+3)}${r.gtn})`,7)));
    rows.push([]);
    return rows;
  }

  function epiModelSpec(){
    const cfg=typeof v16CanonicalConfig==='function'?v16CanonicalConfig():null;
    const e=(typeof v13BuilderState!=='undefined'&&v13BuilderState.epi)||{forecastName:'Foresight EPI Model',disease:'Disease',country:'United States',startYear:2026,endYear:2040,horizon:15,entry:'Prevalence Pool',basis:'Prevalence'};
    const modelName=e.forecastName||cfg?.modelInformation?.modelName||'Foresight EPI Model';
    const architecture=[
      [cell('Stage',1),cell('Configured structure',1)],
      ['Model family','Epidemiology-Based Forecast'],['Disease / indication',e.disease],['Country',e.country],['Forecast period',`${e.startYear}-${e.endYear}`],['Patient pool entry',e.entry],['Epidemiology basis',e.basis||'Not applicable'],['Age groups',e.populationSplit?`${e.populationGroupCount} groups: ${(e.populationGroupLabels||[]).slice(0,e.populationGroupCount).join(', ')}`:(e.poolGroupCount?`${e.poolGroupCount} pool age groups`:'No age split')],['Diagnosis',e.diagnosis?`${e.diagnosisSegments||0} segment(s), ${e.diagnosisSubsegments||0} subsegment(s)`:'Not included'],['Treatment',e.treatment?`${e.treatmentSegments||0} segment(s), ${e.treatmentSubsegments||0} subsegment(s)`:'Not included'],['Line of therapy',e.lot?`${e.lotCount||1} line(s)`:'Not included'],['Product classes',e.classes?`${e.classCount||1} class(es)`:'Not included'],['Products',e.classes?`${e.classMarketedProducts||0} marketed + ${e.classPipelineProducts||0} pipeline per class`:e.product?`${e.directMarketedProducts||0} marketed + ${e.directPipelineProducts||0} pipeline`:'Not included']
    ];
    const inputRows=[[cell('Category',1),cell('Model stage',1),cell('Input / assumption',1),cell('Unit',1),cell('Source / note',1)]];
    if(cfg&&Array.isArray(cfg.inputRows))cfg.inputRows.slice(1).forEach(r=>inputRows.push(r.map((v,i)=>i===4?cell(v||'',9):v)));
    if(inputRows.length===1)[['Epidemiology','Patient Potential Tower','Population','patients','User input'],['Epidemiology','Patient Potential Tower','Prevalence / incidence','rate','User input'],['Commercial','Product conversion','Price and GTN','currency / %','User input']].forEach(r=>inputRows.push(r));
    const info=[[cell('Foresight Generated EPI Model',2)],[],[cell('Field',1),cell('Value',1)],['Model name',modelName],['Disease / indication',e.disease],['Country',e.country],['Start year',e.startYear],['End year',e.endYear],['Horizon',e.endYear-e.startYear+1],['Generated',new Date().toISOString()],['Generation method','Browser-based static Excel generator'],['Governance','Draft model; review and approve assumptions before use']];
    const assumptions=scenarioAssumptions();
    const outputs=[...forecastBlock('Base Scenario','C',e.startYear,e.endYear,1),...forecastBlock('Strong Scenario','D',e.startYear,e.endYear,23),...forecastBlock('Weak Scenario','E',e.startYear,e.endYear,45)];
    const validation=[[cell('Status',1),cell('Check',1),cell('Detail',1)],[cell('Passed',8),'Model information',`${modelName} · ${e.country} · ${e.startYear}-${e.endYear}`],[cell('Passed',8),'Patient-pool configuration',`${e.entry}${e.basis?' · '+e.basis:''}`],[cell('Review',9),'Assumption evidence','Replace illustrative assumption defaults with source-linked approved inputs'],[cell('Review',9),'Scenario assumptions','Strong and Weak defaults are illustrative starting points'],[cell('Passed',8),'Formula scaffold','Patient-to-volume-to-revenue formulas included'],[cell('Passed',8),'Base protection','Scenario sheets are separate and do not overwrite approved inputs']];
    return {filename:`${safeName(modelName,'Foresight_EPI_Model')}.xlsx`,status:'Browser Excel engine ready · formula-wired .xlsx model',sheets:[
      {name:'Model Information',rows:info,widths:[32,90],merges:['A1:B1']},
      {name:'Model Architecture',rows:architecture,widths:[28,90],freezeRows:1,autoFilter:`A1:B${architecture.length}`},
      {name:'Input Register',rows:inputRows,widths:[20,26,42,18,52],freezeRows:1,autoFilter:`A1:E${inputRows.length}`},
      {name:'Assumptions',rows:assumptions,widths:[34,16,15,15,15,55],freezeRows:1,autoFilter:`A1:F${assumptions.length}`},
      {name:'Forecast Outputs',rows:outputs,widths:[34,14,...Array.from({length:e.endYear-e.startYear+1},()=>14)],freezeRows:2,merges:[`A1:${colName(e.endYear-e.startYear+3)}1`,`A23:${colName(e.endYear-e.startYear+3)}23`,`A45:${colName(e.endYear-e.startYear+3)}45`]},
      {name:'Patient Potential Tower',rows:[[cell('Year',1),cell('Prevalent',1),cell('Diagnosed',1),cell('Treated',1),cell('Eligible',1),cell('Product Patients',1)],...Array.from({length:e.endYear-e.startYear+1},(_,i)=>{const yr=e.startYear+i,fc=colName(i+3);return [yr,formula(`'Forecast Outputs'!${fc}5`,4),formula(`'Forecast Outputs'!${fc}7`,4),formula(`'Forecast Outputs'!${fc}9`,4),formula(`'Forecast Outputs'!${fc}11`,4),formula(`'Forecast Outputs'!${fc}13`,4)]})],widths:[12,18,18,18,18,20],freezeRows:1},
      {name:'Validation',rows:validation,widths:[14,34,92],freezeRows:1}
    ]};
  }

  function marketModelSpec(){
    const m=(typeof v13BuilderState!=='undefined'&&v13BuilderState.market)||{name:'Hemaris Switzerland Forecast',product:'Hemaris',country:'Switzerland',currency:'CHF',start:2022,horizon:10,external:'Client-owned syndicated market data',internal:'Internal sales data',granularity:'SKU / strength / pack',method:'Historical trend + events',products:3,competitors:4,skus:5,output:'Baseline vs Event-adjusted',service:2.5,supply:5,cogs:'Per unit'};
    const yrs=Array.from({length:m.horizon},(_,i)=>m.start+i);
    const setup=[[cell('Foresight Generated Market Model',2)],[],[cell('Field',1),cell('Value',1)],['Forecast name',m.name],['Product',m.product],['Country',m.country],['Currency',m.currency],['Forecast period',`${m.start}-${m.start+m.horizon-1}`],['External market source',m.external],['Internal sales source',m.internal],['Granularity',m.granularity],['Forecast method',m.method],['Products',m.products],['Competitors',m.competitors],['SKUs',m.skus],['Output',m.output],['Generated',new Date().toISOString()]];
    const inputs=[[cell('Metric',1),cell('Unit',1),...yrs.map(y=>cell(y,1))],['Baseline market volume','units',...yrs.map((_,i)=>cell(i===0?1000000:'',3))],['Annual market growth','%',...yrs.map(()=>cell(-0.02,5))],['Product share','%',...yrs.map(()=>cell(0.20,5))],['Price per unit',m.currency,...yrs.map(()=>cell(1.00,10))],['Volume event impact','%',...yrs.map(()=>cell(0,5))],['Price event impact','%',...yrs.map(()=>cell(0,5))],['Event probability','%',...yrs.map(()=>cell(1,5))],['COGS per unit',m.currency,...yrs.map(()=>cell(0.35,10))],['Service fee','%',...yrs.map(()=>cell((m.service||0)/100,5))],['Supply fee','%',...yrs.map(()=>cell((m.supply||0)/100,5))]];
    const outputRows=[[cell('Metric',1),cell('Unit',1),...yrs.map(y=>cell(y,1))]];
    const rows={market:2,share:4,price:5,vEvent:6,pEvent:7,prob:8,cogs:9,service:10,supply:11};
    const out={baseVol:2,adjVol:3,baseRev:4,adjRev:5,grossMargin:6,gmPct:7};
    const f=(col,row)=>`'Market Inputs'!${col}${row}`;
    outputRows.push(['Baseline product volume','units',...yrs.map((_,i)=>{const c=colName(i+3);return formula(i===0?`${f(c,rows.market)}*${f(c,rows.share)}`:`${colName(i+2)}${out.baseVol}*(1+${f(c,3)})`,4)})]);
    outputRows.push(['Event-adjusted volume','units',...yrs.map((_,i)=>{const c=colName(i+3);return formula(`${c}${out.baseVol}*(1+${f(c,rows.vEvent)}*${f(c,rows.prob)})`,4)})]);
    outputRows.push(['Baseline revenue',m.currency,...yrs.map((_,i)=>{const c=colName(i+3);return formula(`${c}${out.baseVol}*${f(c,rows.price)}`,7)})]);
    outputRows.push(['Event-adjusted revenue',m.currency,...yrs.map((_,i)=>{const c=colName(i+3);return formula(`${c}${out.adjVol}*${f(c,rows.price)}*(1+${f(c,rows.pEvent)}*${f(c,rows.prob)})`,7)})]);
    outputRows.push(['Gross margin',m.currency,...yrs.map((_,i)=>{const c=colName(i+3);return formula(`${c}${out.adjRev}-${c}${out.adjVol}*${f(c,rows.cogs)}-${c}${out.adjRev}*(${f(c,rows.service)}+${f(c,rows.supply)})`,7)})]);
    outputRows.push(['Gross margin %','%',...yrs.map((_,i)=>{const c=colName(i+3);return formula(`IFERROR(${c}${out.grossMargin}/${c}${out.adjRev},0)`,6)})]);
    const sku=[[cell('SKU',1),cell('Strength',1),cell('Pack',1),cell('Volume share',1),cell('Price',1),cell('COGS / unit',1)],...Array.from({length:m.skus||5},(_,i)=>[`SKU ${i+1}`,'', '',cell(i===0?1/(m.skus||5):'',5),cell('',10),cell('',10)])];
    return {filename:`${safeName(m.name,'Foresight_Market_Model')}.xlsx`,status:'Browser Excel engine ready · formula-wired .xlsx model',sheets:[
      {name:'Model Setup',rows:setup,widths:[32,80],merges:['A1:B1']},
      {name:'Market Inputs',rows:inputs,widths:[30,14,...yrs.map(()=>14)],freezeRows:1},
      {name:'Forecast Outputs',rows:outputRows,widths:[30,14,...yrs.map(()=>14)],freezeRows:1},
      {name:'SKU Structure',rows:sku,widths:[22,18,18,18,18,18],freezeRows:1,autoFilter:`A1:F${sku.length}`},
      {name:'Validation',rows:[[cell('Status',1),cell('Check',1),cell('Detail',1)],[cell('Passed',8),'Forecast period',`${m.start}-${m.start+m.horizon-1}`],[cell('Review',9),'Historical data','Populate client-owned market and internal sales history'],[cell('Review',9),'Events','Set event impact, probability, and timing'],[cell('Passed',8),'Financial conversion','Revenue, COGS, fee, and gross-margin formulas included']],widths:[14,34,90],freezeRows:1}
    ]};
  }

  function demoSeries(base){
    const strong=[],weak=[];
    years.forEach((yr,i)=>{const ramp=Math.max(0,Math.min(1,(yr-2025)/7)),late=Math.max(0,Math.min(1,(yr-2032)/8));strong.push((base[i]||0)*(1+0.16*ramp-0.04*late));weak.push((base[i]||0)*(1-0.10*ramp-0.08*late))});
    return {base:[...base],strong,weak};
  }

  function scenarioWorkbookSpec(){
    const current=(typeof v16ScenarioState!=='undefined'&&v16ScenarioState.current)||null;
    if(current){
      const outputs=[[cell('Year',1),cell('Base patients',1),cell('Scenario patients',1),cell('Patient delta',1),cell('Base revenue',1),cell('Scenario revenue',1),cell('Revenue delta',1)],...years.map((y,i)=>[y,current.patientBase[i]||0,current.patientSim[i]||0,(current.patientSim[i]||0)-(current.patientBase[i]||0),current.revenueBase[i]||0,current.revenueSim[i]||0,(current.revenueSim[i]||0)-(current.revenueBase[i]||0)])];
      return {filename:`Atopic_Dermatitis_${safeName(current.name,'Draft_Scenario')}.xlsx`,status:'Calculated scenario workbook · Base protected',sheets:[{name:'Scenario Summary',rows:[[cell('Field',1),cell('Value',1)],['Scenario',current.name],['Scenario ID',current.id],['Generated',new Date().toISOString()],['Governance','Separate draft; approved Base remains unchanged']],widths:[28,90]},{name:'Calculated Outputs',rows:outputs,widths:[12,20,20,20,20,20,20],freezeRows:1},{name:'Assumptions',rows:[[cell('Lever',1),cell('Value',1)],...Object.entries(current.assumptions||{}).map(([k,v])=>[k,v])],widths:[34,30],freezeRows:1}]};
    }
    const d=typeof scenarioData==='function'?scenarioData():{arr:GLOBAL_DATA.revenue.total,kind:'globalRevenue',title:'Global Net Revenue'},s=demoSeries(d.arr);
    return {filename:'Atopic_Dermatitis_Base_Strong_Weak_Demo.xlsx',status:'Scenario comparison workbook · illustrative Strong and Weak',sheets:[{name:'Scenario Outputs',rows:[[cell('Year',1),cell('Base',1),cell('Strong - illustrative',1),cell('Weak - illustrative',1),cell('Strong vs Base',1),cell('Weak vs Base',1)],...years.map((y,i)=>[y,s.base[i],s.strong[i],s.weak[i],s.strong[i]-s.base[i],s.weak[i]-s.base[i]])],widths:[12,22,22,22,22,22],freezeRows:1},{name:'Governance',rows:[[cell('Field',1),cell('Value',1)],['Base','Workbook-derived'],['Strong and Weak','Illustrative demo paths because cached source outputs overlap'],['Generated',new Date().toISOString()],['Base protection','No source workbook is overwritten']],widths:[30,95]}]};
  }

  async function saveBlob(blob,name){
    if(window.showSaveFilePicker){
      try{const handle=await window.showSaveFilePicker({suggestedName:name,types:[{description:'Microsoft Excel Workbook',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}]});const writable=await handle.createWritable();await writable.write(blob);await writable.close();return true}catch(err){if(err&&err.name==='AbortError')return false}
    }
    const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);return true;
  }

  window.v18CheckEngine=async function(){
    const dot=byId('v18-engine-dot'),text=byId('v18-engine-text');
    if(dot)dot.style.background='#12b76a';
    if(text)text.textContent='Browser Excel generator ready · no local service required';
    return true;
  };

  window.v15OpenModelSaveModal=function(type){
    try{
      const spec=type==='market'?marketModelSpec():epiModelSpec();
      window.v15PendingWorkbookPayload={kind:'v24-xlsx',spec,type,filename:spec.filename,status:spec.status};
      byId('modelSaveTitle').textContent='Generate Excel Model';
      byId('modelSaveDescription').textContent='The browser will generate a formula-wired Excel workbook. No localhost service is required.';
      byId('modelSaveStatus').textContent=spec.status;
      byId('modelSaveFilename').value=spec.filename;
      const typeCell=document.querySelector('#modelSaveModal .save-summary-cell b');if(typeCell)typeCell.textContent='Microsoft Excel Workbook (.xlsx)';
      byId('modelSaveModal').classList.add('open');
    }catch(err){if(typeof toast==='function')toast(err.message||'Unable to prepare the Excel model.');}
  };

  window.downloadScenarioWorkbook=function(){
    try{const spec=scenarioWorkbookSpec();window.v15PendingWorkbookPayload={kind:'v24-xlsx',spec,type:'scenario',filename:spec.filename,status:spec.status};byId('modelSaveTitle').textContent='Save Scenario Workbook';byId('modelSaveDescription').textContent='The export contains scenario outputs, assumptions, and governance metadata.';byId('modelSaveStatus').textContent=spec.status;byId('modelSaveFilename').value=spec.filename;byId('modelSaveModal').classList.add('open')}catch(err){if(typeof toast==='function')toast(err.message||'Unable to prepare the scenario workbook.')}
  };

  window.v15ConfirmWorkbookSave=async function(){
    const p=window.v15PendingWorkbookPayload;if(!p||p.kind!=='v24-xlsx'){if(typeof toast==='function')toast('No generated workbook is ready.');return}
    let name=(byId('modelSaveFilename').value||p.filename).trim();if(!name.toLowerCase().endsWith('.xlsx'))name+='.xlsx';
    const btn=document.querySelector('#modelSaveModal .btn.red'),old=btn?btn.textContent:'';
    try{if(btn){btn.disabled=true;btn.textContent='Generating Excel…'}byId('modelSaveStatus').textContent='Building workbook sheets and formulas in the browser…';const blob=await V24MiniXLSX.build(p.spec.sheets);const saved=await saveBlob(blob,name);if(!saved)return;if(typeof v15CloseSaveModal==='function')v15CloseSaveModal();if(p.type==='epi')byId('v13EpiSuccess')?.classList.add('open');if(p.type==='market')byId('v13MarketSuccess')?.classList.add('open');if(typeof toast==='function')toast(p.type==='scenario'?'Scenario workbook saved.':'Excel model generated and saved.')}catch(err){byId('modelSaveStatus').textContent=err.message||'Excel generation failed.';if(typeof toast==='function')toast(err.message||'The Excel model could not be generated.')}finally{if(btn){btn.disabled=false;btn.textContent=old||'Choose location & save'}}
  };

  function ensureTip(svg,id){let tip=byId(id);if(!tip){tip=document.createElement('div');tip.id=id;tip.className='tooltip';svg.parentElement.appendChild(tip)}return tip}
  function attachLineHover(svg,tip,series,W,H,pad,max,formatter,labels=years){
    svg.querySelectorAll('.v24-hover-layer,.v24-hover-guide').forEach(n=>n.remove());
    const x=i=>pad.l+(W-pad.l-pad.r)*(i/Math.max(1,labels.length-1)),y=v=>pad.t+(max-v)/max*(H-pad.t-pad.b);
    const g=document.createElementNS(NS,'g');g.setAttribute('class','v24-hover-guide');g.style.display='none';
    const line=document.createElementNS(NS,'line');line.setAttribute('y1',pad.t);line.setAttribute('y2',H-pad.b);line.setAttribute('stroke','#667085');line.setAttribute('stroke-dasharray','3 3');line.setAttribute('opacity','.65');g.appendChild(line);
    const circles=series.map(s=>{const c=document.createElementNS(NS,'circle');c.setAttribute('r','4');c.setAttribute('fill','#fff');c.setAttribute('stroke',s.color);c.setAttribute('stroke-width','2.5');g.appendChild(c);return c});svg.appendChild(g);
    const step=(W-pad.l-pad.r)/Math.max(1,labels.length-1);
    labels.forEach((lab,i)=>{const rect=document.createElementNS(NS,'rect');rect.setAttribute('class','v24-hover-layer');rect.setAttribute('x',pad.l+i*step-step/2);rect.setAttribute('y',pad.t);rect.setAttribute('width',Math.max(step,10));rect.setAttribute('height',H-pad.t-pad.b);rect.setAttribute('fill','transparent');rect.addEventListener('mousemove',ev=>{const xv=x(i);line.setAttribute('x1',xv);line.setAttribute('x2',xv);series.forEach((s,j)=>{circles[j].setAttribute('cx',xv);circles[j].setAttribute('cy',y(s.values[i]||0))});g.style.display='block';tip.innerHTML=`<b>${lab}</b><br>${series.map(s=>`${s.name}: ${formatter(s.values[i]||0)}`).join('<br>')}`;tip.style.display='block';const wrap=svg.parentElement.getBoundingClientRect();let left=ev.clientX-wrap.left+12,top=ev.clientY-wrap.top-12;left=Math.min(left,Math.max(8,wrap.width-190));top=Math.max(6,top);tip.style.left=left+'px';tip.style.top=top+'px'});rect.addEventListener('mouseleave',()=>{g.style.display='none';tip.style.display='none'});svg.appendChild(rect)});
  }

  window.drawEpiLines=function(svg,W,H,pad,focusYear,withHover=true){
    if(!svg)return;const e=DATA.epi,series=[{name:'Prevalence',color:'#a7abb1',values:e.Prevalence},{name:'Diagnosed',color:'#555b63',values:e.Diagnosed},{name:'Treated',color:'#c8102e',values:e.Treated},{name:'Eligible',color:'#ff9099',values:e.Eligible}],max=Math.max(...e.Prevalence)*1.08,x=i=>xScale(i,W,pad);let h=gridSvg(W,H,pad,max,fmtPatient,4);years.forEach((yr,i)=>{if(i%2===0||i===years.length-1)h+=`<text x="${x(i)}" y="${H-12}" text-anchor="middle" font-size="8" fill="#98a2b3">${yr}</text>`});series.forEach(s=>{h+=`<path d="${linePath(s.values,W,H,pad,max)}" fill="none" stroke="${s.color}" stroke-width="2.4"/>`});const fi=yi(focusYear);h+=`<line x1="${x(fi)}" y1="${pad.t}" x2="${x(fi)}" y2="${H-pad.b}" stroke="#c8102e" stroke-dasharray="4 4" opacity=".35"/>`;svg.innerHTML=h;const tip=ensureTip(svg,svg.id==='overviewEpiChart'?'overviewEpiTip':'epiTip');attachLineHover(svg,tip,series,W,H,pad,max,fmtPatient)
  };

  window.renderOverviewEpi=function(){drawEpiLines(byId('overviewEpiChart'),1100,255,{l:66,r:18,t:18,b:38},state.overviewYear,true)};

  window.renderOverviewChart=function(){
    const svg=byId('overviewChart');if(!svg)return;const W=980,H=390,pad={l:68,r:22,t:24,b:52},d=overviewSeries(),format=state.overviewMeasure==='revenue'?fmtGlobalRevenue:fmtGlobalVolume,demo=demoSeries(d.total),all=state.overviewView==='scenario'?[...demo.base,...demo.strong,...demo.weak]:[...d.total,...d.us,...d.ex],max=Math.max(...all)*1.12||1,x=i=>xScale(i,W,pad),y=v=>pad.t+(max-v)/max*(H-pad.t-pad.b);let h=gridSvg(W,H,pad,max,format);years.forEach((yr,i)=>{if(i%2===0||i===years.length-1)h+=`<text x="${x(i)}" y="${H-18}" text-anchor="middle" font-size="9" fill="#98a2b3">${yr}</text>`});let hoverSeries=[];
    if(state.overviewView==='scenario'){
      h+=`<path d="${linePath(demo.weak,W,H,pad,max)}" fill="none" stroke="#98a2b3" stroke-width="2.5" stroke-dasharray="6 5"/><path d="${linePath(demo.strong,W,H,pad,max)}" fill="none" stroke="#17191d" stroke-width="2.8"/><path d="${linePath(demo.base,W,H,pad,max)}" fill="none" stroke="#c8102e" stroke-width="3.3" stroke-linecap="round"/>`;
      hoverSeries=[{name:'Base',color:'#c8102e',values:demo.base},{name:'Strong · illustrative',color:'#17191d',values:demo.strong},{name:'Weak · illustrative',color:'#98a2b3',values:demo.weak}];
      if(state.overviewLabels)years.forEach((_,i)=>{[hoverSeries[0],hoverSeries[1],hoverSeries[2]].forEach((s,j)=>{const v=s.values[i]||0;if(v>0&&i%2===0)h+=`<text x="${x(i)}" y="${y(v)-8-j*2}" text-anchor="middle" font-size="7.5" fill="${s.color}">${format(v)}</text>`})});
      byId('overviewLegend').innerHTML='<span><i class="sw" style="height:3px;background:#c8102e"></i>Base · workbook</span><span><i class="sw" style="height:3px;background:#17191d"></i>Strong · illustrative</span><span><i class="sw" style="height:3px;background:#98a2b3"></i>Weak · illustrative</span><span class="epi-demo-chip">Hover for values</span>';
      byId('overviewChartTitle').textContent=`Global ${state.overviewMeasure==='revenue'?'Net Revenue':'Volume'} — Scenario Outlook`;byId('overviewChartSub').textContent='Base is workbook-derived; Strong and Weak are illustrative demo paths';const alert=byId('overviewAlert');if(alert){alert.style.display='flex';alert.className='epi-info';alert.innerHTML='<div>i</div><div><b>Illustrative scenario spread is enabled for the demo.</b> Hover any year to see Base, Strong, and Weak values.</div>'}const sf=byId('overviewScenarioField');if(sf)sf.style.display='none';
    }else{
      const bw=Math.max(10,(W-pad.l-pad.r)/years.length*.58);years.forEach((_,i)=>{const us=d.us[i]||0,ex=d.ex[i]||0,total=us+ex,base=y(0),usy=y(us),ty=y(total);h+=`<rect x="${x(i)-bw/2}" y="${usy}" width="${bw}" height="${Math.max(0,base-usy)}" fill="#cfd3d8"/><rect x="${x(i)-bw/2}" y="${ty}" width="${bw}" height="${Math.max(0,usy-ty)}" fill="#8d96a2"/>`});h+=`<path d="${linePath(d.total,W,H,pad,max)}" fill="none" stroke="#c8102e" stroke-width="2.6"/>`;hoverSeries=[{name:'US',color:'#cfd3d8',values:d.us},{name:'Ex-US',color:'#8d96a2',values:d.ex},{name:'Global total',color:'#c8102e',values:d.total}];byId('overviewLegend').innerHTML='<span><i class="sw" style="background:#cfd3d8"></i>US</span><span><i class="sw" style="background:#8d96a2"></i>Ex-US</span><span><i class="sw" style="height:3px;background:#c8102e"></i>Global Total</span>';byId('overviewChartTitle').textContent=`Global ${state.overviewMeasure==='revenue'?'Net Revenue':'Volume'} — Geographic Contribution`;byId('overviewChartSub').textContent=`${state.scenario} scenario · US + Ex-US`;const alert=byId('overviewAlert');if(alert)alert.style.display='none';const sf=byId('overviewScenarioField');if(sf)sf.style.display='block';
    }
    const fi=yi(state.overviewYear),fv=hoverSeries[0].values[fi]||0;h+=`<line x1="${x(fi)}" y1="${pad.t}" x2="${x(fi)}" y2="${H-pad.b}" stroke="#c8102e" stroke-dasharray="4 4" opacity=".45"/><circle cx="${x(fi)}" cy="${y(fv)}" r="5" fill="#fff" stroke="#c8102e" stroke-width="3"/>`;svg.innerHTML=h;attachLineHover(svg,ensureTip(svg,'overviewTip'),hoverSeries,W,H,pad,max,format)
  };

  window.renderScenarios=function(){
    const d=scenarioData(),s=demoSeries(d.arr),i=yi(state.scenarioYear),base=s.base[i]||0,strong=s.strong[i]||0,weak=s.weak[i]||0,spread=strong-weak;byId('scenarioKpis').innerHTML=[['Base',fmtMetric(base,d.kind),String(state.scenarioYear)],['Strong',fmtMetric(strong,d.kind),'Illustrative'],['Weak',fmtMetric(weak,d.kind),'Illustrative'],['Scenario spread',fmtMetric(spread,d.kind),'Strong − Weak']].map(c=>`<div class="kpi"><div class="kpil">${c[0]}</div><div class="kpiv">${c[1]}</div><div class="kpisub">${c[2]}</div></div>`).join('');byId('scenarioChartTitle').textContent=`${d.title} · Base / Strong / Weak`;const svg=byId('scenarioChart'),W=1100,H=390,pad={l:70,r:20,t:24,b:50},max=Math.max(...s.base,...s.strong,...s.weak)*1.12||1,x=j=>xScale(j,W,pad),y=v=>pad.t+(max-v)/max*(H-pad.t-pad.b);let h=gridSvg(W,H,pad,max,v=>fmtMetric(v,d.kind),5);years.forEach((yr,j)=>{if(j%2===0||j===years.length-1)h+=`<text x="${x(j)}" y="${H-17}" text-anchor="middle" font-size="9" fill="#98a2b3">${yr}</text>`});h+=`<path d="${linePath(s.weak,W,H,pad,max)}" fill="none" stroke="#98a2b3" stroke-width="2.5" stroke-dasharray="6 5"/><path d="${linePath(s.strong,W,H,pad,max)}" fill="none" stroke="#17191d" stroke-width="2.8"/><path d="${linePath(s.base,W,H,pad,max)}" fill="none" stroke="#c8102e" stroke-width="3.2"/>`;const fi=yi(state.scenarioYear);h+=`<line x1="${x(fi)}" y1="${pad.t}" x2="${x(fi)}" y2="${H-pad.b}" stroke="#c8102e" stroke-dasharray="4 4" opacity=".4"/>`;svg.innerHTML=h;attachLineHover(svg,ensureTip(svg,'scenarioTip'),[{name:'Base',color:'#c8102e',values:s.base},{name:'Strong · illustrative',color:'#17191d',values:s.strong},{name:'Weak · illustrative',color:'#98a2b3',values:s.weak}],W,H,pad,max,v=>fmtMetric(v,d.kind));const pk=peak(s.base);byId('scenarioInsight').innerHTML=`<b>The demo shows a visible and governed scenario range.</b> Base peaks in ${pk.year} at ${fmtMetric(pk.value,d.kind)}. Strong and Weak are illustrative paths because the uploaded cached scenario outputs overlap. Hover any year to see all three values.`
  };

  function relabel(){
    document.title='Foresight V24 — Deployable Forecasting Demo';
    const replacements=[['Patient Funnel','Patient Potential Tower'],['Patient Flow','Patient Potential Tower']];
    const walker=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);const nodes=[];while(walker.nextNode())nodes.push(walker.currentNode);nodes.forEach(n=>{let t=n.nodeValue;replacements.forEach(([a,b])=>{t=t.split(a).join(b)});n.nodeValue=t});
    document.querySelectorAll('[data-epi-page="flow"], [data-local-page="flow"]').forEach(el=>el.textContent='Patient Potential Tower');
    const status=byId('v18-engine-text');if(status)status.textContent='Browser Excel generator ready · no local service required';const dot=byId('v18-engine-dot');if(dot)dot.style.background='#12b76a';
  }

  window.downloadSource=function(){if(typeof toast==='function')toast('The source workbook is excluded from this public-safe deployment package.');};

  window.addEventListener('load',()=>{relabel();setTimeout(()=>{relabel();try{if(window.activePlatformSection==='workspace'){renderOverview();renderFlow();renderScenarios()}}catch(e){}},500)});
  const observer=new MutationObserver(()=>{clearTimeout(window.__v24RelabelTimer);window.__v24RelabelTimer=setTimeout(relabel,40)});observer.observe(document.documentElement,{childList:true,subtree:true});
})();
