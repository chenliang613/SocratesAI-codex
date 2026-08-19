import fs from 'node:fs/promises';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const root = path.resolve(process.cwd(), 'tmp/ai-governance-2028/pkg');
const out = path.resolve(process.cwd(), 'AI治理2028.pptx');
const emu = (n) => Math.round(n * 9525);
const esc = (v) => String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const color = (v) => `<a:solidFill><a:srgbClr val="${v.replace('#','')}"/></a:solidFill>`;
const noLine = '<a:ln><a:noFill/></a:ln>';
let id = 2;

function tx(text, x, y, w, h, size, fill='0F172A', opts={}) {
  const {bold=false, font='Microsoft YaHei', align='l', valign='t', margin=0, tracking=0} = opts;
  const runs = String(text).split('\n').map((line, i) => `<a:p><a:pPr algn="${align}" marL="0" marR="0" lvl="0"/>`+
    `<a:r><a:rPr lang="zh-CN" sz="${Math.round(size*100)}" b="${bold?1:0}" kern="0" spc="${tracking}">${color(fill)}<a:latin typeface="${font}"/><a:ea typeface="${font}"/></a:rPr><a:t>${esc(line)}</a:t></a:r><a:endParaRPr lang="zh-CN" sz="${Math.round(size*100)}"/></a:p>`).join('');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${noLine}</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="${emu(margin)}" rIns="${emu(margin)}" tIns="0" bIns="0" anchor="${valign}"/><a:lstStyle/>${runs}</p:txBody></p:sp>`;
}
function rect(x,y,w,h,fill, line='none') { return `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${color(fill)}${line==='none'?noLine:`<a:ln w="${emu(1)}">${color(line)}</a:ln>`}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`; }
function line(x,y,w,h,fill='B8BCC4') { return `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="Rule ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${color(fill)}${noLine}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`; }
function circle(x,y,d,fill) { return `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="Circle ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(d)}" cy="${emu(d)}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>${color(fill)}${noLine}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`; }

const shapes = [];
shapes.push(rect(0,0,1280,720,'FFFFFF'));
shapes.push(rect(72,56,12,56,'3D8DFF'));
shapes.push(tx('AI GOVERNANCE  /  2028',104,54,450,24,14,'3D8DFF',{bold:true,tracking:120}));
shapes.push(tx('AI治理2028：从“安全护栏”走向AI规模化的操作系统',72,93,1136,55,30,'111827',{bold:true}));
shapes.push(tx('未来2–3年，治理将同时成为产品信任的竞争力、产业扩张的基础设施，以及贯穿数据—模型—智能体—应用的运营闭环。',72,158,1115,28,15,'475569'));
shapes.push(line(72,205,1136,2,'111827'));

const cols=[72,452,832];
const heads=['01  竞争力：从“合规项”到“产品能力”','02  发展力：从“约束”到“产业基础设施”','03  生命力：从“单点审查”到“全生命周期”'];
const big=['可信能力将成为厂商\n规模化交付的门槛','治理机制将释放\n可复制、可推广的应用','风险控制将嵌入\n每一次训练、调用与迭代'];
const detail=['安全、评测、红队、内容溯源与权限控制\n被打包为可销售、可验证、可运营的能力。','标准、工具链、行业规则与监管沙盒协同演进；\n在发展中治理，在治理中发展。','覆盖数据、模型、智能体、应用、运行监测与\n责任追溯；由“上线前”延伸至“运行中”。'];
for(let i=0;i<3;i++){
 shapes.push(rect(cols[i],236,332,8,'3D8DFF'));
 shapes.push(tx(heads[i],cols[i],260,332,40,15,'3D8DFF',{bold:true}));
 shapes.push(tx(big[i],cols[i],314,332,66,23,'111827',{bold:true}));
 shapes.push(line(cols[i],398,332,1,'CBD5E1'));
 shapes.push(tx(detail[i],cols[i],416,332,70,15,'475569'));
}
shapes.push(rect(72,515,1136,104,'F1F5F9'));
shapes.push(tx('治理能力的演进路径',94,534,200,20,14,'475569',{bold:true}));
shapes.push(line(314,565,794,2,'94A3B8'));
const years=[['2026','建立治理基线','规则、评测、责任可审计'],['2027','嵌入生产体系','治理进入开发、交付与运行'],['2028','形成价值闭环','以可信驱动规模化与生态协同']];
const nodes=[330,604,878];
for(let i=0;i<3;i++){
 shapes.push(circle(nodes[i],550,30,'3D8DFF'));
 shapes.push(tx(years[i][0],nodes[i]-6,554,42,20,11,'FFFFFF',{bold:true,align:'c'}));
 shapes.push(tx(years[i][1],nodes[i]-10,587,210,22,15,'111827',{bold:true}));
 shapes.push(tx(years[i][2],nodes[i]-10,609,230,18,12,'64748B'));
}
shapes.push(tx('企业行动：将治理能力产品化、工程化、运营化，并以业务价值和可信指标共同衡量。',72,648,1136,26,17,'111827',{bold:true}));
shapes.push(tx('资料依据：用户提供的AI治理三项定义；本页为2026–2028前瞻性判断。',72,688,800,14,10,'94A3B8'));

const slide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
const content = {
 '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
 '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`,
 'ppt/presentation.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle/></p:presentation>`,
 'ppt/_rels/presentation.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`,
 'ppt/slides/slide1.xml': slide,
 'ppt/slides/_rels/slide1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>`,
 'ppt/notesSlides/notesSlide1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${tx('[Sources]\\n- 用户提供的三项AI治理定义。\\n- 本页为2026–2028前瞻性判断，非统计预测。',0,0,600,300,12,'000000')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
 'ppt/notesSlides/_rels/notesSlide1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/></Relationships>`,
 'docProps/core.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>AI治理2028</dc:title><dc:creator>Codex</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-08-19T00:00:00Z</dcterms:created></cp:coreProperties>`,
 'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office PowerPoint</Application><Slides>1</Slides><Notes>1</Notes></Properties>`
};

async function main(){
 await fs.rm(root,{recursive:true,force:true}); await fs.mkdir(root,{recursive:true});
 for(const [rel,data] of Object.entries(content)){const f=path.join(root,rel); await fs.mkdir(path.dirname(f),{recursive:true}); await fs.writeFile(f,data);}
 await fs.rm(out,{force:true}); execFileSync('zip',['-X','-q','-r',out,'.'],{cwd:root}); console.log(out);
}
main().catch(e=>{console.error(e);process.exit(1)});
