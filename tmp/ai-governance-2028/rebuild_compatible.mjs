import fs from 'node:fs/promises';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const work = path.resolve(process.cwd(), 'tmp/ai-governance-2028/compatible');
const source = path.resolve(process.cwd(), '微软AI治理分析.pptx');
const draft = path.resolve(process.cwd(), 'tmp/ai-governance-2028/pkg');
const finalFile = path.resolve(process.cwd(), 'AI治理2028.pptx');

async function text(f){ return fs.readFile(f, 'utf8'); }
async function main(){
  await fs.rm(work,{recursive:true,force:true}); await fs.mkdir(work,{recursive:true});
  execFileSync('unzip',['-q',source,'-d',work]);
  await fs.copyFile(path.join(draft,'ppt/slides/slide1.xml'),path.join(work,'ppt/slides/slide1.xml'));
  const presentationPath=path.join(work,'ppt/presentation.xml');
  let presentation=await text(presentationPath);
  presentation=presentation.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst><p:sldId id="256" r:id="R370fb6bdcafd49fa" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" /></p:sldIdLst>');
  await fs.writeFile(presentationPath,presentation);
  const relPath=path.join(work,'ppt/_rels/presentation.xml.rels');
  let rels=await text(relPath);
  rels=rels.replace(/<Relationship Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide" Target="\/ppt\/slides\/slide(?:[2-9]|1[0-9])\.xml" Id="[^"]+" \/>/g,'');
  await fs.writeFile(relPath,rels);
  const typesPath=path.join(work,'[Content_Types].xml');
  let types=await text(typesPath);
  types=types.replace(/<Override PartName="\/ppt\/slides\/slide(?:[2-9]|1[0-9])\.xml"[^>]*\/>/g,'').replace(/<Override PartName="\/ppt\/notesSlides\/notesSlide(?:[2-9]|1[0-9])\.xml"[^>]*\/>/g,'');
  await fs.writeFile(typesPath,types);
  const corePath=path.join(work,'docProps/core.xml');
  let core=await text(corePath);
  core=core.replace(/<dc:title>.*?<\/dc:title>/,'<dc:title>AI治理2028</dc:title>');
  await fs.writeFile(corePath,core);
  for(let n=2;n<=15;n++){
    await fs.rm(path.join(work,`ppt/slides/slide${n}.xml`),{force:true});
    await fs.rm(path.join(work,`ppt/slides/_rels/slide${n}.xml.rels`),{force:true});
    await fs.rm(path.join(work,`ppt/notesSlides/notesSlide${n}.xml`),{force:true});
    await fs.rm(path.join(work,`ppt/notesSlides/_rels/notesSlide${n}.xml.rels`),{force:true});
  }
  await fs.rm(finalFile,{force:true});
  execFileSync('zip',['-X','-q','-r',finalFile,'.'],{cwd:work});
  console.log(finalFile);
}
main().catch(err=>{console.error(err);process.exit(1)});
