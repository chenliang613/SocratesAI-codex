import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const deck = await PresentationFile.importPptx(await FileBlob.load("/Users/apple/SocratesAI-codex/行业智能化分析.pptx"));
const snap = await deck.inspect({
  kind: "slide,textbox,shape,notes,layout",
  include: "id,slide,name,title,text,textPreview,bbox,isPlaceholder,placeholders",
  maxChars: 500000,
});
await fs.writeFile("/Users/apple/SocratesAI-codex/tmp/industry-intelligence-update/inspect-large.ndjson", snap.ndjson);
console.log(snap.ndjson.length);
