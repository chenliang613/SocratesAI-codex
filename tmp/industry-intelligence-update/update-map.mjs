import fs from "node:fs/promises";

const dir = "/Users/apple/SocratesAI-codex/tmp/industry-intelligence-update";
const map = JSON.parse(await fs.readFile(`${dir}/template-frame-map.json`, "utf8"));
const records = (await fs.readFile(`${dir}/inspect-large.ndjson`, "utf8")).trim().split(/\n/).map(JSON.parse);
const bySlide = new Map();
for (const r of records) {
  if (!r.slide || !r.id?.startsWith("sh/")) continue;
  if (!bySlide.has(r.slide)) bySlide.set(r.slide, []);
  bySlide.get(r.slide).push(r);
}

for (const entry of map.outputSlides) {
  if (entry.editTargets.length === 0) {
    entry.narrativeRole = "preserve-only";
    continue;
  }
  const layout = JSON.parse(await fs.readFile(`${dir}/template-inspect/layouts/source-slide-${String(entry.sourceSlide).padStart(2,"0")}.layout.json`, "utf8"));
  const layoutById = new Map(layout.elements.map(e => [String(e.id), e]));
  for (const t of entry.editTargets) {
    const src = layoutById.get(String(t.sourceElementId));
    if (!src) throw new Error(`Missing layout element ${entry.sourceSlide}/${t.sourceElementId}`);
    const candidates = bySlide.get(entry.sourceSlide) || [];
    const hit = candidates.find(r => r.name === src.name && JSON.stringify(r.bbox) === JSON.stringify(src.bbox));
    if (!hit) throw new Error(`Missing inspect shape for ${entry.sourceSlide}/${src.name}/${JSON.stringify(src.bbox)}`);
    t.sourceElementId = hit.id;
  }
}
await fs.writeFile(`${dir}/template-frame-map.json`, JSON.stringify(map, null, 2) + "\n");
