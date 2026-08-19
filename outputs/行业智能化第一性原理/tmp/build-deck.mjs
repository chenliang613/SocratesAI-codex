import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const workspaceRoot = "/Users/apple/SocratesAI-codex/outputs/行业智能化第一性原理";
const imageDir = path.join(workspaceRoot, "imagegen", "images");
const outputDir = path.join(workspaceRoot, "output");
const finalPptx = "/Users/apple/SocratesAI-codex/行业智能化第一性原理.pptx";
const workspacePptx = path.join(outputDir, "行业智能化第一性原理.pptx");

const titles = [
  "行业智能化第一性原理",
  "先回答为什么必然发生，再回答构建什么、如何规模复制",
  "行业不同，产业化规律相同",
  "六条原理构成一条从价值到规模的闭环",
  "原理一：正 ROI 是智能化项目的生存线",
  "原理二：统一业务语义，Agent 才能跨系统工作",
  "原理三：通用模型决定地板，行业知识决定天花板",
  "原理四：PoC、标杆、复制分别验证三件不同的事",
  "原理五：不能模板化的场景，只能做项目，不能做产业",
  "原理六：第 N 次交付成本，决定应用层价值倍数",
  "企业软件竞争焦点正在从界面迁移到业务逻辑",
  "目标架构：让意图通过语义与业务逻辑，稳定变成结果",
  "执行策略：把一次成功沉淀为可重复的复制资产",
  "行业智能化的终局，是可度量、可执行、可复制",
];

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function readImageBytes(filePath) {
  const bytes = await fs.readFile(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function main() {
  await fs.mkdir(outputDir, { recursive: true });
  const deck = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });

  for (let index = 0; index < 14; index += 1) {
    const slideNumber = String(index + 1).padStart(2, "0");
    const imagePath = path.join(imageDir, `slide-${slideNumber}.png`);
    const imageBytes = await readImageBytes(imagePath);
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    slide.images.add({
      blob: imageBytes,
      contentType: "image/png",
      alt: `第${index + 1}页：${titles[index]}`,
      fit: "cover",
      position: { left: 0, top: 0, width: 1280, height: 720 },
    });
    slide.speakerNotes.textFrame.setText(
      `[Sources]\nS01 | /Users/apple/SocratesAI-codex/Input Minds.md | 用户提供材料；本页观点与文案基于该材料重组。`,
    );
    slide.speakerNotes.setVisible(false);
  }

  const montage = await deck.export({
    format: "webp",
    montage: true,
    scale: 1,
  });
  await writeBlob(path.join(outputDir, "contact-sheet.webp"), montage);

  const pptx = await PresentationFile.exportPptx(deck);
  await pptx.save(workspacePptx);
  await pptx.save(finalPptx);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
