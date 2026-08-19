import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "/Users/apple/SocratesAI-codex/微软AI治理分析.pptx";
const PREVIEW_DIR = "/Users/apple/SocratesAI-codex/tmp/ms_ai_governance_rebuild/previews";

const W = 1280;
const H = 720;
const FONT = "PingFang SC";
const C = {
  ink: "#111111",
  muted: "#5F6670",
  rule: "#B8BCC4",
  panel: "#F1F2F4",
  panel2: "#E8EDF2",
  blue: "#0078D4",
  blue2: "#50A7E0",
  bluePale: "#E7F3FB",
  green: "#2E7D32",
  greenPale: "#EAF5EA",
  amber: "#B26A00",
  amberPale: "#FFF3D6",
  red: "#C62828",
  redPale: "#FCE8E8",
  white: "#FFFFFF",
};

const URL = {
  euAct: "https://eur-lex.europa.eu/eli/reg/2024/1689/oj",
  euNavigate: "https://digital-strategy.ec.europa.eu/en/faqs/navigating-ai-act",
  euFramework: "https://digital-strategy.ec.europa.eu/en/policies/regulatory-framework-ai",
  euGpai: "https://digital-strategy.ec.europa.eu/en/policies/contents-code-gpai",
  euGpaiGuide: "https://digital-strategy.ec.europa.eu/en/policies/guidelines-gpai-providers",
  euTransparency: "https://digital-strategy.ec.europa.eu/en/policies/guidelines-transparency-ai-generated-content",
  euTransparencyCode: "https://digital-strategy.ec.europa.eu/en/faqs/code-practice-transparency-ai-generated-content",
  msReport: "https://www.microsoft.com/en-us/corporate-responsibility/responsible-ai-transparency-report/",
  msAssurance: "https://learn.microsoft.com/en-us/compliance/assurance/assurance-artificial-intelligence",
  msTrust: "https://www.microsoft.com/en-us/trust-center/compliance/eu-ai-act",
  msActBlog: "https://blogs.microsoft.com/on-the-issues/2025/01/15/innovating-in-line-with-the-european-unions-ai-act/",
  msStandard: "https://cdn-dynmedia-1.microsoft.com/is/content/microsoftcorp/microsoft/bade/documents/products-and-services/en-us/ai/RAIS-Reference-Guide-v2.pdf",
  msContentSafety: "https://learn.microsoft.com/en-us/azure/ai-services/content-safety/overview",
  msPurview: "https://learn.microsoft.com/en-sg/purview/compliance-manager-assessments",
  msShared: "https://learn.microsoft.com/en-us/azure/security/fundamentals/shared-responsibility-ai",
};

function addText(slide, name, text, left, top, width, height, opts = {}) {
  const box = slide.shapes.add({
    geometry: "textbox",
    name,
    position: { left, top, width, height },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  box.text = text;
  box.text.style = {
    typeface: FONT,
    fontSize: opts.fontSize ?? 22,
    color: opts.color ?? C.ink,
    bold: opts.bold ?? false,
    alignment: opts.alignment ?? "left",
    verticalAlignment: opts.verticalAlignment ?? "top",
    autoFit: opts.autoFit ?? "shrinkText",
  };
  return box;
}

function addRect(slide, name, left, top, width, height, fill = C.panel, lineFill = "none", lineWidth = 0, radius = false) {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect",
    name,
    position: { left, top, width, height },
    fill,
    line: { style: "solid", fill: lineFill, width: lineWidth },
  });
}

function addLine(slide, name, left, top, width, height = 0, color = C.rule, lineWidth = 2, arrow = false) {
  return slide.shapes.add({
    geometry: "straightConnector1",
    name,
    position: { left, top, width, height },
    fill: "none",
    line: {
      style: "solid",
      fill: color,
      width: lineWidth,
      ...(arrow ? { endArrowType: "triangle" } : {}),
    },
  });
}

function addCircle(slide, name, cx, cy, diameter, fill, lineFill = fill) {
  return slide.shapes.add({
    geometry: "ellipse",
    name,
    position: { left: cx - diameter / 2, top: cy - diameter / 2, width: diameter, height: diameter },
    fill,
    line: { style: "solid", fill: lineFill, width: 1 },
  });
}

function addChrome(slide, title, section, page) {
  slide.background.fill = C.white;
  addText(slide, `section-${page}`, section, 42, 26, 380, 24, { fontSize: 15, bold: true, color: C.blue });
  addText(slide, `title-${page}`, title, 42, 60, 1120, 70, { fontSize: 48, bold: true, autoFit: "none" });
  addLine(slide, `title-rule-${page}`, 42, 142, 1196, 0, C.rule, 1);
  addText(slide, `footer-${page}`, String(page).padStart(2, "0"), 1185, 670, 53, 22, { fontSize: 14, color: C.muted, alignment: "right" });
}

function notes(slide, sources, extra = "") {
  const lines = ["[Sources]", ...sources.map((s) => `- ${s}`)];
  if (extra) lines.push("", extra);
  slide.speakerNotes.textFrame.setText(lines.join("\n"));
  slide.speakerNotes.setVisible(true);
}

function addPanelText(slide, prefix, x, y, w, h, header, body, tone = "neutral") {
  const fill = tone === "blue" ? C.bluePale : tone === "green" ? C.greenPale : tone === "amber" ? C.amberPale : tone === "red" ? C.redPale : C.panel;
  addRect(slide, `${prefix}-panel`, x, y, w, h, fill);
  addText(slide, `${prefix}-header`, header, x + 24, y + 22, w - 48, 42, { fontSize: 32, bold: true, color: tone === "red" ? C.red : tone === "amber" ? C.amber : tone === "green" ? C.green : tone === "blue" ? C.blue : C.ink });
  addText(slide, `${prefix}-body`, body, x + 24, y + 78, w - 48, h - 100, { fontSize: 21, color: C.ink });
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function styleTable(table, rows, cols, options = {}) {
  table.borders.assign({ style: "solid", fill: C.rule, width: 1 });
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r, c);
      cell.fill = r === 0 ? C.panel2 : (r % 2 === 0 ? "#FAFAFA" : C.white);
      cell.text.style = {
        typeface: FONT,
        fontSize: r === 0 ? (options.headerSize ?? 18) : (options.bodySize ?? 17),
        bold: r === 0 || (options.boldFirstColumn && c === 0),
        color: C.ink,
        verticalAlignment: "middle",
      };
    }
  }
}

async function main() {
  await fs.mkdir(PREVIEW_DIR, { recursive: true });
  const deck = Presentation.create({ slideSize: { width: W, height: H } });

  // 1 — Cover, adapted from Codex Grid stacked-text-flow.
  {
    const s = deck.slides.add();
    s.background.fill = C.white;
    addRect(s, "cover-accent", 42, 144, 18, 282, C.blue);
    addText(s, "cover-eyebrow", "MANAGEMENT BRIEF · 2026.08.04", 42, 42, 460, 36, { fontSize: 20, bold: true, color: C.blue });
    addText(s, "cover-date", "欧盟 AI 法案执行期分析", 850, 42, 388, 36, { fontSize: 20, alignment: "right", color: C.muted });
    addText(s, "cover-title", "微软 AI 治理战略", 88, 176, 1000, 108, { fontSize: 80, bold: true, autoFit: "none" });
    addText(s, "cover-subtitle", "如何满足欧盟《AI 法案》", 88, 300, 1000, 78, { fontSize: 52, bold: true, color: C.blue, autoFit: "none" });
    addText(s, "cover-desc", "治理架构 · 义务映射 · 共享责任 · 执行路线图", 88, 506, 820, 46, { fontSize: 28, color: C.muted });
    addText(s, "cover-scope", "基于公开材料的战略与控制分析，不构成法律意见或独立审计结论。", 88, 572, 850, 40, { fontSize: 18, color: C.muted });
    notes(s, [
      `European Commission — AI Act regulatory framework: ${URL.euFramework}`,
      `Microsoft — 2025 Responsible AI Transparency Report: ${URL.msReport}`,
    ]);
  }

  // 2 — Executive conclusion, adapted from metric-led layout.
  {
    const s = deck.slides.add();
    addChrome(s, "结论：治理底座成熟，胜负在证据链与责任边界", "EXECUTIVE CONCLUSION", 2);
    addText(s, "s2-lead", "微软已把原则转化为公司级政策、工程门禁与风险工具；但“使用微软产品”不等于自动满足《AI 法案》。", 42, 166, 1196, 66, { fontSize: 27, bold: true });
    addPanelText(s, "s2-a", 42, 256, 374, 340, "已具备", "• 董事会与高管监督\n• Responsible AI Council / ORA / Aether\n• Responsible AI Standard\n• 发布前安全审查与敏感用途复核\n• 持续监测与纵深防御", "green");
    addPanelText(s, "s2-b", 453, 256, 374, 340, "正在兑现", "• 禁止用途纳入 Restricted Use Policy\n• 中央筛查与文档工作流\n• 签署 GPAI 行为准则\n• 30 项工具、155+ 功能\n• 40 份 Transparency Notes", "blue");
    addPanelText(s, "s2-c", 864, 256, 374, 340, "仍需闭环", "• 第 50 条内容标记的端到端继承\n• 高风险 QMS / 合格评定证据\n• 微软—模型伙伴—客户的角色认定\n• 客户场景中的 FRIA 与人工监督\n• 跨主体严重事件响应", "amber");
    addText(s, "s2-foot", "判断：战略方向与法案高度同构；合规成熟度取决于每个版本、场景和责任主体能否形成可验证记录。", 42, 620, 1120, 38, { fontSize: 20, color: C.muted });
    notes(s, [
      `Microsoft — Responsible AI Transparency Report (govern/map/measure/manage; tools and Transparency Notes): ${URL.msReport}`,
      `Microsoft — Artificial Intelligence assurance overview (governance roles): ${URL.msAssurance}`,
      `Microsoft Trust Center — EU AI Act compliance approach: ${URL.msTrust}`,
    ], "The readiness characterization and gap conclusions are analytical judgments based on public evidence, not audit findings.");
  }

  // 3 — EU timeline.
  {
    const s = deck.slides.add();
    addChrome(s, "法律已进入执行期，高风险义务窗口被延长", "EU AI ACT · CURRENT STATE", 3);
    addRect(s, "s3-current-tag", 998, 164, 240, 44, C.blue);
    addText(s, "s3-current-text", "当前：执法已经开始", 1012, 173, 212, 28, { fontSize: 19, bold: true, color: C.white, alignment: "center" });
    addLine(s, "s3-axis", 80, 368, 1120, 0, C.rule, 3);
    const events = [
      { x: 95, date: "2025.02.02", title: "禁止实践 / AI 素养", body: "定义、禁止性规则与第 4 条开始适用", top: true, color: C.green },
      { x: 310, date: "2025.08.02", title: "GPAI 义务", body: "治理、通用模型义务与处罚框架", top: false, color: C.green },
      { x: 525, date: "2026.08.02", title: "透明度与执法", body: "第 50 条适用；AI Office 执法权生效", top: true, color: C.blue },
      { x: 740, date: "2026.12.02", title: "旧系统标记宽限", body: "此前上市的生成式 AI 完成机器可读标记", top: false, color: C.amber },
      { x: 955, date: "2027.12.02", title: "附录 III 高风险", body: "就业、教育、关键设施等敏感用途", top: true, color: C.amber },
      { x: 1170, date: "2028.08.02", title: "附录 I 产品高风险", body: "嵌入医疗器械、玩具、电梯等产品", top: false, color: C.amber },
    ];
    for (let i = 0; i < events.length; i++) {
      const e = events[i];
      addCircle(s, `s3-node-${i}`, e.x, 368, i === 2 ? 26 : 18, e.color);
      const boxX = Math.max(42, Math.min(1038, e.x - 100));
      const y = e.top ? 226 : 410;
      addText(s, `s3-date-${i}`, e.date, boxX, y, 200, 28, { fontSize: 20, bold: true, color: e.color, alignment: "center" });
      addText(s, `s3-title-${i}`, e.title, boxX, y + 32, 200, 30, { fontSize: 21, bold: true, alignment: "center" });
      addText(s, `s3-body-${i}`, e.body, boxX, y + 68, 200, 58, { fontSize: 16, color: C.muted, alignment: "center" });
    }
    addText(s, "s3-omnibus", "AI Omnibus 于 2026.07.27 生效：高风险日期后移，但透明度、GPAI 和执法节点不后移。", 42, 624, 1140, 40, { fontSize: 21, bold: true, color: C.blue });
    notes(s, [
      `European Commission — Navigating the AI Act (Omnibus dates and enforcement): ${URL.euNavigate}`,
      `European Commission — Transparency rules quick guidance: ${URL.euTransparency}`,
      `EUR-Lex — Regulation (EU) 2024/1689: ${URL.euAct}`,
    ]);
  }

  // 4 — Regulatory logic.
  {
    const s = deck.slides.add();
    addChrome(s, "合规不是单一清单，而是“风险层级 + 供应链角色”", "REGULATORY LOGIC", 4);
    const rows = [
      { y: 172, h: 82, label: "不可接受风险", color: C.red, fill: C.redPale, text: "禁止设计、投放或使用；必须以用途筛查、合同限制和升级机制阻断。" },
      { y: 266, h: 118, label: "高风险系统", color: C.amber, fill: C.amberPale, text: "风险管理、数据治理、技术文档、日志、透明度、人工监督、准确性/稳健性/网络安全、QMS、合格评定、登记、上市后监测。" },
      { y: 396, h: 82, label: "透明度风险", color: C.blue, fill: C.bluePale, text: "告知人与 AI 交互；机器可读标记；深度伪造、情绪识别与公共利益文本披露。" },
      { y: 490, h: 82, label: "最小风险", color: C.green, fill: C.greenPale, text: "原则上可自愿采用行为准则；仍受 GDPR、消费者保护、知识产权等横向法律约束。" },
    ];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      addRect(s, `s4-band-${i}`, 42, r.y, 1196, r.h, r.fill);
      addRect(s, `s4-color-${i}`, 42, r.y, 12, r.h, r.color);
      addText(s, `s4-label-${i}`, r.label, 78, r.y + 20, 220, 40, { fontSize: 28, bold: true, color: r.color });
      addText(s, `s4-text-${i}`, r.text, 330, r.y + 17, 870, r.h - 24, { fontSize: 21, verticalAlignment: "middle" });
    }
    addRect(s, "s4-gpai", 42, 598, 1196, 54, C.ink);
    addText(s, "s4-gpai-text", "独立的 GPAI 层：所有通用模型承担透明度 / 版权义务；系统性风险模型另承担评估、缓解、事件与网络安全义务。", 65, 611, 1150, 30, { fontSize: 20, bold: true, color: C.white, alignment: "center" });
    notes(s, [
      `European Commission — AI Act regulatory framework and risk levels: ${URL.euFramework}`,
      `European Commission — Navigating the AI Act (high-risk requirements): ${URL.euNavigate}`,
      `European Commission — GPAI provider guidelines: ${URL.euGpaiGuide}`,
    ]);
  }

  // 5 — Governance architecture, simple native diagram with connectors created first.
  {
    const s = deck.slides.add();
    addChrome(s, "微软用“中央定标 + 分布式问责”把原则嵌入工程", "MICROSOFT GOVERNANCE", 5);
    // Connectors first.
    addLine(s, "s5-v1", 640, 232, 0, 32, C.rule, 2);
    addLine(s, "s5-v2", 640, 342, 0, 42, C.rule, 2);
    addLine(s, "s5-hubline", 250, 384, 780, 0, C.rule, 2);
    for (const x of [250, 640, 1030]) addLine(s, `s5-down-${x}`, x, 384, 0, 28, C.rule, 2);
    addLine(s, "s5-spoke-line", 250, 512, 780, 0, C.rule, 2);
    for (const x of [250, 640, 1030]) addLine(s, `s5-up-${x}`, x, 490, 0, 22, C.rule, 2);
    addLine(s, "s5-bottom", 640, 512, 0, 40, C.rule, 2, true);

    addRect(s, "s5-board", 390, 166, 500, 66, C.panel);
    addText(s, "s5-board-t", "董事会 / CEO / 高管团队", 414, 178, 452, 28, { fontSize: 27, bold: true, alignment: "center" });
    addText(s, "s5-board-b", "董事会环境、社会与公共政策委员会监督", 414, 207, 452, 18, { fontSize: 15, color: C.muted, alignment: "center" });
    addRect(s, "s5-council", 350, 264, 580, 78, C.blue);
    addText(s, "s5-council-t", "Responsible AI Council", 378, 278, 524, 34, { fontSize: 32, bold: true, color: C.white, alignment: "center" });
    addText(s, "s5-council-b", "跨研究、政策、工程的重大问题决策与一致性执行", 378, 315, 524, 22, { fontSize: 17, color: C.white, alignment: "center" });
    addRect(s, "s5-ora", 90, 412, 320, 78, C.bluePale);
    addText(s, "s5-ora-t", "ORA", 108, 425, 284, 28, { fontSize: 27, bold: true, color: C.blue, alignment: "center" });
    addText(s, "s5-ora-b", "公司政策 · 治理流程 · 敏感用途复核", 108, 458, 284, 20, { fontSize: 15, alignment: "center" });
    addRect(s, "s5-aether", 480, 412, 320, 78, C.panel);
    addText(s, "s5-aether-t", "Aether / Research", 498, 425, 284, 28, { fontSize: 27, bold: true, alignment: "center" });
    addText(s, "s5-aether-b", "前沿风险 · 社会技术研究 · 技术建议", 498, 458, 284, 20, { fontSize: 15, alignment: "center" });
    addRect(s, "s5-control", 870, 412, 320, 78, C.panel);
    addText(s, "s5-control-t", "Legal / Security / Policy", 888, 425, 284, 28, { fontSize: 24, bold: true, alignment: "center" });
    addText(s, "s5-control-b", "法规解释 · 安全门禁 · 合同与公共政策", 888, 458, 284, 20, { fontSize: 15, alignment: "center" });
    addRect(s, "s5-engineering", 150, 552, 980, 74, C.panel2);
    addText(s, "s5-engineering-t", "产品工程团队 + Division Leads + Responsible AI Champs", 178, 568, 924, 30, { fontSize: 28, bold: true, alignment: "center" });
    addText(s, "s5-engineering-b", "在模型、平台、应用各层承担落地责任，并把问题与例外向中央治理升级", 178, 602, 924, 22, { fontSize: 17, color: C.muted, alignment: "center" });
    notes(s, [
      `Microsoft Learn — Artificial Intelligence assurance overview (Board, Council, ORA, Aether, engineering roles): ${URL.msAssurance}`,
      `Microsoft — Responsible AI governance and transparency: ${URL.msReport}`,
    ]);
  }

  // 6 — Lifecycle alignment.
  {
    const s = deck.slides.add();
    addChrome(s, "Govern—Map—Measure—Manage 与法案生命周期高度同构", "CONTROL LIFECYCLE", 6);
    const xs = [42, 347, 652, 957];
    for (let i = 0; i < 3; i++) addLine(s, `s6-arrow-${i}`, xs[i] + 258, 348, 47, 0, C.blue, 3, true);
    const steps = [
      { k: "GOVERN", zh: "治理", body: "政策、角色、培训、文档工作流", law: "QMS · 问责 · AI 素养" },
      { k: "MAP", zh: "识别", body: "用途、利益相关方、风险分类与影响", law: "禁止实践 · 高风险分类 · FRIA" },
      { k: "MEASURE", zh: "度量", body: "安全、公平、稳健性、红队与评估", law: "准确性 · 稳健性 · 网络安全" },
      { k: "MANAGE", zh: "处置", body: "缓解、人工监督、监控、反馈与事件", law: "上市后监测 · 纠正 · 严重事件" },
    ];
    for (let i = 0; i < steps.length; i++) {
      const x = xs[i];
      addRect(s, `s6-box-${i}`, x, 220, 258, 310, i === 0 ? C.bluePale : C.panel);
      addText(s, `s6-k-${i}`, steps[i].k, x + 22, 242, 214, 30, { fontSize: 18, bold: true, color: C.blue });
      addText(s, `s6-zh-${i}`, steps[i].zh, x + 22, 283, 214, 54, { fontSize: 40, bold: true });
      addText(s, `s6-body-${i}`, steps[i].body, x + 22, 354, 214, 74, { fontSize: 21 });
      addLine(s, `s6-rule-${i}`, x + 22, 444, 214, 0, C.rule, 1);
      addText(s, `s6-law-${i}`, steps[i].law, x + 22, 461, 214, 52, { fontSize: 17, bold: true, color: C.muted });
    }
    addRect(s, "s6-bottom", 42, 574, 1196, 70, C.ink);
    addText(s, "s6-bottom-text", "微软优势：发布前审查已制度化；关键要求：把同一证据链延伸到版本变更、客户配置与上线后事件。", 68, 593, 1144, 32, { fontSize: 22, bold: true, color: C.white, alignment: "center" });
    notes(s, [
      `Microsoft — 2025 Responsible AI Transparency Report (NIST AI RMF functions; pre-deployment and ongoing monitoring): ${URL.msReport}`,
      `Microsoft Responsible AI Standard v2 Reference Guide: ${URL.msStandard}`,
      `European Commission — High-risk system obligations: ${URL.euNavigate}`,
    ]);
  }

  // 7 — Prohibited practices and literacy.
  {
    const s = deck.slides.add();
    addChrome(s, "禁止实践与 AI 素养：最强适配项仍需可审计证明", "NEAR-TERM OBLIGATIONS", 7);
    addRect(s, "s7-callout", 42, 166, 1196, 58, C.greenPale);
    addText(s, "s7-callout-text", "适配度：高　|　法案已适用　|　客户改变预期用途时，风险分类必须重做", 66, 181, 1148, 28, { fontSize: 22, bold: true, color: C.green, alignment: "center" });
    addText(s, "s7-left-title", "微软公开控制", 42, 258, 540, 42, { fontSize: 32, bold: true });
    addText(s, "s7-left-body", "01　Restricted Use Policy 纳入法案禁止实践\n\n02　对已上市系统进行集中盘点与筛查\n\n03　中央工作流保留问题、复核与升级记录\n\n04　更新合同，限制客户不当使用\n\n05　员工在开发 / 部署前查阅政策并接受培训", 42, 320, 540, 288, { fontSize: 22 });
    addLine(s, "s7-divider", 620, 250, 0, 378, C.rule, 1);
    addText(s, "s7-right-title", "审计需要看到的证据", 666, 258, 572, 42, { fontSize: 32, bold: true, color: C.blue });
    addText(s, "s7-right-body", "• 系统清册是否覆盖模型、平台、应用与定制代理\n\n• 禁止用途筛查是否随版本与用途变化触发\n\n• 角色化 AI 素养课程、完成率与能力评估\n\n• 例外申请、拒绝记录与高层升级轨迹\n\n• 合同限制能否被技术策略和监控共同执行", 666, 320, 548, 288, { fontSize: 22 });
    addText(s, "s7-foot", "判断：微软公司级政策设计足以支撑第 5 条与第 4 条；合规风险主要转向覆盖范围和持续执行。", 42, 630, 1140, 28, { fontSize: 19, bold: true, color: C.muted });
    notes(s, [
      `Microsoft Trust Center — Prohibited practices and Restricted Use Policy: ${URL.msTrust}`,
      `Microsoft — Innovating in line with the EU AI Act (screening, contracts, central workflow): ${URL.msActBlog}`,
      `European Commission — AI Act timeline and AI literacy: ${URL.euNavigate}`,
    ]);
  }

  // 8 — High-risk mapping table.
  {
    const s = deck.slides.add();
    addChrome(s, "高风险系统：控制组件齐全，合格评定证据尚需产品化", "HIGH-RISK MAPPING", 8);
    addText(s, "s8-sub", "高风险规则已延至 2027/2028，但现在应按目标状态做干运行；下表为公开证据评估。", 42, 158, 1130, 34, { fontSize: 20, color: C.muted });
    const values = [
      ["义务域", "欧盟要求", "微软公开控制", "判断"],
      ["风险管理", "全生命周期风险管理体系", "RAI Standard；Map/Measure/Manage；发布前审查", "较强"],
      ["数据治理", "训练/验证/测试数据质量与代表性", "隐私、安全、Purview 与工程实践", "需场景证据"],
      ["文档与透明度", "技术文档、使用说明、下游信息", "Transparency Notes；集中式文档工作流", "较强"],
      ["日志与可追溯", "自动日志、记录保存、版本追踪", "Azure 日志与监控能力；客户需配置留存", "共享责任"],
      ["人工监督", "可理解、可干预、具备授权与能力", "RAI Standard；产品级人工控制；客户流程", "共享责任"],
      ["稳健与安全", "准确性、鲁棒性、网络安全", "红队、PyRIT、Content Safety、Secure Future", "较强"],
      ["QMS / 合格评定", "质量体系、评估、登记、上市后监测", "跨职能合规工作组；标准仍在演进", "关键缺口"],
    ];
    const table = s.tables.add({ rows: values.length, columns: 4, left: 42, top: 210, width: 1196, height: 420, columnWidths: [170, 320, 526, 180], values });
    styleTable(table, values.length, 4, { headerSize: 19, bodySize: 17, boldFirstColumn: true });
    const tones = [null, C.greenPale, C.amberPale, C.greenPale, C.bluePale, C.bluePale, C.greenPale, C.redPale];
    for (let r = 1; r < values.length; r++) {
      table.getCell(r, 3).fill = tones[r];
      table.getCell(r, 3).text.style = { typeface: FONT, fontSize: 17, bold: true, color: r === 7 ? C.red : r === 2 ? C.amber : r === 4 || r === 5 ? C.blue : C.green, verticalAlignment: "middle", alignment: "center" };
    }
    addText(s, "s8-foot", "“较强”表示存在公开控制证据，不代表特定产品已完成欧盟合格评定。", 42, 642, 1120, 28, { fontSize: 17, color: C.muted });
    notes(s, [
      `European Commission — High-risk provider and deployer obligations: ${URL.euNavigate}`,
      `Microsoft — Artificial Intelligence assurance overview: ${URL.msAssurance}`,
      `Microsoft — Responsible AI Transparency Report: ${URL.msReport}`,
      `Microsoft — EU AI Act implementation approach and tools: ${URL.msActBlog}`,
    ], "Assessment labels are analytical judgments based on publicly available evidence; they are not certifications.");
  }

  // 9 — GPAI.
  {
    const s = deck.slides.add();
    addChrome(s, "GPAI：签署行为准则是优势，模型供应链仍是核心考题", "GENERAL-PURPOSE AI", 9);
    addText(s, "s9-lead", "微软既可能是 Phi 等模型的提供者，也可能是模型平台、集成商或应用部署者；角色必须按每项服务认定。", 42, 164, 1196, 54, { fontSize: 25, bold: true });
    addPanelText(s, "s9-left", 42, 242, 570, 356, "所有 GPAI 提供者", "• 向 AI Office 维护技术文档\n• 向下游系统提供者交付能力 / 限制信息\n• 实施欧盟版权合规政策\n• 发布训练内容摘要\n• 非欧盟提供者指定授权代表\n\n微软证据：Transparency Notes、Phi 发布审查、集中文档流程；已列为 GPAI 行为准则签署方。", "blue");
    addPanelText(s, "s9-right", 668, 242, 570, 356, "系统性风险模型追加义务", "• 通知欧委会并持续模型评估\n• 识别与缓解系统性风险\n• 报告严重事件\n• 确保模型网络安全\n• 提供安全与安全框架 / 模型报告\n\n关键挑战：OpenAI 与其他模型伙伴的信息传递、角色边界、模型版本变化及合同承诺必须一致。", "amber");
    addText(s, "s9-foot", "执法权自 2026.08.02 生效；2025.08.02 前上市的 GPAI 模型可在 2027.08.02 前完成相关过渡义务。", 42, 620, 1170, 34, { fontSize: 19, bold: true, color: C.red });
    notes(s, [
      `European Commission — GPAI provider obligations and enforcement timing: ${URL.euGpaiGuide}`,
      `European Commission — GPAI Code of Practice signatories (includes Microsoft): ${URL.euGpai}`,
      `Microsoft — 2025 Responsible AI Transparency Report (Phi reviews and downstream governance): ${URL.msReport}`,
    ]);
  }

  // 10 — Article 50 transparency chain.
  {
    const s = deck.slides.add();
    addChrome(s, "第 50 条已生效：透明度必须穿透生成、界面与分发", "ARTICLE 50 TRANSPARENCY", 10);
    const xs = [48, 293, 538, 783, 1028];
    for (let i = 0; i < 4; i++) addLine(s, `s10-arrow-${i}`, xs[i] + 178, 322, 67, 0, C.blue, 3, true);
    const chain = [
      ["01", "模型输出", "识别生成 / 操纵内容"],
      ["02", "机器可读标记", "元数据、来源与可检测性"],
      ["03", "界面告知", "明确说明用户正在与 AI 交互"],
      ["04", "场景披露", "深伪、情绪/生物特征、公共利益文本"],
      ["05", "分发与审计", "导出后保留标记并可回溯"],
    ];
    for (let i = 0; i < chain.length; i++) {
      addRect(s, `s10-box-${i}`, xs[i], 250, 178, 148, i === 1 || i === 4 ? C.bluePale : C.panel);
      addText(s, `s10-num-${i}`, chain[i][0], xs[i] + 18, 268, 42, 28, { fontSize: 18, bold: true, color: C.blue });
      addText(s, `s10-title-${i}`, chain[i][1], xs[i] + 18, 306, 142, 36, { fontSize: 24, bold: true, alignment: "center" });
      addText(s, `s10-body-${i}`, chain[i][2], xs[i] + 14, 350, 150, 36, { fontSize: 16, color: C.muted, alignment: "center" });
    }
    addText(s, "s10-left-h", "微软可复用的能力", 42, 446, 540, 38, { fontSize: 30, bold: true });
    addText(s, "s10-left-b", "Transparency Notes、产品内告知、内容凭证实践、平台级安全与元数据能力。", 42, 494, 540, 82, { fontSize: 22 });
    addLine(s, "s10-divider", 620, 438, 0, 166, C.rule, 1);
    addText(s, "s10-right-h", "最容易失效的环节", 666, 446, 572, 38, { fontSize: 30, bold: true, color: C.red });
    addText(s, "s10-right-b", "API 下游改写、截图/转码、第三方模型、旧系统、多语种和多模态会让标签丢失；必须做端到端验证。", 666, 494, 548, 82, { fontSize: 22 });
    addRect(s, "s10-legacy", 42, 614, 1196, 38, C.amberPale);
    addText(s, "s10-legacy-t", "过渡：2026.08.02 前上市的生成式 AI 系统，其机器可读标记义务宽限至 2026.12.02。", 62, 620, 1156, 26, { fontSize: 18, bold: true, color: C.amber, alignment: "center" });
    notes(s, [
      `European Commission — Guidelines on transparency of AI-generated content: ${URL.euTransparency}`,
      `European Commission — Transparency Code Q&A and legacy grace period: ${URL.euTransparencyCode}`,
      `Microsoft — Responsible AI Transparency Report (Transparency Notes and content credentials example): ${URL.msReport}`,
    ]);
  }

  // 11 — Customer tooling.
  {
    const s = deck.slides.add();
    addChrome(s, "微软工具栈可以加速合规，但不能替代客户的法律判断", "CUSTOMER ENABLEMENT", 11);
    addText(s, "s11-lead", "最有效的组合是“分类—预防—测试—监测”形成同一证据流水线，而不是单点采购工具。", 42, 164, 1196, 46, { fontSize: 25, bold: true });
    const lanes = [
      { x: 42, h: "分类与评估", tools: "Purview Compliance Manager\nEU AI Act 模板\nFoundry evaluations\n影响评估模板", use: "把义务转成控制与改进项" },
      { x: 347, h: "预防与约束", tools: "Azure AI Content Safety\nPrompt Shields\nProtected Material\nEntra / Purview DLP", use: "过滤伤害、攻击与敏感数据" },
      { x: 652, h: "测试与挑战", tools: "PyRIT\nAI Red Team\nGroundedness\nTask Adherence", use: "验证失效模式与代理行为" },
      { x: 957, h: "监测与证据", tools: "日志 / 监控\n模型与代理评估\nTransparency Notes\nCompliance Manager", use: "保留版本、结果与改进记录" },
    ];
    for (let i = 0; i < lanes.length; i++) {
      const l = lanes[i];
      addRect(s, `s11-lane-${i}`, l.x, 238, 258, 352, i === 0 ? C.bluePale : C.panel);
      addText(s, `s11-h-${i}`, l.h, l.x + 22, 262, 214, 48, { fontSize: 28, bold: true, color: i === 0 ? C.blue : C.ink });
      addLine(s, `s11-r-${i}`, l.x + 22, 326, 214, 0, C.rule, 1);
      addText(s, `s11-tools-${i}`, l.tools, l.x + 22, 346, 214, 144, { fontSize: 21, bold: true });
      addText(s, `s11-use-${i}`, l.use, l.x + 22, 510, 214, 58, { fontSize: 18, color: C.muted });
    }
    addRect(s, "s11-caveat", 42, 614, 1196, 42, C.redPale);
    addText(s, "s11-caveat-t", "限制：部分功能仍处于预览、语言/区域覆盖不同；客户必须在自己的数据、用户和预期用途上验证。", 62, 622, 1156, 26, { fontSize: 18, bold: true, color: C.red, alignment: "center" });
    notes(s, [
      `Microsoft — EU AI Act implementation tools (Purview, Content Safety, Foundry, PyRIT): ${URL.msActBlog}`,
      `Microsoft Learn — Azure AI Content Safety capabilities and limitations: ${URL.msContentSafety}`,
      `Microsoft Learn — Purview Compliance Manager AI regulation assessments: ${URL.msPurview}`,
    ]);
  }

  // 12 — Shared responsibility table.
  {
    const s = deck.slides.add();
    addChrome(s, "部署模式越开放，客户承担的合规责任越多", "SHARED RESPONSIBILITY", 12);
    addText(s, "s12-sub", "下表是典型责任分配：最终角色仍取决于谁决定预期用途、重大修改、数据与上线方式。", 42, 158, 1160, 34, { fontSize: 20, color: C.muted });
    const values = [
      ["控制任务", "SaaS / Copilot", "PaaS / Foundry", "IaaS / 自建"],
      ["基础模型与云服务安全", "微软主责", "微软主责", "共享"],
      ["应用设计与提示/代理逻辑", "共享", "客户主责", "客户主责"],
      ["业务数据、权限与保留", "客户主责", "客户主责", "客户主责"],
      ["预期用途与风险分类", "客户主责", "客户主责", "客户主责"],
      ["人工监督与员工/个人告知", "客户主责", "客户主责", "客户主责"],
      ["FRIA 与工作场所程序", "客户主责", "客户主责", "客户主责"],
      ["模型文档与下游信息", "微软支持", "共享", "客户主责"],
      ["监控、事件与纠正行动", "共享", "共享", "客户主责"],
    ];
    const table = s.tables.add({ rows: values.length, columns: 4, left: 42, top: 210, width: 1196, height: 414, columnWidths: [340, 285, 285, 286], values });
    styleTable(table, values.length, 4, { headerSize: 19, bodySize: 18, boldFirstColumn: true });
    for (let r = 1; r < values.length; r++) {
      for (let c = 1; c < 4; c++) {
        const v = values[r][c];
        table.getCell(r, c).fill = v.includes("客户") ? C.amberPale : v.includes("微软") ? C.bluePale : C.panel;
        table.getCell(r, c).text.style = { typeface: FONT, fontSize: 18, bold: true, color: v.includes("客户") ? C.amber : v.includes("微软") ? C.blue : C.ink, verticalAlignment: "middle", alignment: "center" };
      }
    }
    addText(s, "s12-foot", "关键提醒：微软提供文档和工具；部署者仍需完成分类、FRIA、人员告知、人工监督与本地监控。", 42, 642, 1170, 30, { fontSize: 19, bold: true, color: C.red });
    notes(s, [
      `Microsoft Learn — AI shared responsibility model: ${URL.msShared}`,
      `Microsoft Learn — Artificial Intelligence assurance overview, customer responsibilities: ${URL.msAssurance}`,
      `European Commission — High-risk deployer obligations and FRIA: ${URL.euNavigate}`,
    ], "Responsibility allocation is a generalized analytical mapping and must be validated against product terms and each deployment context.");
  }

  // 13 — Readiness assessment.
  {
    const s = deck.slides.add();
    addChrome(s, "公开证据显示五项优势、四个优先缺口", "READINESS ASSESSMENT", 13);
    addText(s, "s13-sub", "交通灯仅表示公开控制覆盖度，不代表监管认可或产品合格评定。", 42, 158, 900, 30, { fontSize: 19, color: C.muted });
    const items = [
      [C.green, "治理与问责", "董事会—Council—ORA—工程团队的分层机制清晰", "维持独立挑战与例外升级"],
      [C.green, "禁止实践 / 素养", "政策、筛查、合同与培训已有公司级路径", "验证清册覆盖和培训有效性"],
      [C.green, "风险评估与发布门禁", "发布前审查、敏感用途、红队与工作流可规模化", "把版本变更纳入再评估"],
      [C.amber, "GPAI 供应链", "已签 GPAI 准则；公开文档实践成熟", "固化伙伴模型的信息与事件 RACI"],
      [C.amber, "第 50 条透明度", "具备 Transparency Notes 与内容凭证实践", "保证导出、转码、API 下游仍保留标记"],
      [C.amber, "高风险合格评定", "控制组件存在；统一标准与最终证据仍在演进", "2027 前完成 QMS 和评定干运行"],
    ];
    addText(s, "s13-h1", "状态", 42, 208, 90, 28, { fontSize: 18, bold: true, color: C.muted });
    addText(s, "s13-h2", "控制域与公开证据", 132, 208, 700, 28, { fontSize: 18, bold: true, color: C.muted });
    addText(s, "s13-h3", "管理优先项", 872, 208, 330, 28, { fontSize: 18, bold: true, color: C.muted });
    for (let i = 0; i < items.length; i++) {
      const y = 246 + i * 67;
      addLine(s, `s13-line-${i}`, 42, y + 58, 1196, 0, C.rule, 1);
      addCircle(s, `s13-dot-${i}`, 76, y + 25, 22, items[i][0]);
      addText(s, `s13-domain-${i}`, items[i][1], 132, y, 250, 34, { fontSize: 23, bold: true });
      addText(s, `s13-evidence-${i}`, items[i][2], 392, y, 440, 44, { fontSize: 18 });
      addText(s, `s13-priority-${i}`, items[i][3], 872, y, 330, 44, { fontSize: 18, bold: true, color: items[i][0] === C.amber ? C.amber : C.muted });
    }
    addRect(s, "s13-bottom", 42, 654, 1196, 1, C.rule);
    notes(s, [
      `Microsoft — 2025 Responsible AI Transparency Report: ${URL.msReport}`,
      `Microsoft Trust Center — EU AI Act approach: ${URL.msTrust}`,
      `European Commission — AI Act implementation timeline and obligations: ${URL.euNavigate}`,
      `European Commission — Article 50 transparency guidelines: ${URL.euTransparency}`,
    ], "Traffic-light ratings are analytical judgments derived from public evidence, not audit or certification results.");
  }

  // 14 — Roadmap, adapted from Codex Grid Gantt layout.
  {
    const s = deck.slides.add();
    addChrome(s, "建议路线图：先守住已生效义务，再为高风险评定做干运行", "EXECUTION ROADMAP", 14);
    const left = 270;
    const top = 190;
    const colW = 184;
    const headers = ["立即\n0–30 天", "短期\n31–90 天", "2026 Q4", "2027 H1", "2027 H2" ];
    for (let i = 0; i < headers.length; i++) {
      addRect(s, `s14-head-${i}`, left + i * colW, top, colW, 62, i < 2 ? C.bluePale : C.panel);
      addText(s, `s14-head-t-${i}`, headers[i], left + i * colW + 8, top + 10, colW - 16, 44, { fontSize: 18, bold: true, color: i < 2 ? C.blue : C.ink, alignment: "center", verticalAlignment: "middle" });
      addLine(s, `s14-grid-${i}`, left + i * colW, top, 0, 420, C.rule, 1);
    }
    addLine(s, "s14-grid-end", left + 5 * colW, top, 0, 420, C.rule, 1);
    const tracks = [
      { label: "清册与角色分类", start: 0, span: 2, color: C.blue, text: "统一 AI 清册 / 角色 / 风险" },
      { label: "已生效义务", start: 0, span: 3, color: C.red, text: "禁止实践、素养、第 50 条、GPAI" },
      { label: "证据与合同", start: 1, span: 3, color: C.blue2, text: "证据目录、下游文档、供应商 RACI" },
      { label: "客户部署控制", start: 1, span: 4, color: C.amber, text: "FRIA、人工监督、通知、事件流程" },
      { label: "高风险干运行", start: 2, span: 3, color: C.green, text: "QMS、合格评定、登记、上市后监测" },
    ];
    for (let i = 0; i < tracks.length; i++) {
      const y = 276 + i * 70;
      addText(s, `s14-label-${i}`, tracks[i].label, 42, y + 10, 205, 34, { fontSize: 21, bold: true });
      addLine(s, `s14-row-${i}`, left, y + 52, 5 * colW, 0, C.rule, 1);
      const x = left + tracks[i].start * colW + 10;
      const w = tracks[i].span * colW - 20;
      addRect(s, `s14-bar-${i}`, x, y, w, 46, tracks[i].color, "none", 0, true);
      addText(s, `s14-bar-t-${i}`, tracks[i].text, x + 14, y + 9, w - 28, 28, { fontSize: 17, bold: true, color: C.white, alignment: "center", verticalAlignment: "middle" });
    }
    addRect(s, "s14-gate", 42, 642, 1196, 32, C.ink);
    addText(s, "s14-gate-t", "管理门禁：任何进入欧盟市场的 AI 版本，没有角色、风险、证据负责人和上线后监测计划，不得发布。", 58, 647, 1164, 22, { fontSize: 17, bold: true, color: C.white, alignment: "center" });
    notes(s, [
      `European Commission — AI Act dates and high-risk preparation window: ${URL.euNavigate}`,
      `Microsoft — EU AI Act implementation approach: ${URL.msActBlog}`,
      `Microsoft — Responsible AI Transparency Report (central workflow and lifecycle controls): ${URL.msReport}`,
    ], "Roadmap and management gate are recommendations developed for this analysis.");
  }

  // 15 — Close with decisions, adapted from stacked-text-flow close.
  {
    const s = deck.slides.add();
    s.background.fill = C.white;
    addText(s, "s15-eyebrow", "BOARD / EXECUTIVE DECISIONS", 42, 42, 420, 32, { fontSize: 20, bold: true, color: C.blue });
    addText(s, "s15-title", "管理层需立即拍板的三件事", 42, 126, 1100, 80, { fontSize: 64, bold: true, autoFit: "none" });
    const decisions = [
      ["01", "建立唯一的欧盟 AI 控制面", "统一清册、角色分类、风险评级、证据目录与例外升级。"],
      ["02", "把合规变成发布门禁", "禁止实践、第 50 条、GPAI 与客户责任在每个版本发布前完成签核。"],
      ["03", "用共享责任合同化、工程化", "让微软、模型伙伴、集成商和客户对文档、监控、事件与纠正行动各自负责。"],
    ];
    for (let i = 0; i < decisions.length; i++) {
      const y = 258 + i * 118;
      addLine(s, `s15-line-${i}`, 42, y - 18, 1196, 0, C.rule, 1);
      addText(s, `s15-num-${i}`, decisions[i][0], 42, y, 70, 44, { fontSize: 26, bold: true, color: C.blue });
      addText(s, `s15-head-${i}`, decisions[i][1], 132, y, 430, 42, { fontSize: 30, bold: true });
      addText(s, `s15-body-${i}`, decisions[i][2], 590, y, 620, 56, { fontSize: 21 });
    }
    addRect(s, "s15-close", 42, 624, 1196, 48, C.blue);
    addText(s, "s15-close-t", "战略优势不是“有一套原则”，而是让每次 AI 决策都留下可复核、可追踪、可纠正的证据。", 62, 634, 1156, 28, { fontSize: 21, bold: true, color: C.white, alignment: "center" });
    notes(s, [
      `Microsoft — Responsible AI governance architecture and practices: ${URL.msAssurance}`,
      `European Commission — AI Act roles, enforcement and obligations: ${URL.euNavigate}`,
    ], "The three decisions are recommendations developed for this analysis.");
  }

  for (let i = 0; i < deck.slides.items.length; i++) {
    const slide = deck.slides.items[i];
    const stem = `slide-${String(i + 1).padStart(2, "0")}`;
    const png = await deck.export({ slide, format: "png", scale: 1 });
    await writeBlob(`${PREVIEW_DIR}/${stem}.png`, png);
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${PREVIEW_DIR}/${stem}.layout.json`, await layout.text());
  }
  const montage = await deck.export({ format: "webp", montage: true, scale: 1 });
  await writeBlob(`${PREVIEW_DIR}/deck-montage.webp`, montage);
  const pptx = await PresentationFile.exportPptx(deck);
  await pptx.save(OUT);
  console.log(`Created ${OUT}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
