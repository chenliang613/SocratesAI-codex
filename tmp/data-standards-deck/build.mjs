import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "/Users/apple/SocratesAI-codex/outputs/数据国际标准分析.pptx";
const TMP = "/Users/apple/SocratesAI-codex/tmp/data-standards-deck";
const W = 1280, H = 720;
const C = {
  ink: "#111827", muted: "#5B6472", rule: "#C9CED6", panel: "#F1F3F5",
  blue: "#2E7DFF", cyan: "#7DD3FC", darkBlue: "#164E93", green: "#1C9B75",
  amber: "#D99000", red: "#B54747", white: "#FFFFFF"
};

function box(slide, left, top, width, height, fill = "none", lineFill = "none", radius = "square") {
  return slide.shapes.add({ geometry: radius === "round" ? "roundRect" : "rect", position: { left, top, width, height }, fill, line: { style: "solid", fill: lineFill, width: lineFill === "none" ? 0 : 1 } });
}
function text(slide, value, left, top, width, height, style = {}) {
  const s = slide.shapes.add({ geometry: "textbox", position: { left, top, width, height }, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  s.text = value;
  s.text.style = { typeface: "Arial", fontSize: 18, color: C.ink, alignment: "left", verticalAlignment: "top", autoFit: "shrinkText", insets: { top: 0, right: 0, bottom: 0, left: 0 }, ...style };
  return s;
}
function rule(slide, left, top, width, color = C.rule, height = 2) { box(slide, left, top, width, height, color); }
function title(slide, value, kicker = "国际标准分析") {
  text(slide, kicker.toUpperCase(), 42, 34, 460, 24, { fontSize: 14, bold: true, color: C.blue, letterSpacing: 1 });
  text(slide, value, 42, 66, 1160, 62, { fontSize: 36, bold: true, color: C.ink });
  rule(slide, 42, 144, 1196, C.rule, 1);
}
function footer(slide, n) { text(slide, `数据国际标准分析  ·  ${String(n).padStart(2, "0")}`, 42, 674, 500, 20, { fontSize: 11, color: C.muted }); }
function note(slide, urls, extra = "") { slide.speakerNotes.textFrame.setText(`[Sources]\n${urls.map(u => `- ${u}`).join("\n")}${extra ? `\n\n${extra}` : ""}`); slide.speakerNotes.setVisible(true); }
function bullet(slide, items, left, top, width, lineH = 31, fs = 19, color = C.ink) {
  items.forEach((v, i) => { text(slide, `•  ${v}`, left, top + i * lineH, width, lineH + 6, { fontSize: fs, color }); });
}
function label(slide, value, left, top, width, color = C.blue) { text(slide, value.toUpperCase(), left, top, width, 22, { fontSize: 13, bold: true, color, letterSpacing: 0.8 }); }
function callout(slide, value, left, top, width, height, fill = C.panel, accent = C.blue, fs = 22) {
  box(slide, left, top, width, height, fill);
  box(slide, left, top, 7, height, accent);
  text(slide, value, left + 24, top + 18, width - 42, height - 30, { fontSize: fs, bold: true });
}
function table(slide, rows, left, top, widths, rowH = 44, header = true, fs = 15) {
  let y = top;
  rows.forEach((row, ri) => {
    let x = left;
    const fill = ri === 0 && header ? C.darkBlue : (ri % 2 ? C.panel : C.white);
    row.forEach((cell, ci) => {
      box(slide, x, y, widths[ci], rowH, fill, C.white);
      text(slide, String(cell), x + 12, y + 10, widths[ci] - 22, rowH - 15, { fontSize: ri === 0 && header ? fs - 1 : fs, bold: ri === 0 && header, color: ri === 0 && header ? C.white : C.ink, verticalAlignment: "middle" });
      x += widths[ci];
    });
    y += rowH;
  });
}
function addBar(slide, value, max, left, top, width, color, labelText) {
  text(slide, labelText, left, top - 2, 190, 24, { fontSize: 16, bold: true });
  box(slide, left + 190, top + 3, width, 18, C.panel);
  box(slide, left + 190, top + 3, width * value / max, 18, color);
  text(slide, String(value), left + 190 + width + 12, top - 3, 48, 24, { fontSize: 16, bold: true, color });
}

const p = Presentation.create({ slideSize: { width: W, height: H } });

// 1 Cover
{ const s = p.slides.add(); s.background.fill = C.white; rule(s, 42, 42, 170, C.blue, 8); text(s, "SC32 × SC38", 42, 88, 600, 40, { fontSize: 22, bold: true, color: C.blue }); text(s, "数据国际标准分析", 42, 160, 900, 120, { fontSize: 66, bold: true }); text(s, "从数据定义、语义注册到云上流动与共享的标准体系", 46, 310, 920, 48, { fontSize: 25, color: C.muted }); box(s, 42, 438, 550, 122, C.panel); text(s, "ISO/IEC JTC 1\nSC32 · Data management and interchange\nSC38 · Cloud computing and distributed platforms", 70, 464, 480, 82, { fontSize: 19, bold: true, color: C.ink }); text(s, "资料基准：ISO 官方委员会范围与目录 · 2026-08-07", 42, 660, 560, 22, { fontSize: 13, color: C.muted }); note(s, ["https://www.iso.org/committee/45342.html", "https://www.iso.org/committee/601355.html"]); }

// 2 Scope
{ const s = p.slides.add(); title(s, "两个委员会分工不同，但共同构成数据基础设施"); text(s, "SC32 规定数据如何被定义、组织、交换与解释；SC38 规定数据如何在云与分布式平台中流动、共享与被治理。", 42, 172, 1160, 42, { fontSize: 22, bold: true }); box(s, 42, 250, 540, 300, C.panel); label(s, "SC32 · Data management and interchange", 72, 280, 420, C.darkBlue); bullet(s, ["参考模型与框架", "数据域、类型、结构及语义", "持久存储、并发访问、数据交换", "元数据注册与互操作"], 72, 330, 430, 40, 20); box(s, 660, 250, 540, 300, "#EAF4FF"); label(s, "SC38 · Cloud & distributed platforms", 690, 280, 430, C.blue); bullet(s, ["云计算与分布式平台基础概念", "云间 / 边缘 / 多云的交互", "数据流、数据类别、数据使用", "共享协议、可移植性与数字主权"], 690, 330, 430, 40, 20); footer(s, 2); note(s, ["https://www.iso.org/committee/45342.html", "https://www.iso.org/committee/601355.html"]); }

// 3 status
{ const s = p.slides.add(); title(s, "阅读标准目录，先看状态，再看编号"); text(s, "ISO 官方目录同时展示已发布、制定中、撤销与删除项目；对于工程选型，状态决定可引用性与落地风险。", 42, 172, 1100, 40, { fontSize: 21 }); const rows = [["状态", "工程含义", "建议动作"], ["Published", "可作为现行依据；仍需确认版本与修订", "纳入企业标准基线"], ["Under development", "方向性信号，内容与时间仍可能变化", "跟踪、试点，不作硬性合规依据"], ["Withdrawn", "历史版本已被撤销或替代", "仅用于兼容性与演进分析"]]; table(s, rows, 42, 245, [220, 510, 466], 70, true, 17); callout(s, "结论：标准体系是“版本化资产”，不能只维护一个编号清单。", 42, 570, 1158, 68, "#EEF6FF", C.blue, 22); footer(s, 3); note(s, ["https://www.iso.org/committee/45342/x/catalogue/", "https://www.iso.org/committee/601355/x/catalogue/"]); }

// 4 SC32 map
{ const s = p.slides.add(); title(s, "SC32 的数据标准形成一条从语义到执行的链"); text(s, "SC32 的价值不只是数据库语言；它把“数据是什么、如何命名、如何注册、如何交换、如何分析”连成一套可复用的基础层。", 42, 172, 1160, 44, { fontSize: 21 }); const xs = [42, 282, 522, 762, 1002]; const labs = [["01", "语义与概念", "2382 · 5394 · 21838"], ["02", "类型与标识", "11404 · 5218 · 6523"], ["03", "元数据注册", "11179 · 19763 · 20944"], ["04", "交换与数据库", "9075 · 13249 · 14957"], ["05", "使用与分析", "5207 · 5212 · 29075 · 39075"]]; xs.forEach((x, i) => { box(s, x, 285, 200, 190, i === 2 ? "#EAF4FF" : C.panel); text(s, labs[i][0], x + 20, 305, 50, 28, { fontSize: 18, bold: true, color: i === 2 ? C.blue : C.muted }); text(s, labs[i][1], x + 20, 350, 165, 45, { fontSize: 22, bold: true }); text(s, labs[i][2], x + 20, 422, 165, 35, { fontSize: 15, color: C.muted }); if (i < 4) text(s, "→", x + 208, 360, 34, 40, { fontSize: 28, bold: true, color: C.blue, alignment: "center" }); }); callout(s, "SC32 的核心产出是“可被不同系统共同理解的对象”：数据元素、概念系统、数据集、模型与查询语言。", 42, 540, 1158, 78, "#F6FBFF", C.blue, 20); footer(s, 4); note(s, ["https://www.iso.org/committee/45342.html", "https://www.iso.org/committee/45342/x/catalogue/"]); }

// 5 Metadata
{ const s = p.slides.add(); title(s, "SC32 的元数据主线已扩展到数据集、模型与可计算数据"); text(s, "ISO/IEC 11179:2023 及其配套系列，把注册对象从数据元素扩展到概念系统、数据集、可计算数据与模型。", 42, 172, 1160, 44, { fontSize: 21 }); const rows = [["系列", "当前重点", "作用"], ["ISO/IEC 11179-1/-2/-3", "框架、分类、注册公共设施", "定义元数据注册的共同语义"], ["ISO/IEC 11179-30/-31", "元数据 / 数据规格注册", "让数据对象可发现、可描述、可复用"], ["ISO/IEC 11179-32/-33", "概念系统 / 数据集注册", "把词表、代码表与数据产品连接"], ["ISO/IEC 11179-34/-35", "可计算数据 / 模型注册", "支撑算法、模型与机器可读资产"]]; table(s, rows, 42, 248, [240, 420, 498], 67, true, 16); text(s, "配套：ISO/IEC 19763 MFI（互操作元模型）、ISO/IEC 20944 MDR-IB（API / 协议 / 绑定）、ISO/IEC 19583（元数据使用与映射）。", 42, 604, 1150, 34, { fontSize: 17, color: C.muted }); footer(s, 5); note(s, ["https://www.iso.org/committee/45342/x/catalogue/", "https://www.iso.org/committee/45342/x/catalogue/"], "关键目录位置：ISO/IEC 11179-1:2023、11179-3:2023、11179-30:2023、11179-31:2023、11179-32:2023、11179-33:2023、11179-34:2024、11179-35:2023。"); }

// 6 DB
{ const s = p.slides.add(); title(s, "数据库语言仍是 SC32 的主干，但对象正在从关系表扩展到图"); text(s, "SQL 2023 系列继续提供关系数据库的标准语言基础；GQL 2024 则为属性图查询建立独立的国际标准入口。", 42, 172, 1150, 42, { fontSize: 21 }); const rows = [["技术族", "现行代表", "标准化关注点"], ["SQL / ISO/IEC 9075", "Part 1, 2, 3, 4, 9, 14（2023）", "框架、基础语法、CLI、存储模块、外部数据、SQL/XML"], ["GQL / ISO/IEC 39075", "GQL（2024）", "图模式、图构造、图查询与图数据处理"], ["SQL packages / 13249", "空间、全文、图像、数据挖掘等历史包", "数据库内置能力的领域化扩展"], ["数据表示 / 11404、14957", "通用数据类型、数据值格式", "跨语言、跨系统的数据类型与格式一致性"]]; table(s, rows, 42, 245, [240, 390, 528], 62, true, 16); callout(s, "判断：关系 + 图并存，企业数据平台需要同时管理表语义、图语义与跨模型映射。", 42, 570, 1158, 72, "#F6FBFF", C.blue, 21); footer(s, 6); note(s, ["https://www.iso.org/committee/45342/x/catalogue/", "https://www.iso.org/standard/80327.html"]); }

// 7 usage
{ const s = p.slides.add(); title(s, "SC32 的新方向从“如何存”转向“如何负责任地使用数据"); text(s, "2024 年的数据使用系列与制定中的可信使用、组织评估和高级分析函数，显示标准重心正在上移到数据使用治理。", 42, 172, 1160, 43, { fontSize: 21 }); addBar(s, 2, 4, 80, 265, 480, C.blue, "Published"); addBar(s, 2, 4, 80, 325, 480, C.amber, "Under dev."); text(s, "数据使用主线", 80, 405, 200, 28, { fontSize: 18, bold: true }); bullet(s, ["ISO/IEC 5207: 术语与用例", "ISO/IEC 5212: 数据使用指导", "ISO/IEC CD 24927: 组织中的数据使用评估", "ISO/IEC CD 25985: 可信数据使用", "ISO/IEC CD 29075-1/-2: 高级分析函数库"], 80, 445, 600, 34, 17); box(s, 820, 260, 360, 280, "#EAF4FF"); text(s, "对企业的含义", 850, 292, 280, 32, { fontSize: 22, bold: true, color: C.darkBlue }); text(s, "数据治理不应只停留在“有没有目录”。\n\n还要回答：\n• 谁可以用\n• 用于什么目的\n• 是否可审计\n• 结果如何被复现", 850, 350, 280, 160, { fontSize: 18 }); footer(s, 7); note(s, ["https://www.iso.org/committee/45342/x/catalogue/"]); }

// 8 SC38 foundation
{ const s = p.slides.add(); title(s, "SC38 先建立云与分布式平台的共同语言"); text(s, "SC38 的基础层已经从“云是什么”演进到多云、平台分类、平台能力与数字主权。", 42, 172, 1160, 42, { fontSize: 21 }); const rows = [["层次", "代表标准", "分析"], ["词汇 / 概念 / 架构", "ISO/IEC 22123-1/-2/-3:2023", "当前基础三件套；替代撤销的 17788 / 17789"], ["数字平台分类", "ISO/IEC TS 5928:2023", "统一平台一词在云与分布式系统中的含义"], ["平台能力", "ISO/IEC TS 7339:2024", "描述平台能力类型与 PaaS"], ["多云 / 主权", "ISO/IEC 5140:2024、TS 10866:2024", "把跨云使用与组织自主性纳入标准语言"]]; table(s, rows, 42, 245, [230, 420, 508], 60, true, 16); callout(s, "SC38 的变化：从基础设施标准，转向平台生态与组织控制权标准。", 42, 570, 1158, 70, "#F6FBFF", C.blue, 21); footer(s, 8); note(s, ["https://www.iso.org/committee/601355.html", "https://www.iso.org/committee/601355/x/catalogue/", "https://www.iso.org/cms/live/live/en/sites/isoorg/contents/data/standard/08/18/81848.html"]); }

// 9 data flow
{ const s = p.slides.add(); title(s, "SC38 的数据主线：从数据流与分类，到共享协议与处理策略"); text(s, "SC38 直接触及数据跨设备、云服务与分布式平台的流动；其标准更像“数据运行环境”的治理层。", 42, 172, 1160, 42, { fontSize: 21 }); const rows = [["标准", "解决的问题", "落地关注"], ["ISO/IEC 19944-1/-2", "数据流、数据类别、数据使用", "识别数据在设备—云—服务间如何移动"], ["ISO/IEC 22624", "基于分类的数据处理", "按数据类别设计云服务处理规则"], ["ISO/IEC 23751", "数据共享协议（DSA）框架", "把参与方、目的、责任、约束写进协议"], ["ISO/IEC 19941", "云互操作与可移植性", "降低跨云迁移与切换障碍"]]; table(s, rows, 42, 245, [250, 430, 478], 60, true, 16); box(s, 42, 570, 1158, 48, "#EAF4FF"); text(s, "关键判断：SC32 负责“数据对象可理解”，SC38 负责“数据在平台中可被安全、可控地使用”。", 64, 593, 1110, 28, { fontSize: 19, bold: true, color: C.darkBlue }); footer(s, 9); note(s, ["https://www.iso.org/standard/79573.html", "https://www.iso.org/committee/601355/x/catalogue/"]); }

// 10 emerging
{ const s = p.slides.add(); title(s, "SC38 的下一波重点是数据空间、多云管理、边缘与 AI 服务"); text(s, "目录中的制定中项目，集中在跨组织数据空间与分布式运行的“协作控制面”。", 42, 172, 1120, 40, { fontSize: 21 }); const items = [["数据空间", "20151-1 FDIS：概念与特征；20151-2 AWI：信任框架", C.blue], ["多云管理", "10822-1 已发布；-2 身份、-3 编排、-4 监控制定中", C.darkBlue], ["边缘网络", "19274 DIS：云、边缘、多云与联邦云的网络基础", C.green], ["数字主权", "10866：组织自主性与数字主权框架", C.amber], ["AI 服务", "26191 AWI：云计算对 AI 服务的支持", C.red]]; items.forEach((it, i) => { const y = 250 + i * 74; box(s, 42, y, 12, 54, it[2]); text(s, it[0], 76, y + 4, 190, 28, { fontSize: 20, bold: true }); text(s, it[1], 280, y + 4, 880, 42, { fontSize: 17, color: C.ink }); }); callout(s, "趋势：标准化对象从单一云服务，升级为跨组织、跨平台的数据协作网络。", 42, 625, 1158, 44, "#F6FBFF", C.blue, 19); footer(s, 10); note(s, ["https://www.iso.org/committee/601355/x/catalogue/", "https://www.iso.org/cms/live/live/en/sites/isoorg/contents/data/standard/08/65/86589.html", "https://www.iso.org/cms/live/live/es/sites/isoorg/contents/data/standard/08/57/85772.html", "https://www.iso.org/cms/live/live/en/sites/isoorg/contents/data/standard/09/42/94290.html"]); }

// 11 crosswalk
{ const s = p.slides.add(); title(s, "把 SC32 与 SC38 放进一条企业数据架构链"); text(s, "企业不应按委员会采购标准，而应按数据生命周期与控制点组合标准。", 42, 172, 1100, 40, { fontSize: 21 }); const cols = [{x:42, name:"定义", code:"SC32", body:"概念、数据元素、类型、标识", color:C.darkBlue}, {x:282, name:"注册", code:"SC32", body:"元数据、数据集、模型、词表", color:C.blue}, {x:522, name:"交换", code:"SC32", body:"SQL、GQL、绑定、格式", color:C.blue}, {x:762, name:"运行", code:"SC38", body:"云、边缘、多云、可移植性", color:C.green}, {x:1002, name:"共享", code:"SC38", body:"数据流、DSA、数据空间、主权", color:C.amber}]; cols.forEach((c, i) => { box(s, c.x, 280, 200, 250, i === 2 ? "#EAF4FF" : C.panel); text(s, c.code, c.x + 18, 302, 160, 25, { fontSize: 14, bold: true, color: c.color }); text(s, c.name, c.x + 18, 350, 160, 38, { fontSize: 26, bold: true }); text(s, c.body, c.x + 18, 420, 160, 74, { fontSize: 18, color: C.muted }); if (i < cols.length - 1) text(s, "→", c.x + 206, 370, 30, 40, { fontSize: 26, bold: true, color: C.blue }); }); callout(s, "管理层看的是“可控数据流”，技术团队落的是“可复用标准对象”。", 42, 575, 1158, 70, "#F6FBFF", C.blue, 22); footer(s, 11); note(s, ["https://www.iso.org/committee/45342.html", "https://www.iso.org/committee/601355.html", "https://www.iso.org/committee/45342/x/catalogue/", "https://www.iso.org/committee/601355/x/catalogue/"]); }

// 12 implications
{ const s = p.slides.add(); title(s, "对企业数据平台的四个直接影响"); const rows = [["影响", "标准抓手", "优先动作"], ["数据产品可发现", "11179-30/-31/-33/-34/-35", "建立数据集、模型与可计算数据注册"], ["跨系统语义一致", "5394、11179、19763、21838", "统一概念系统、命名、代码表与映射"], ["多云数据可控", "19944、22624、23751、19941", "把数据类别、用途、共享责任写入协议"], ["分析与 AI 可复现", "5207、5212、24927、25985、29075、39075", "记录使用目的、函数、图 / 表查询与审计轨迹"]]; table(s, rows, 42, 222, [260, 440, 458], 67, true, 16); box(s, 42, 590, 1158, 45, C.ink); text(s, "优先级原则：先把“对象”和“责任”标准化，再扩展平台和算法能力。", 65, 601, 1100, 26, { fontSize: 19, bold: true, color: C.white }); footer(s, 12); note(s, ["https://www.iso.org/committee/45342/x/catalogue/", "https://www.iso.org/committee/601355/x/catalogue/"]); }

// 13 roadmap
{ const s = p.slides.add(); title(s, "建议的落地路线：先建语义底座，再接云上数据流"); const steps = [["0–3个月", "盘点与版本基线", "按 Published / Under development / Withdrawn 清理标准清单"], ["3–6个月", "元数据与概念系统", "落地 11179、概念系统、数据集与模型注册"], ["6–12个月", "数据交换与平台接入", "用 SQL / GQL / API 绑定与云数据流标准建立接口契约"], ["12个月以后", "跨组织协作与可信使用", "试点 DSA、数据空间、多云切换与数据使用审计"]]; steps.forEach((st, i) => { const y = 220 + i * 94; box(s, 42, y, 160, 64, i === 3 ? C.darkBlue : C.blue); text(s, st[0], 58, y + 18, 125, 28, { fontSize: 18, bold: true, color: C.white }); text(s, st[1], 246, y + 4, 280, 28, { fontSize: 22, bold: true }); text(s, st[2], 246, y + 37, 870, 25, { fontSize: 17, color: C.muted }); if (i < 3) rule(s, 118, y + 64, 2, C.rule, 30); }); callout(s, "不要把“标准采用”理解为一次性认证；它更像数据平台的持续版本治理。", 42, 610, 1158, 52, "#F6FBFF", C.blue, 20); footer(s, 13); note(s, ["https://www.iso.org/committee/45342/x/catalogue/", "https://www.iso.org/committee/601355/x/catalogue/"]); }

// 14 inventory
{ const s = p.slides.add(); title(s, "附录：SC32 数据相关标准按系列归并的当前清单"); text(s, "口径：覆盖 ISO 官方 SC32 目录中仍有现行版本或制定中项目的主要数据主线；同系列的撤销历史版本不重复展开。", 42, 172, 1160, 38, { fontSize: 18, color: C.muted }); const rows = [["主题", "系列 / 代表编号", "状态特征"], ["语义 / 标识", "2382-4；5207；5212；5218；5394；6523-1/-2", "现行 + 2024/2025 更新"], ["元数据注册", "11179-1/-2/-3/-4/-5/-6；11179-30/-31/-32/-33/-34/-35", "2023–2024 主线；部分制定中"], ["互操作 / 元模型", "19763；19583；20944；21838；24707", "现行 + 持续扩展"], ["数据库 / 查询", "9075；13249；11404；39075", "SQL 2023、GQL 2024；持续修订"], ["数据交换 / 业务", "13238-3；14957；14662；15944", "现行与历史兼容并存"], ["数据使用 / 分析", "24927；25985；29075-1/-2", "制定中为主；5207/5212 已发布"]]; table(s, rows, 42, 238, [250, 610, 298], 54, true, 14); text(s, "完整目录仍应以 ISO 页面实时查询为准；编号、阶段与撤销状态可能变化。", 42, 636, 1100, 26, { fontSize: 15, color: C.red }); footer(s, 14); note(s, ["https://www.iso.org/committee/45342/x/catalogue/"]); }

// 15 inventory SC38 + sources
{ const s = p.slides.add(); title(s, "附录：SC38 数据相关标准与主要来源"); const rows = [["主题", "代表编号", "状态 / 作用"], ["基础架构", "22123-1/-2/-3；5928；7339", "现行基础词汇、概念、架构与平台分类"], ["数据流与处理", "19944-1/-2；22624；23751", "现行数据流、分类、处理与共享协议"], ["互操作 / 多云", "19941；5140；10822 系列", "现行 + 制定中"], ["数据空间 / 主权", "20151 系列；10866；19274", "FDIS / AWI / DIS 等制定中"], ["生态与服务", "19086；23187；23613；26191", "SLA、云服务伙伴、计量与 AI 支持"]]; table(s, rows, 42, 220, [250, 500, 408], 58, true, 15); label(s, "主要来源", 42, 625, 200, C.blue); text(s, "ISO SC32 委员会范围与目录 · ISO SC38 委员会范围与目录 · ISO/IEC 19944-1 标准页 · ISO/IEC TS 5928 标准页", 42, 652, 1160, 24, { fontSize: 13, color: C.muted }); footer(s, 15); note(s, ["https://www.iso.org/committee/45342.html", "https://www.iso.org/committee/45342/x/catalogue/", "https://www.iso.org/committee/601355.html", "https://www.iso.org/committee/601355/x/catalogue/", "https://www.iso.org/standard/79573.html", "https://www.iso.org/cms/live/live/en/sites/isoorg/contents/data/standard/08/18/81848.html"]); }

await fs.mkdir("/Users/apple/SocratesAI-codex/outputs", { recursive: true });
for (const [i, slide] of p.slides.items.entries()) {
  const stem = `slide-${String(i + 1).padStart(2, "0")}`;
  await fs.writeFile(`${TMP}/${stem}.layout.json`, await (await slide.export({ format: "layout" })).text());
  const png = await p.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(`${TMP}/${stem}.png`, new Uint8Array(await png.arrayBuffer()));
}
const montage = await p.export({ format: "webp", montage: true, scale: 1 });
await fs.writeFile(`${TMP}/deck-montage.webp`, new Uint8Array(await montage.arrayBuffer()));
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(OUT);
console.log(OUT);
