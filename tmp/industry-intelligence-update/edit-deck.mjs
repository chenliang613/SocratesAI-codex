import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const ROOT = "/Users/apple/SocratesAI-codex";
const TMP = `${ROOT}/tmp/industry-intelligence-update`;
const STARTER = `${TMP}/template-starter.pptx`;
const OUT = `${ROOT}/行业智能化分析.pptx`;
const RENDER = `${TMP}/final-render`;

const deck = await PresentationFile.importPptx(await FileBlob.load(STARTER));

function setText(slide, name, value) {
  const hit = slide.shapes.items.find((shape) => shape.name === name);
  if (!hit) throw new Error(`Missing shape ${name}`);
  hit.text = value;
  return hit;
}

function setNotes(slide, urls, context="") {
  const block = ["[Sources]", ...urls.map(u => `- ${u}`), "[/Sources]", context].filter(Boolean).join("\n");
  slide.speakerNotes.textFrame.setText(block);
}

// Slide 22 — evidence section divider.
{
  const s = deck.slides.getItem(21);
  setText(s, "text-42-38", "行业证据更新 · 2026");
  setText(s, "text-42-176", "补充证据\n与行业观察");
  setText(s, "text-42-500", "用最新公开数据检验六层架构，并补足标准、政策与落地指标");
  setText(s, "text-1080-466", "量化证据\n开放标准\n监管要求");
  setNotes(s, [
    "https://hai.stanford.edu/ai-index/2026-ai-index-report",
    "https://www.iea.org/reports/key-questions-on-energy-and-ai/executive-summary"
  ]);
}

// Slide 23 — adoption and risk evidence.
{
  const s = deck.slides.getItem(22);
  setText(s, "text-42-24", "证据与行业观察");
  setText(s, "text-42-55", "AI 已进入广泛采用期，但 Agent 仍处于早期规模化");
  setText(s, "text-1178-674", "23");
  setText(s, "text-42-158", "企业 AI 的主要矛盾已从“是否采用”转向“能否可靠地嵌入流程并兑现结果”。");
  setText(s, "text-66-320", "01");
  setText(s, "text-66-380", "88% 已采用");
  setText(s, "text-66-446", "2025年，88%的受访组织已采用 AI；70%已在至少一个职能使用生成式 AI。");
  setText(s, "text-476-320", "02");
  setText(s, "text-476-380", "Agent 仍是个位数");
  setText(s, "text-476-446", "Agent 在绝大多数业务职能中的部署率仍为个位数，生产化远落后于试用。");
  setText(s, "text-886-320", "03");
  setText(s, "text-886-380", "362 起事故");
  setText(s, "text-886-446", "2025年记录的 AI 事故由2024年的233起升至362起，治理压力同步上升。");
  setNotes(s, [
    "https://hai.stanford.edu/ai-index/2026-ai-index-report/economy",
    "https://hai.stanford.edu/ai-index/2026-ai-index-report/responsible-ai"
  ], "口径：组织采用率来自 Stanford 2026 AI Index 汇总的调查；事故数来自 AI Incident Database，经 Stanford 2026 AI Index 引用。");
}

// Slide 24 — energy evidence.
{
  const s = deck.slides.getItem(23);
  setText(s, "text-42-24", "证据与行业观察");
  setText(s, "text-42-55", "能源将成为 AI 扩张速度的第一约束");
  setText(s, "text-1178-674", "24");
  setText(s, "text-42-154", "效率持续提升，但需求规模、Agent 使用与基础设施投资推动总用电继续上升。");
  setText(s, "text-42-294", "950 TWh");
  setText(s, "text-42-350", "IEA预计数据中心用电：485 TWh（2025）→约950 TWh（2030），接近翻倍。");
  setText(s, "text-442-294", "3×");
  setText(s, "text-442-350", "AI专用数据中心2030年用电约为2025年的3倍。");
  setText(s, "text-842-294", "17%");
  setText(s, "text-842-350", "2025年数据中心用电同比增长17%，显著高于全球总用电3%的增速。");
  setText(s, "text-66-569", "管理含义");
  setText(s, "text-210-563", "电力可获得性、并网周期、能源设备/芯片供应链与融资能力，将共同决定算力落地节奏。");
  setNotes(s, [
    "https://www.iea.org/reports/key-questions-on-energy-and-ai/executive-summary",
    "https://www.iea.org/news/data-centre-electricity-use-surged-in-2025-even-with-tightening-bottlenecks-driving-a-scramble-for-solutions"
  ]);
}

// Slide 25 — model economics and reliability.
{
  const s = deck.slides.getItem(24);
  setText(s, "text-42-24", "证据与行业观察");
  setText(s, "text-42-55", "模型更强、更便宜，但可靠性仍呈“锯齿状前沿”");
  setText(s, "text-1178-674", "25");
  setText(s, "text-42-158", "性能趋同与成本下降，正在把竞争焦点推向可靠性、场景适配和端到端交付。");
  setText(s, "text-66-320", "01");
  setText(s, "text-66-380", "推理成本 ↓280×");
  setText(s, "text-66-446", "2022.11—2024.10，达到 GPT‑3.5 水平的推理成本下降超过280倍。");
  setText(s, "text-476-320", "02");
  setText(s, "text-476-380", "Agent 成功率 ≈66%");
  setText(s, "text-476-446", "OSWorld任务成功率由12%升至约66%，但结构化基准上仍约三分之一失败。");
  setText(s, "text-886-320", "03");
  setText(s, "text-886-380", "中美差距 ≈2.7%");
  setText(s, "text-886-446", "截至2026年3月，美国与中国头部模型的性能差距已缩至约2.7%。");
  setNotes(s, [
    "https://hai.stanford.edu/news/ai-index-2025-state-of-ai-in-10-charts",
    "https://hai.stanford.edu/ai-index/2026-ai-index-report",
    "https://hai.stanford.edu/ai-index/2026-ai-index-report/technical-performance"
  ]);
}

function editMatrix(s, page, kicker, title, headers, rows, urls, context="") {
  setText(s, "text-42-24", kicker);
  setText(s, "text-42-55", title);
  setText(s, "text-1178-674", String(page));
  const headerNames = ["text-56-169","text-274-169","text-564-169","text-874-169"];
  headerNames.forEach((n,i)=>setText(s,n,headers[i]));
  const ys=[219,301,383,465,547];
  for (let r=0;r<5;r++) {
    const names=[`text-56-${ys[r]}`,`text-274-${ys[r]}`,`text-564-${ys[r]}`,`text-874-${ys[r]}`];
    names.forEach((n,i)=>setText(s,n,rows[r][i]));
  }
  setNotes(s, urls, context);
}

// Slide 26 — agent protocol landscape.
editMatrix(
  deck.slides.getItem(25), 26, "开放标准",
  "Agent 标准正把“点对点集成”改造成可组合生态",
  ["标准 / 控制面","连接对象","核心作用","治理与企业含义"],
  [
    ["MCP","Agent ↔ 工具/数据","统一发现、上下文与工具调用","2025.12进入AAIF，走向中立治理"],
    ["A2A","Agent ↔ Agent","跨框架协作、能力发现与任务交接","2025.06进入Linux Foundation"],
    ["AGENTS.md","开发Agent ↔ 代码库","传递项目级上下文、规则与操作边界","AAIF创始项目之一"],
    ["支付/商务协议","Agent ↔ 商业系统","表达授权、意图、凭证与交易证明","将Agent扩展到可验证商业活动"],
    ["企业控制面","身份 ↔ 全链路","权限、策略、观测、审计与回滚","开放协议不能替代企业安全边界"]
  ],
  [
    "https://www.anthropic.com/news/donating-the-model-context-protocol-and-establishing-of-the-agentic-ai-foundation",
    "https://cloud.google.com/blog/products/ai-machine-learning/agent2agent-protocol-is-getting-an-upgrade",
    "https://www.linuxfoundation.org/blog/linux-foundation-newsletter-december-2025"
  ],
  "说明：支付/商务协议为行业演进方向概括；企业控制面为基于开放协议边界推导的企业架构要求。"
);

// Slide 27 — application patterns.
{
  const s=deck.slides.getItem(26);
  setText(s,"text-42-24","行业应用");
  setText(s,"text-42-55","行业落地呈现六种共性模式：知识、代码、服务、运营、科学与决策");
  setText(s,"text-1178-674","27");
  const rows=[
    ["知识工作","搜索 · 摘要 · 写作 · 分析｜现实使用中约57%偏增能、43%偏自动化"],
    ["软件研发","代码生成 · 测试 · 迁移 · 运维｜研究汇总显示生产率提升约26%"],
    ["客户运营","客服 · 销售 · 营销｜结构化、可测量、反馈快的任务最先兑现"],
    ["产业运营","质检 · 排产 · 供应链 · 设备维护｜需连接OT/IoT、规则与审批"],
    ["科学与医疗","文献 · 实验 · 诊疗辅助｜潜在价值高，但验证和责任要求更重"],
    ["管理决策","经营分析 · 资源配置 · 风险预警｜必须保留可解释的证据链"]
  ];
  const left=["text-62-169","text-62-239","text-62-309","text-62-379","text-62-449","text-62-519"];
  const right=["text-260-170","text-260-240","text-260-310","text-260-380","text-260-450","text-260-520"];
  rows.forEach((row,i)=>{setText(s,left[i],row[0]);setText(s,right[i],row[1]);});
  setText(s,"text-56-602","规模化共性：明确任务边界｜建立可验证输出｜接入真实系统｜异常升级｜持续评测");
  setNotes(s,[
    "https://www.anthropic.com/news/the-anthropic-economic-index",
    "https://hai.stanford.edu/ai-index/2026-ai-index-report/economy",
    "https://www.cac.gov.cn/2025-08/27/c_1758018277755538.htm"
  ],"生产率数字是 Stanford 2026 AI Index 对研究结果的汇总，不代表所有组织都能获得相同收益。");
}

// Slide 28 — policy landscape.
editMatrix(
  deck.slides.getItem(27), 28, "政策与治理",
  "政策环境从原则走向目标、时间表与可操作要求",
  ["地区 / 框架","关键节点","核心要求","企业影响"],
  [
    ["中国“AI+”","2027 / 2030","终端与智能体普及率>70% / >90%","六大重点领域推进深度融合"],
    ["EU：GPAI / 透明","2025.08 / 2026.08","GPAI治理；AI交互与生成内容透明","建立模型清单、标识与披露机制"],
    ["EU：高风险","2027.12 / 2028.08","风险评估、数据质量、日志、人类监督","对敏感场景实施分级准入"],
    ["NIST AI RMF","持续更新","Govern · Map · Measure · Manage","用生命周期方法建立控制体系"],
    ["企业治理基线","现在","身份、数据、评测、审计、事件响应","将合规变成产品与平台能力"]
  ],
  [
    "https://www.cac.gov.cn/2025-08/27/c_1758018277755538.htm",
    "https://digital-strategy.ec.europa.eu/en/policies/regulatory-framework-ai",
    "https://www.nist.gov/itl/ai-risk-management-framework"
  ],
  "欧盟日期依据 European Commission 2026-07 更新后的实施时间表。"
);

// Slide 29 — measurement scorecard.
{
  const s=deck.slides.getItem(28);
  setText(s,"text-42-24","价值衡量");
  setText(s,"text-42-55","行业智能化必须用五类指标共同验收");
  setText(s,"text-1178-674","29");
  const rows=[
    ["01","业务价值","收入增量、单位成本、处理周期、产能或风险损失改善；设置基线与对照组。"],
    ["02","任务质量","正确率、完成率、人工返工率、异常升级率；用真实任务持续回归。"],
    ["03","风险控制","越权、敏感数据暴露、幻觉/误导、审计完整性，以及事件数量与严重度。"],
    ["04","单位经济性","单任务推理与工具成本、延迟、资源利用率，以及规模化后的边际成本。"],
    ["05","采用与组织","活跃用户、覆盖流程、人工接管、培训与满意度；避免只看调用量。"]
  ];
  const ys=[156,250,344,438,532];
  rows.forEach((row,i)=>{setText(s,`text-42-${ys[i]}`,row[0]);setText(s,`text-120-${ys[i]}`,row[1]);setText(s,`text-360-${ys[i]}`,row[2]);});
  setNotes(s,[
    "https://hai.stanford.edu/ai-index/2026-ai-index-report/economy",
    "https://www.nist.gov/itl/ai-risk-management-framework"
  ],"该五类指标体系综合了价值实现、持续评测和生命周期风险管理要求。");
}

await fs.mkdir(RENDER,{recursive:true});
for (const [i,slide] of deck.slides.items.entries()) {
  const stem=`slide-${String(i+1).padStart(2,"0")}`;
  const png=await deck.export({slide,format:"png",scale:1});
  await fs.writeFile(`${RENDER}/${stem}.png`,new Uint8Array(await png.arrayBuffer()));
  const layout=await slide.export({format:"layout"});
  await fs.writeFile(`${RENDER}/${stem}.layout.json`,await layout.text());
}
const inspect=await deck.inspect({kind:"slide,textbox,shape,notes,layout",maxChars:250000});
await fs.writeFile(`${RENDER}/final-inspect.ndjson`,inspect.ndjson);
const pptx=await PresentationFile.exportPptx(deck);
await pptx.save(OUT);
console.log(`Updated ${OUT}: ${deck.slides.items.length} slides`);
