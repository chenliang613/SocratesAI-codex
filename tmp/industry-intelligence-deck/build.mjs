import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "/Users/apple/SocratesAI-codex/行业智能化分析.pptx";
const RENDER_DIR = "/Users/apple/SocratesAI-codex/tmp/industry-intelligence-deck/rendered";
const W = 1280, H = 720;
const C = { ink: "#000000", muted: "#5F6670", panel: "#EDEDED", rule: "#B8BCC4", blue: "#3D8DFF", sky: "#6DCBF4", pale: "#D0EDFA", white: "#FFFFFF", dark: "#15171A" };
const FONT = "PingFang SC";

const p = Presentation.create({ slideSize: { width: W, height: H } });

function text(slide, value, left, top, width, height, size=20, opts={}) {
  const s = slide.shapes.add({
    geometry: "textbox", name: opts.name || `text-${left}-${top}`,
    position: { left, top, width, height }, fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = value;
  s.text.style = {
    fontSize: size, typeface: FONT, color: opts.color || C.ink,
    bold: !!opts.bold, alignment: opts.align || "left",
    verticalAlignment: opts.valign || "top", autoFit: opts.autoFit || "shrinkText",
  };
  return s;
}

function rect(slide, left, top, width, height, fill=C.panel, opts={}) {
  return slide.shapes.add({
    geometry: opts.geometry || "rect", name: opts.name || `rect-${left}-${top}`,
    position: { left, top, width, height }, fill,
    line: opts.line || { style: "solid", fill: opts.stroke || fill, width: opts.lineWidth ?? 0 },
    borderRadius: opts.radius,
  });
}

function line(slide, left, top, width, height=0, color=C.rule, weight=1) {
  return slide.shapes.add({ geometry: "line", position: { left, top, width, height }, fill: "none", line: { style: "solid", fill: color, width: weight } });
}

function base(title, page, kicker="行业智能化分析") {
  const slide = p.slides.add();
  slide.background.fill = C.white;
  text(slide, kicker, 42, 24, 400, 26, 14, { color: C.muted, bold: true });
  text(slide, title, 42, 55, 1160, 68, 42, { bold: true });
  line(slide, 42, 128, 1196, 0, C.rule, 1);
  text(slide, String(page).padStart(2,"0"), 1178, 674, 60, 22, 13, { color: C.muted, align: "right" });
  return slide;
}

function notes(slide, extra="") {
  slide.speakerNotes.textFrame.setText(`[Sources]\n- /Users/apple/SocratesAI-codex/AI_Layout.md\n[/Sources]${extra ? `\n\n${extra}` : ""}`);
}

function bulletList(slide, items, x, y, w, h, size=20, color=C.ink) {
  return text(slide, items.map(v => `• ${v}`).join("\n"), x, y, w, h, size, { color });
}

// 1 Cover — Codex Grid cover silhouette.
{
  const s = p.slides.add(); s.background.fill = C.white;
  text(s, "战略分析报告 · 2026", 42, 38, 360, 34, 20, { color: C.muted, bold: true });
  text(s, "行业智能化\n分析", 42, 176, 850, 230, 82, { bold: true, autoFit: "none" });
  text(s, "从能源与算力的物理约束，到 Agent 与应用的价值兑现", 42, 500, 850, 70, 28, { color: C.muted });
  rect(s, 1050, 38, 188, 594, C.pale);
  rect(s, 1050, 431, 188, 201, C.blue);
  text(s, "六层架构\n跨层竞争\n企业路径", 1080, 466, 130, 130, 24, { color: C.white, bold: true });
  notes(s);
}

// 2 Executive thesis.
{
  const s = base("行业智能化不是“加一个模型”，而是重构整条价值链", 2);
  text(s, "电力、算力、模型、数据、流程与场景共同构成结果交付系统；任一层薄弱，都会成为规模化瓶颈。", 42, 158, 1160, 78, 28, { bold: true });
  const xs = [42, 452, 862];
  const heads = ["底层决定上限", "中层决定可用性", "上层决定回报"];
  const bodies = [
    "能源与算力具有重资产、高壁垒和寡头化特征，决定智能供给的规模、成本与稳定性。",
    "模型与数据平台把通用能力转化为可信、可治理、可调用的企业智能。",
    "Agent 与应用贴近业务流程，竞争焦点从“功能使用”转向“结果交付”。"
  ];
  xs.forEach((x,i)=>{ rect(s,x,300,376,264,i===1?C.pale:"#F4F4F4"); text(s,`0${i+1}`,x+24,320,60,42,22,{color:C.blue,bold:true}); text(s,heads[i],x+24,380,320,44,28,{bold:true}); text(s,bodies[i],x+24,446,320,92,19,{color:C.muted}); });
  notes(s);
}

// 3 Definition & shift.
{
  const s = base("行业智能化的本质：把“信息系统”升级为“行动系统”", 3);
  text(s, "传统数字化", 42, 176, 540, 40, 28, { bold: true });
  text(s, "行业智能化", 698, 176, 540, 40, 28, { bold: true, color: C.blue });
  line(s, 640, 168, 0, 430, C.rule, 2);
  const left = ["人发起操作，系统被动响应", "以界面与功能模块为中心", "跨系统协同依赖人工搬运", "价值衡量偏向上线率、活跃度"];
  const right = ["目标触发任务，Agent 主动编排", "以语义、工具和业务结果为中心", "跨数据、模型与系统自动协作", "价值衡量转向周期、质量、成本与风险"];
  bulletList(s,left,42,246,540,310,23,C.muted); bulletList(s,right,698,246,540,310,23,C.ink);
  rect(s,698,575,540,48,C.blue); text(s,"关键变化：软件从“供人操作”转向“可被智能体调用”",718,585,500,30,19,{color:C.white,bold:true});
  notes(s);
}

// 4 Six-layer architecture diagram.
{
  const s = base("六层架构共同完成“电力 → 智能 → 业务结果”的转化", 4);
  const layers = [
    ["应用层 Application","把能力嵌入销售、客服、研发、财务等场景，兑现 ROI"],
    ["Agent 层 Agent","理解意图、规划任务、调用工具、记忆状态、执行与审批"],
    ["数据平台层 Data","统一数据底座、治理、语义层、向量检索与知识库"],
    ["模型层 Model","预训练、推理、对齐、多模态与 API 服务"],
    ["算力层 Compute","芯片、服务器、云、网络互联与软件栈"],
    ["能源层 Energy","电力、数据中心、散热、电网与新能源"]
  ];
  layers.forEach((d,i)=>{ const y=154+i*78; const fill=i<2?C.pale:(i===2?"#E8EEF4":(i===3?"#ECECEC":"#F4F4F4")); rect(s,126,y,1028,62,fill); text(s,d[0],150,y+13,300,34,24,{bold:true,color:i<2?C.blue:C.ink}); text(s,d[1],470,y+14,650,30,19,{color:C.muted}); });
  text(s,"重资产 · 高壁垒 · 寡头",42,634,360,28,17,{color:C.muted}); line(s,342,649,588,0,C.blue,3); text(s,"轻资产 · 近场景 · 多样化",948,634,290,28,17,{color:C.muted,align:"right"});
  notes(s);
}

// 5 Value and control points.
{
  const s = base("价值流向上兑现，约束却从底层逐级向上传导", 5);
  const cols=[42,348,654,960];
  const heads=["供给约束","能力转化","组织嵌入","结果兑现"];
  const bodies=["电力、芯片、互联与资本开支决定可获得的计算规模。","模型质量、数据语义与推理成本决定智能是否可用。","Agent 的可靠性、权限和流程编排决定能否进入生产。","应用的场景设计与变革管理决定能否形成可量化 ROI。"];
  cols.forEach((x,i)=>{ text(s,`0${i+1}`,x,180,56,36,20,{color:C.blue,bold:true}); text(s,heads[i],x,230,244,44,26,{bold:true}); line(s,x,294,244,0,i===0?C.blue:C.rule,3); text(s,bodies[i],x,320,244,170,20,{color:C.muted}); if(i<3) text(s,"→",x+258,246,38,40,28,{color:C.blue,bold:true,align:"center"}); });
  rect(s,42,548,1160,74,C.dark); text(s,"管理含义：不要只选择“最强模型”，要识别本企业最稀缺的那一层，并围绕瓶颈配置投资。",70,566,1104,38,24,{color:C.white,bold:true,align:"center"});
  notes(s);
}

function layerSlide(page, title, definition, value, bottlenecks, trends, players, accent=C.blue) {
  const s=base(title,page,"六层架构 · 深度拆解");
  text(s,definition,42,154,1160,76,27,{bold:true});
  const x=[42,442,842], heads=["核心价值","关键瓶颈","结构性趋势"], vals=[value,bottlenecks,trends];
  x.forEach((xx,i)=>{ line(s,xx,272,344,0,i===0?accent:C.rule,4); text(s,heads[i],xx,294,344,40,25,{bold:true}); text(s,vals[i],xx,350,344,162,19,{color:C.muted}); });
  rect(s,42,550,1144,72,"#F2F2F2"); text(s,"代表参与者",66,569,130,30,20,{bold:true,color:accent}); text(s,players,210,563,944,46,17,{color:C.ink});
  notes(s); return s;
}

layerSlide(6,"能源层：AI 的物理地基正从成本项变成战略资源","大模型训练与推理，本质上是把稳定电力转化为可调用的智能。","为高密度数据中心持续提供稳定、廉价、低碳的电力与散热能力。","电网容量与并网周期、电价波动、碳排放约束、液冷能力与 PUE。","科技巨头通过长期购电协议、核电与新能源绑定供给；选址开始围绕“电力可获得性”重构。","国家电网、南方电网、NextEra、Constellation；GDS、Equinix、Digital Realty；AWS、微软、谷歌");
layerSlide(7,"算力层：利润集中于“芯片 + 互联 + 软件生态”的协同","算力层把电力转化为模型可消费的并行计算，是当前壁垒最高、资本最密集的环节。","提供训练与推理所需的 GPU/加速器、服务器、云算力、网络互联和开发栈。","先进制程、HBM、封装与互联、集群利用率，以及 CUDA 等软件生态锁定。","训练继续寡头化；推理因成本、延迟和场景差异而多元化；云厂商加速自研芯片。","NVIDIA、AMD、Intel；Google TPU、AWS Trainium/Inferentia、Microsoft Maia；华为昇腾、寒武纪；TSMC、SK 海力士",C.blue);
layerSlide(8,"模型层：通用能力趋同后，行业知识密度成为差异化来源","模型层把算力与数据转化为理解、推理、生成和工具使用能力，并通过 API 或权重供给上层。","提升推理、多模态、长上下文、工具使用与可控生成能力。","高质量数据、对齐与安全、推理成本、幻觉、评测和行业适配。","闭源前沿与开放权重并行；模型厂商向 Agent 框架与工具协议延伸，模型与 Agent 边界变薄。","OpenAI、Anthropic、Google、xAI；Meta、Mistral、DeepSeek、Qwen、GLM；Midjourney、ElevenLabs",C.sky);
layerSlide(9,"数据平台层：企业智能的上限取决于“同一种业务语言”","数据平台把分散、异构、权限复杂的数据，转换为模型和 Agent 可理解、可信任、可追溯的上下文。","采集、存储、治理、实时处理、语义统一、向量检索与知识服务。","数据孤岛、指标口径冲突、血缘不清、权限错配、实时性与合规。","湖仓与数据云向 AI 原生演进；语义层、向量检索和 Agent 数据接口成为标准能力。","Snowflake、Databricks、BigQuery、MaxCompute；Pinecone、Milvus；Confluent、dbt、Fivetran",C.sky);
layerSlide(10,"Agent 层：它把模型从“会回答”推进到“会完成任务”","Agent 组合模型、数据、工具、记忆和护栏，对目标进行分解、执行、检查与交付。","承担意图理解、任务规划、工具调用、状态管理、审批与多 Agent 协作。","可靠性、可控性、权限边界、错误恢复、工具协议与端到端评测。","从单点 Copilot 走向跨系统自主执行；Agent 入口成为新的人机界面与平台控制点。","Google Gemini/A2A、Microsoft Copilot Studio、OpenAI、Anthropic/MCP、阿里百炼、字节 Coze、腾讯元宝；Manus、Devin、Sierra",C.blue);
layerSlide(11,"应用层：竞争从“功能丰富”转向“结果可交付”","应用层将智能嵌入真实流程，是产业价值最终变成收入、效率、质量与风险改善的地方。","围绕销售、客服、营销、研发、财务、供应链等场景交付业务结果。","场景标准化、行业 Know-how、系统集成、流程责任重构和用户信任。","SaaS 走向 Agentic SaaS 与去 UI 化；能力被调用的频率，逐步替代页面访问量成为价值指标。","Salesforce、SAP、ServiceNow、Workday；Notion、Adobe；Harvey、Abridge、Cursor、Intercom；金蝶、用友、销售易",C.blue);

// 12 Economics.
{
  const s=base("六层的经济结构不同：越靠底层越集中，越靠上层越分散",12,"跨层规律");
  const rows=[
    ["能源 / 算力","重资产、供给受限","规模、资本与工程壁垒","高集中度"],
    ["模型","研发密集、边际成本下降","能力、生态与分发","头部集中 + 开放生态"],
    ["数据平台","迁移与治理成本高","语义、信任与数据引力","平台化整合"],
    ["Agent","标准尚未固化","入口、工具网络与可靠性","竞争最激烈"],
    ["应用","轻资产、近客户","流程 Know-how 与交付","高度多样化"]
  ];
  const xs=[42,260,550,860], widths=[200,270,290,378];
  ["层级","成本结构","主要壁垒","市场形态"].forEach((h,i)=>{rect(s,xs[i],158,widths[i],46,C.dark);text(s,h,xs[i]+14,169,widths[i]-28,26,18,{color:C.white,bold:true});});
  rows.forEach((r,ri)=>r.forEach((v,i)=>{const y=204+ri*82;rect(s,xs[i],y,widths[i],82,ri%2?"#F6F6F6":C.white,{stroke:C.rule,lineWidth:1});text(s,v,xs[i]+14,y+15,widths[i]-28,54,i===0?20:18,{bold:i===0,color:i===0?C.blue:C.ink});}));
  notes(s);
}

// 13 Vertical integration matrix.
{
  const s=base("巨头通过纵向整合控制成本、入口与数据回路",13,"跨层竞争");
  const companies=["Google","Microsoft","Amazon","Alibaba"], layers=["能源/数据中心","算力/云","模型","Agent","应用/入口"];
  const marks={Google:[1,1,1,1,1],Microsoft:[1,1,1,1,1],Amazon:[1,1,1,1,1],Alibaba:[0,1,1,1,1]};
  const x0=354,y0=176,cw=166,rh=78;
  text(s,"企业",42,y0,250,40,22,{bold:true}); layers.forEach((l,i)=>text(s,l,x0+i*cw,y0,cw-10,54,17,{bold:true,align:"center"}));
  companies.forEach((co,r)=>{const y=y0+64+r*rh; text(s,co,42,y+18,250,40,24,{bold:true}); line(s,42,y+rh-4,1196,0,C.rule,1); layers.forEach((l,i)=>{rect(s,x0+i*cw+54,y+16,46,46,marks[co][i]?C.blue:"#F1F1F1",{geometry:"ellipse"}); text(s,marks[co][i]?"●":"—",x0+i*cw+54,y+19,46,34,20,{color:marks[co][i]?C.white:C.rule,bold:true,align:"center"});});});
  text(s,"共同逻辑",42,570,160,30,20,{bold:true,color:C.blue}); text(s,"向下掌握供给成本，向上掌握用户入口，中间通过数据反馈形成持续优化闭环。",210,568,990,38,22,{bold:true});
  notes(s,"矩阵表示 AI_Layout.md 中列举的跨层布局方向，不代表各层自有资产比例相同。");
}

// 14 Agent battle.
{
  const s=base("Agent 层是价值重新分配的主战场",14,"跨层竞争");
  text(s,"模型厂商向上、应用厂商向下、平台厂商横向扩张，三股力量在 Agent 入口交汇。",42,154,1160,58,26,{bold:true});
  const xs=[42,452,862]; const heads=["模型厂商向上","应用厂商向下","平台厂商横向"];
  const bodies=["把模型、工具调用、协议和执行环境打包，直接承接用户任务。","把客户关系、行业流程和专有数据封装成垂直 Agent。","提供构建、治理、观测、权限与多 Agent 协作的 Agent OS。"];
  const wins=["胜负手：模型能力 + 开发者生态","胜负手：场景深度 + 分发入口","胜负手：标准 + 控制面 + 数据连接"];
  xs.forEach((x,i)=>{rect(s,x,258,376,276,i===1?C.pale:"#F2F2F2"); text(s,heads[i],x+24,285,328,40,27,{bold:true}); text(s,bodies[i],x+24,350,328,100,19,{color:C.muted}); line(s,x+24,466,328,0,C.rule,1); text(s,wins[i],x+24,484,328,34,18,{bold:true,color:C.blue});});
  rect(s,42,564,1196,58,C.dark); text(s,"对企业而言，关键不是押注单一 Agent，而是保留模型可替换、工具可组合、权限可审计的架构选择权。",64,578,1152,30,21,{color:C.white,bold:true,align:"center"});
  notes(s);
}

// 15 Maturity path.
{
  const s=base("企业将沿四级成熟度，从助手走向自主运营",15,"企业落地");
  const stages=[
    ["L1 助手化","单点生成与问答","人负责判断与执行"],
    ["L2 流程嵌入","在一个流程内检索、生成、建议","系统可写回，人负责审批"],
    ["L3 跨系统 Agent","跨数据与工具完成端到端任务","异常与高风险节点人工介入"],
    ["L4 自主运营","多 Agent 协作、持续优化目标","人管理策略、边界与结果"]
  ];
  const xs=[42,340,638,936];
  line(s,82,310,1050,0,C.rule,3);
  stages.forEach((v,i)=>{rect(s,xs[i]+26,286,48,48,i<2?C.sky:C.blue,{geometry:"ellipse"}); text(s,String(i+1),xs[i]+26,295,48,30,18,{color:C.white,bold:true,align:"center"}); text(s,v[0],xs[i],190,250,40,24,{bold:true}); text(s,v[1],xs[i],360,250,62,20,{bold:true,color:C.blue}); text(s,v[2],xs[i],444,250,80,18,{color:C.muted});});
  text(s,"演进原则",42,580,120,30,20,{bold:true}); text(s,"每升一级，都必须同步提升数据语义、工具接口、评测体系、权限治理和组织责任，而不是只增加模型自主性。",170,576,1050,42,21,{bold:true});
  notes(s);
}

// 16 Target architecture.
{
  const s=base("企业目标架构应把“智能面”叠加在既有业务与数据之上",16,"企业落地");
  const rows=[
    ["交互与场景面","员工助手 · 客户触点 · 垂直应用 · API"],
    ["Agent 编排面","规划 · 路由 · 记忆 · 工具调用 · 多 Agent 协作 · 审批"],
    ["模型服务面","模型网关 · 多模型路由 · Prompt/Context · 评测 · 成本控制"],
    ["数据与知识面","语义层 · RAG · 向量/知识库 · 主数据 · 权限 · 血缘"],
    ["工具与系统面","CRM · ERP · SCM · OA · 研发平台 · 自动化接口"],
    ["基础设施面","云/边/端算力 · 芯片 · 网络 · 数据中心 · 能源"]
  ];
  rows.forEach((r,i)=>{const y=156+i*70; rect(s,42,y,1196,56,i===1?C.pale:(i===3?"#E8EEF4":"#F5F5F5")); text(s,r[0],62,y+13,180,28,21,{bold:true,color:(i===1||i===3)?C.blue:C.ink}); text(s,r[1],260,y+14,950,28,18,{color:C.muted});});
  rect(s,42,596,1196,36,C.dark); text(s,"横向治理：身份与权限｜安全与合规｜可观测性｜质量评测｜成本与配额｜审计与责任",56,602,1168,24,18,{color:C.white,bold:true,align:"center"});
  notes(s);
}

// 17 Scenario prioritization matrix.
{
  const s=base("场景选择要同时看业务价值与可规模化条件",17,"企业落地");
  const left=172, top=176, size=430;
  rect(s,left,top,size/2,size/2,"#F5F5F5"); rect(s,left+size/2,top,size/2,size/2,C.pale); rect(s,left,top+size/2,size/2,size/2,"#F8F8F8"); rect(s,left+size/2,top+size/2,size/2,size/2,"#EAF5FB");
  line(s,left+size/2,top,0,size,C.white,3); line(s,left,top+size/2,size,0,C.white,3);
  text(s,"战略孵化",left+20,top+22,170,34,22,{bold:true}); text(s,"优先突破",left+235,top+22,170,34,22,{bold:true,color:C.blue}); text(s,"谨慎投入",left+20,top+238,170,34,22,{bold:true,color:C.muted}); text(s,"快速复制",left+235,top+238,170,34,22,{bold:true,color:C.blue});
  text(s,"高业务价值",50,top+4,105,30,17,{bold:true,align:"right"}); text(s,"低业务价值",50,top+390,105,30,17,{color:C.muted,align:"right"}); text(s,"低可规模化",left,top+444,160,28,17,{color:C.muted}); text(s,"高可规模化",left+270,top+444,160,28,17,{bold:true,align:"right"});
  text(s,"四项筛选标准",680,176,300,38,28,{bold:true});
  const criteria=["价值：收入、成本、周期、质量或风险改善是否显著？","数据：是否有可信、授权、可追溯的上下文？","流程：任务边界是否清晰，异常能否被识别和升级？","复制：能否通过模板、工具和语义层跨部门复用？"];
  criteria.forEach((v,i)=>{text(s,`0${i+1}`,680,248+i*86,44,28,18,{color:C.blue,bold:true}); text(s,v,738,242+i*86,470,62,19,{bold:i===0}); line(s,680,318+i*86,520,0,C.rule,1);});
  notes(s,"该优先级矩阵为基于 AI_Layout.md 产业规律推导的企业落地框架。");
}

// 18 Operating model.
{
  const s=base("规模化需要产品、平台、数据、治理四条线协同",18,"企业落地");
  const xs=[42,342,642,942]; const heads=["业务产品线","AI 平台线","数据产品线","风险治理线"];
  const bodies=["定义价值目标、流程边界、用户体验和采用机制；对业务结果负责。","提供模型网关、Agent 编排、工具注册、评测、观测和成本管理。","建设语义层、知识产品、数据契约、质量与权限；对上下文可信负责。","制定分级授权、红线、审计、隐私、安全与事件响应；对风险边界负责。"];
  xs.forEach((x,i)=>{text(s,heads[i],x,176,256,40,25,{bold:true}); line(s,x,232,256,0,i===1?C.blue:C.rule,4); text(s,bodies[i],x,266,256,170,19,{color:C.muted}); text(s,["Owner：业务负责人","Owner：平台负责人","Owner：数据负责人","Owner：风控/安全负责人"][i],x,474,256,30,18,{bold:true,color:C.blue});});
  rect(s,42,548,1156,74,C.dark); text(s,"共同节奏：以场景为单位建立“上线—评测—观测—复盘—扩展”闭环；平台标准由真实场景反向沉淀。",66,566,1108,40,22,{color:C.white,bold:true,align:"center"});
  notes(s);
}

// 19 Governance.
{
  const s=base("Agent 治理的核心不是限制能力，而是建立可控的授权链",19,"企业落地");
  const flow=["身份","意图","计划","工具","执行","结果","审计"];
  flow.forEach((v,i)=>{const x=42+i*168; if(i<6) line(s,x+120,250,48,0,C.blue,2); rect(s,x,214,120,72,i===3?C.blue:C.pale); text(s,v,x,235,120,30,22,{bold:true,align:"center",color:i===3?C.white:C.ink});});
  const cols=[42,440,838]; const heads=["事前：边界","事中：控制","事后：追责"];
  const bodies=["按角色、数据敏感度、工具风险和金额设置权限；高风险动作必须审批。","对工具调用、参数、上下文、异常和成本实时观测；支持暂停、回滚与人工接管。","保留输入、计划、调用、输出与审批记录；用评测和事件复盘持续更新策略。"];
  cols.forEach((x,i)=>{text(s,heads[i],x,374,340,38,25,{bold:true,color:i===1?C.blue:C.ink}); text(s,bodies[i],x,430,340,120,19,{color:C.muted});});
  text(s,"最低治理基线：最小权限｜敏感数据隔离｜高风险人工确认｜全链路日志｜可回滚｜持续评测",42,592,1160,34,20,{bold:true});
  notes(s);
}

// 20 Roadmap.
{
  const s=base("12 个月路线图：先建立可验证闭环，再扩展自主性",20,"行动路线");
  const phases=[
    ["0–3 个月","识别 3–5 个高价值场景；建立模型网关、数据接入和基础评测；上线人机协同试点。"],
    ["3–6 个月","沉淀语义层与工具目录；把试点嵌入流程；建立权限、日志、成本与质量看板。"],
    ["6–9 个月","扩展跨系统 Agent；复用场景模板；引入异常恢复、人工接管和持续评测。"],
    ["9–12 个月","形成多 Agent 协作；按业务指标优化；建立平台化运营和年度投资组合机制。"]
  ];
  const xs=[42,342,642,942];
  line(s,82,262,1050,0,C.rule,3);
  phases.forEach((v,i)=>{rect(s,xs[i]+22,238,48,48,i<2?C.sky:C.blue,{geometry:"ellipse"});text(s,String(i+1),xs[i]+22,247,48,30,18,{color:C.white,bold:true,align:"center"});text(s,v[0],xs[i],166,250,38,24,{bold:true});text(s,v[1],xs[i],328,250,174,19,{color:C.muted});});
  rect(s,42,550,1196,72,"#F2F2F2"); text(s,"阶段门",64,570,100,28,20,{color:C.blue,bold:true}); text(s,"只有当价值、质量、风险、成本四项指标同时达标，才提升自主性或扩大到下一批场景。",172,566,1034,36,22,{bold:true});
  notes(s,"路线图为基于 AI_Layout.md 六层架构推导的建议，可根据企业基础调整节奏。");
}

// 21 Decisions.
{
  const s=base("管理层需要尽快作出五个关键选择",21,"行动路线");
  const items=[
    ["01","价值锚点","以业务结果而非模型参数定义投资优先级。"],
    ["02","架构边界","选择可替换模型、可组合工具、可审计权限的开放架构。"],
    ["03","数据责任","把语义、质量、权限与知识更新明确到业务 Owner。"],
    ["04","自主性边界","按风险分级决定哪些任务自动执行、哪些必须人工批准。"],
    ["05","规模化机制","用平台标准、场景模板和评测基线推动跨部门复制。"]
  ];
  items.forEach((v,i)=>{const y=156+i*94; text(s,v[0],42,y,60,34,20,{bold:true,color:C.blue}); text(s,v[1],120,y,210,34,24,{bold:true}); text(s,v[2],360,y,820,44,20,{color:C.muted}); line(s,42,y+62,1196,0,C.rule,1);});
  notes(s);
}

// 22 Close.
{
  const s=p.slides.add(); s.background.fill=C.white;
  text(s,"结论",42,40,180,34,20,{color:C.muted,bold:true});
  text(s,"真正的行业智能化，\n是让智能稳定地交付结果",42,176,1050,210,72,{bold:true,autoFit:"none"});
  text(s,"竞争优势将来自三种能力的组合：底层资源可获得、企业语义可理解、业务流程可执行。",42,482,920,72,27,{color:C.muted});
  rect(s,1038,40,200,592,C.blue); text(s,"能源\n算力\n模型\n数据\nAgent\n应用",1080,140,120,370,28,{color:C.white,bold:true,align:"center"});
  notes(s);
}

await fs.mkdir(RENDER_DIR, { recursive: true });
for (const [i, slide] of p.slides.items.entries()) {
  const stem = `slide-${String(i+1).padStart(2,"0")}`;
  const png = await p.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(`${RENDER_DIR}/${stem}.png`, new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${RENDER_DIR}/${stem}.layout.json`, await layout.text());
}
const montage = await p.export({ format: "webp", montage: true, scale: 1 });
await fs.writeFile(`${RENDER_DIR}/montage.webp`, new Uint8Array(await montage.arrayBuffer()));
const snapshot = await p.inspect({ kind: "slide,textbox,shape,table,chart,notes", maxChars: 50000 });
await fs.writeFile(`${RENDER_DIR}/inspect.ndjson`, snapshot.ndjson);
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(OUT);
console.log(`Wrote ${OUT} with ${p.slides.items.length} slides`);
