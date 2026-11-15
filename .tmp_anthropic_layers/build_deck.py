from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET


OUT = Path('/Users/apple/SocratesAI-codex/Anthropic分层.pptx')

W, H = 13.333333333333334, 7.5
prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
slide = prs.slides.add_slide(prs.slide_layouts[6])


def rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


INK = '#171717'
MUTED = '#626262'
RULE = '#C9C9C9'
PANEL = '#F2F2F2'
PALE = '#FAF8F6'
ACCENT = '#D97757'
ACCENT_DARK = '#9D4D37'
BLUE = '#4E8ED8'
FONT = 'PingFang SC'


def add_text(x, y, w, h, text, size=16, color=INK, bold=False,
             align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, margin=0,
             name=None, font=FONT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.clear()
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = rgb(color)
    return box


def add_rect(x, y, w, h, fill, line=RULE, radius=True, name=None, lw=1.0):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    s.line.color.rgb = rgb(line)
    s.line.width = Pt(lw)
    if radius:
        try:
            s.adjustments[0] = 0.08
        except Exception:
            pass
    return s


def add_line(x1, y1, x2, y2, color=RULE, width=1.2, dash=False, arrow=False, name=None):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    if name:
        c.name = name
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(width)
    if dash:
        c.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    if arrow:
        # python-pptx has no public arrowhead API; add a small triangle separately.
        pass
    return c


def add_arrowhead(x, y, direction='down', color=ACCENT, size=0.12):
    tri = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                 Inches(x-size/2), Inches(y-size/2), Inches(size), Inches(size))
    tri.fill.solid(); tri.fill.fore_color.rgb = rgb(color)
    tri.line.color.rgb = rgb(color)
    rotations = {'down': 180, 'up': 0, 'right': 90, 'left': 270}
    tri.rotation = rotations[direction]
    return tri


# Canvas and header
bg = slide.background.fill
bg.solid(); bg.fore_color.rgb = rgb('#FFFFFF')
add_text(0.52, 0.25, 2.35, 0.22, 'STRATEGY SNAPSHOT · 2026', 10.5, MUTED, True, name='eyebrow')
add_text(0.52, 0.55, 12.2, 0.55,
         'Anthropic的护城河：模型 × 能力 × 执行闭环',
         35, INK, True, name='title')
add_text(0.52, 1.10, 11.9, 0.32,
         'Claude 提供通用智能；MCP 与 Skills 将其接入业务；Harness 把能力约束为可持续交付的系统闭环。',
         14, MUTED, False, name='subtitle')
add_line(0.52, 1.48, 12.80, 1.48, RULE, 1.0, name='header-rule')

# Connectors first: main upward value flow + feedback return loop.
cx = 7.55
for y_bottom, y_top in [(4.79, 4.67), (3.79, 3.67), (2.79, 2.67)]:
    add_line(cx, y_bottom, cx, y_top, ACCENT, 2.0, name=f'value-flow-{y_bottom}')
    add_arrowhead(cx, y_top, 'up', ACCENT, 0.12)

# Left flywheel loop connectors behind nodes.
lx = 1.60
for y1, y2 in [(2.89, 3.25), (3.89, 4.25), (4.89, 5.25)]:
    add_line(lx, y1, lx, y2, ACCENT, 1.8, name=f'flywheel-down-{y1}')
    add_arrowhead(lx, y2, 'down', ACCENT, 0.11)
add_line(1.22, 5.40, 0.88, 5.40, ACCENT, 1.8, name='flywheel-return-1')
add_line(0.88, 5.40, 0.88, 2.00, ACCENT, 1.8, name='flywheel-return-2')
add_line(0.88, 2.00, 1.22, 2.00, ACCENT, 1.8, name='flywheel-return-3')
add_arrowhead(1.21, 2.00, 'right', ACCENT, 0.11)

# Left flywheel panel.
add_text(0.52, 1.72, 2.05, 0.30, '受控数据飞轮', 18, INK, True, name='flywheel-title')
add_text(0.52, 2.03, 2.08, 0.46, '评测与真实任务信号\n持续校准模型与系统', 11.5, MUTED, name='flywheel-subtitle')

nodes = [
    (1.24, 2.63, '训练数据', '公开 / 授权 / 合成'),
    (1.24, 3.63, 'Claude迭代', '预训练 + 后训练'),
    (1.24, 4.63, '真实任务', '行业场景与工具调用'),
    (1.24, 5.63, '评测反馈', 'Evals / 红队 / 失败轨迹'),
]
for i, (x, y, head, body) in enumerate(nodes):
    add_rect(x, y-0.36, 1.65, 0.65, '#FFFFFF', ACCENT if i in (0,3) else RULE,
             True, f'flywheel-node-{i}', 1.2)
    add_text(x+0.12, y-0.27, 1.42, 0.22, head, 13, INK, True, name=f'flywheel-head-{i}')
    add_text(x+0.12, y-0.03, 1.42, 0.20, body, 9.5, MUTED, name=f'flywheel-body-{i}')

# Main stack background and four bands.
stack_x, stack_w = 3.25, 9.55
layer_specs = [
    (1.78, 0.90, '#ECE7E3', '04  行业方案',
     '研发｜金融｜客服｜科研：交付可完成任务的智能方案',
     '价值兑现'),
    (2.78, 0.90, '#F2F2F2', '03  Harness · 可靠执行',
     '任务分解｜状态管理｜权限沙箱｜可观测｜评测｜恢复',
     '系统闭环'),
    (3.78, 0.90, '#F7F7F7', '02  MCP + Skills · 业务能力',
     '连接数据与工具；封装流程、脚本与行业知识',
     '能力复用'),
    (4.78, 0.90, '#FAF8F6', '01  Claude · 核心智能',
     '推理与多模态｜API / Code / SDK｜安全对齐',
     '智能底座'),
]

for idx, (y, h, fill, head, body, tag) in enumerate(layer_specs):
    add_rect(stack_x, y, stack_w, h, fill, ACCENT if idx == 3 else RULE,
             True, f'layer-{4-idx}', 1.2 if idx == 3 else 0.9)
    add_text(stack_x+0.24, y+0.14, 3.50, 0.38, head, 20, INK, True, name=f'layer-head-{idx}')
    add_text(stack_x+3.80, y+0.10, 4.80, 0.58, body, 14.5, MUTED, False,
             valign=MSO_ANCHOR.MIDDLE, name=f'layer-body-{idx}')
    add_text(stack_x+8.70, y+0.24, 0.60, 0.25, tag, 10.5,
             ACCENT_DARK if idx in (0,3) else MUTED, True, PP_ALIGN.CENTER,
             valign=MSO_ANCHOR.MIDDLE, name=f'layer-tag-{idx}')

# Flywheel link labels between side loop and stack.
add_text(2.74, 2.13, 0.42, 0.32, '反馈', 9.5, ACCENT_DARK, True, PP_ALIGN.CENTER)
add_line(2.85, 2.50, 3.22, 2.50, ACCENT, 1.4)
add_arrowhead(3.20, 2.50, 'right', ACCENT, 0.10)
add_text(2.71, 5.19, 0.48, 0.32, '学习', 9.5, ACCENT_DARK, True, PP_ALIGN.CENTER)
add_line(3.22, 5.42, 2.90, 5.42, ACCENT, 1.4)
add_arrowhead(2.92, 5.42, 'left', ACCENT, 0.10)

# Bottom thesis strip.
add_rect(3.25, 5.92, 9.55, 0.78, '#171717', '#171717', True, 'thesis-strip', 0)
add_text(3.52, 6.08, 1.36, 0.27, '核心洞察', 13, '#FFFFFF', True, name='insight-label')
add_text(4.84, 6.02, 7.62, 0.39,
         '壁垒 = 模型 × 连接标准 × 业务知识 × 执行闭环',
         18, '#FFFFFF', True, valign=MSO_ANCHOR.MIDDLE, name='insight')

# Boundary note + footer.
add_text(0.52, 6.83, 11.70, 0.28,
         '边界提示：公开资料显示 Anthropic 默认不使用商业客户输入/输出训练模型；“数据飞轮”更适合解释为评测、合成数据、红队和明确授权反馈驱动的受控改进机制。',
         9.3, MUTED, False, name='boundary-note')
add_text(12.28, 6.83, 0.52, 0.24, '01', 10, MUTED, True, PP_ALIGN.RIGHT, name='page-number')

# Core properties.
prs.core_properties.title = 'Anthropic分层'
prs.core_properties.subject = 'Anthropic端到端智能系统分层与数据飞轮洞察'
prs.core_properties.author = 'OpenAI Codex'
prs.core_properties.keywords = 'Anthropic, Claude, MCP, Skills, Harness, 数据飞轮'
prs.core_properties.comments = ('Sources: Anthropic MCP introduction; Agent Skills; Effective harnesses for '
                                'long-running agents; Demystifying evals; Anthropic Transparency Hub. '
                                'The four-layer model and controlled data flywheel are analytical synthesis.')

OUT.parent.mkdir(parents=True, exist_ok=True)
raw_out = OUT.with_name('Anthropic分层.compat-build.pptx')
prs.save(raw_out)

# python-pptx inherits the default template's 4:3 type marker even after resizing.
# Normalize the package to Microsoft's exact 16:9 dimensions and "wide" marker.
with ZipFile(raw_out, 'r') as zin, ZipFile(OUT, 'w', ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'ppt/presentation.xml':
            root = ET.fromstring(data)
            ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
            size = root.find('p:sldSz', ns)
            size.set('cx', '12192000')
            size.set('cy', '6858000')
            size.set('type', 'wide')
            data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        zout.writestr(item, data)
raw_out.unlink()
print(OUT)
