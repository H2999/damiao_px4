from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUT = '/home/hzq/RM/px4-autopilot/无人机具身智能系统应用赛项目实现报告.docx'

def font(run, name='宋体', size=10.5, bold=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def border(cell, color='000000', size='4'):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders'); tcPr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = borders.find(qn('w:' + edge))
        if el is None:
            el = OxmlElement('w:' + edge); borders.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), size)
        el.set(qn('w:space'), '0'); el.set(qn('w:color'), color)

def margins(cell, top=90, start=110, bottom=90, end=110):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in('w:tcMar')
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar'); tcPr.append(tcMar)
    for side, value in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        el = tcMar.find(qn('w:' + side))
        if el is None:
            el = OxmlElement('w:' + side); tcMar.append(el)
        el.set(qn('w:w'), str(value)); el.set(qn('w:type'), 'dxa')

def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement('w:tblHeader'); el.set(qn('w:val'), 'true'); trPr.append(el)

def keep_next(p):
    pPr = p._p.get_or_add_pPr(); pPr.append(OxmlElement('w:keepNext'))

def para(doc, text='', indent=True, align=None, after=4):
    p = doc.add_paragraph()
    if indent: p.paragraph_format.first_line_indent = Pt(21)
    if align is not None: p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.25
    r = p.add_run(text); font(r)
    return p

def heading(doc, text, level=1):
    p = doc.add_paragraph(style='Heading %d' % level)
    r = p.add_run(text); font(r, name='黑体', size=16 if level == 1 else 12.5, bold=True)
    keep_next(p)
    return p

def bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.left_indent = Pt(21); p.paragraph_format.space_after = Pt(2)
        r = p.add_run(item); font(r)

def table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers)); t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.style = 'Table Grid'
    h = t.rows[0]; repeat_header(h)
    for i, value in enumerate(headers):
        c = h.cells[i]; c.text = ''; border(c); margins(c); c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(str(value)); font(r, size=9.2, bold=True)
        if widths: c.width = Inches(widths[i])
    for row in rows:
        cells = t.add_row().cells
        for i, value in enumerate(row):
            c = cells[i]; c.text = str(value); border(c); margins(c); c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.05
                for r in p.runs: font(r, size=9.0)
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(1)
    return t

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.72); sec.bottom_margin = Inches(0.65); sec.left_margin = Inches(0.82); sec.right_margin = Inches(0.72)
sec.header_distance = Inches(0.3); sec.footer_distance = Inches(0.3)

normal = doc.styles['Normal']; normal.font.name = '宋体'; normal._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体'); normal.font.size = Pt(10.5)
normal.paragraph_format.line_spacing = 1.25; normal.paragraph_format.space_after = Pt(4)
for level, size in [(1, 16), (2, 12.5), (3, 11)]:
    st = doc.styles['Heading %d' % level]
    st.font.name = '黑体'; st._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); st.font.size = Pt(size); st.font.bold = True; st.font.color.rgb = RGBColor(0,0,0)
    st.paragraph_format.space_before = Pt(8 if level == 1 else 5); st.paragraph_format.space_after = Pt(3)
title_style = doc.styles['Title']; title_style.font.name = '黑体'; title_style._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); title_style.font.size = Pt(21); title_style.font.bold = True; title_style.font.color.rgb = RGBColor(0,0,0)

head = sec.header.paragraphs[0]; head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r = head.add_run('无人机具身智能系统应用赛项目实现报告'); font(r, size=8.5)
foot = sec.footer.paragraphs[0]; foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = foot.add_run('2026 年华北五省（市、自治区）大学生机器人大赛'); font(r, size=8.5)

p = doc.add_paragraph(style='Title'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(60)
r = p.add_run('无人机具身智能系统应用赛'); font(r, name='黑体', size=21, bold=True)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(12)
r = p.add_run('项目实现报告'); font(r, name='黑体', size=17, bold=True)
for line in ['参赛方向：强国赛道', '系统名称：基于 PX4 与视觉感知的微型无人机具身智能系统', '参赛平台：1103 无刷电机四旋翼', '报告人：____________    日期：____________']:
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(10)
    r = p.add_run(line); font(r, size=11)
doc.add_page_break()

heading(doc, '摘要', 1)
para(doc, '本项目面向 2026 年华北五省（市、自治区）大学生机器人大赛强国赛道中的无人机具身智能系统应用赛，设计并实现一套面向室内封闭场地的微型四旋翼自主飞行系统。系统以 PX4 开源飞控完成姿态稳定、位置控制和飞行安全，以香橙派 Zero 3W 完成视觉计算和任务决策，以微空科技 Micovio 提供室内定位，以 IMX219 摄像头完成环境图像采集，并采用 YOLO 进行目标检测、OpenCV 进行二维码识别，结合视觉方法完成障碍物识别和避障控制。')
para(doc, '针对比赛规定的“自主起飞、避障导航、目标识别、无人车二维码扫描、按指令降落”任务，本文给出系统组成、感知与定位方法、任务状态机、控制策略、工程实现和验证计划。报告中的器件选择和算法方案是当前设计方案；电调具体型号、1103 电机 KV 值、桨叶参数、Micovio 接口和部分性能指标需在装机与测试后补录。')
para(doc, '关键词：具身智能；微型无人机；PX4；Micovio；YOLO；OpenCV；二维码识别；室内自主导航')

heading(doc, '一、项目与比赛任务分析', 1)
heading(doc, '1.1 比赛任务', 2)
para(doc, '根据比赛文件，本赛项线上提交项目实现报告，线下进行实体赛。线下自主导航模式满分 100 分，任务依次为：自主起飞；避开障碍物并导航至目标区域；识别并定位指定目标物体；搜索贴有二维码的无人车并扫描二维码；根据二维码中的降落指令飞往指定降落区；在指定区域安全降落。')
table(doc, ['比赛约束','本项目设计响应'], [
    ['场地：9 m × 5 m × 3 m 封闭区域','采用室内定位和视觉感知，不依赖 GNSS；所有路径和速度留出边界余量'],
    ['常规飞行高度不低于 1.2 m','将巡航高度设为不低于 1.2 m，识别和扫描阶段按裁判要求低速悬停'],
    ['单轮时间不超过 5 分钟','采用有限航点搜索、目标优先级和二维码确认机制，设置各状态超时'],
    ['两次飞行机会，至少一次自主模式','以自主模式作为主方案，遥控模式作为调试和风险兜底方案'],
    ['必须具备遥控和紧急安全开关','指定接管人员，保留遥控器和 PX4 失联/急停保护'],
    ['不得自行布置外部标志物或辅助装置','定位与识别全部由机载 Micovio、IMX219 和机载计算完成'],
], widths=[2.0, 4.5])
heading(doc, '1.2 评分导向', 2)
table(doc, ['评分项','分值','技术重点','本项目对应模块'], [
    ['自主起飞','15','解锁检查、起飞高度和稳定状态','PX4 起飞状态机、Micovio 定位'],
    ['避障与导航','20','障碍物识别、路径规划和航迹跟踪','视觉障碍物识别、PX4 位置控制'],
    ['目标识别','20','识别指定物体并确定位置','YOLO 目标检测'],
    ['二维码扫描','20','扫描无人车二维码并正确解读','IMX219 + OpenCV'],
    ['指定区域降落','15','按二维码指令选择降落区并安全降落','任务状态机、PX4 降落控制'],
    ['降落精度','10','无人机投影落在 H 区域内','Micovio 定位、视觉对准、低速下降'],
])

heading(doc, '二、系统总体设计', 1)
heading(doc, '2.1 设计目标', 2)
bullets(doc, ['以自主导航模式完成比赛全流程，遥控器仅用于安全接管，不作为常规控制手段。','在无 GNSS 的室内环境中完成连续定位、目标识别、二维码识别和路径跟踪。','将感知结果转化为可执行的飞行行为：搜索、悬停、避障、接近、确认和降落。','保证系统可在赛前快速检查、现场快速部署，并能用日志复盘每次飞行。'])
heading(doc, '2.2 系统架构', 2)
para(doc, '系统采用“实时飞控层 + 机载智能层”的分层架构。PX4 位于实时飞控层，负责 IMU 融合、姿态控制、位置控制、混控输出和失控保护。香橙派 Zero 3W 位于机载智能层，负责图像采集、YOLO 推理、OpenCV 二维码识别、视觉障碍物判断、任务状态机和与 PX4 的 MAVLink 通信。Micovio 为 PX4 提供室内定位或视觉里程计信息，IMX219 为香橙派提供任务视觉输入。')
table(doc, ['系统层','硬件/软件','输入','输出'], [
    ['飞行执行层','1103 电机、海科电调、寸级桨','PX4 四路电机控制量','推力、姿态和位移'],
    ['实时控制层','PX4 开源飞控','IMU、气压计、Micovio、遥控器','姿态/位置控制、电机输出、保护动作'],
    ['智能计算层','香橙派 Zero 3W、任务程序','图像、定位、飞控状态、电池状态','目标、二维码、障碍物和任务指令'],
    ['感知层','Micovio、IMX219','室内环境和目标','定位数据、视频帧'],
])
heading(doc, '2.3 数据闭环', 2)
para(doc, 'Micovio 输出定位数据并经过坐标转换、时间同步和质量检查后送入 PX4 估计器。IMX219 输出图像后，香橙派分别运行 YOLO 和 OpenCV：YOLO 输出目标类别、置信度和边界框，OpenCV 输出二维码文本和角点；视觉障碍物模块输出风险等级和安全距离。任务状态机融合这些结果与 PX4 飞行状态，生成速度、位置或航点指令，PX4 再将指令转换为姿态和电机输出。')

heading(doc, '三、无人机平台设计', 1)
heading(doc, '3.1 硬件组成', 2)
table(doc, ['子系统','选型','作用','待确认参数'], [
    ['机体','小型四旋翼 X 形结构','承载动力、飞控、视觉和计算设备','整机质量、轴距、重心、包络'],
    ['动力','1103 无刷电机 ×4，寸级桨','提供升力和姿态控制力矩','KV、桨叶规格、推力和电流'],
    ['电调','海科电调 ×4','执行 PX4 电机控制量','型号、电流、协议、输入电压'],
    ['飞控','PX4 开源飞控','姿态、位置、保护和电机混控','具体控制器型号、固件版本'],
    ['机载计算','香橙派 Zero 3W','运行视觉算法和任务状态机','供电、散热、推理帧率'],
    ['室内定位','微空科技 Micovio','输出室内位置/视觉里程计','接口、频率、坐标系、时延'],
    ['相机','IMX219','采集目标、二维码和障碍物图像','分辨率、帧率、镜头和畸变参数'],
])
heading(doc, '3.2 结构和布局', 2)
para(doc, '电池和飞控布置在几何中心附近，降低转动惯量和重心偏移；香橙派安装在机体上层并留出散热空间；IMX219 朝向主要搜索方向且无遮挡；Micovio 按其视场和安装要求固定。电机和电调线束远离相机视场和桨叶，飞控采用软安装但不得产生明显摆动。所有设备安装完成后重新测量重心，不用名义位置代替实际位置。')
heading(doc, '3.3 动力匹配', 2)
para(doc, '悬停时单电机平均推力按 T_h = mg/4 估算，总推重比按 R = 4T_max/(mg) 计算。由于当前尚未给出 1103 电机 KV、桨叶和电池的完整参数，报告不虚构推力和续航结论。正式参赛前需使用实际电机、电调、桨叶和电池完成静态推力测试，记录单电机推力、电流、温升和电池压降，并确认全机推重比满足自主飞行和避障机动需要。')

heading(doc, '四、感知与识别方案', 1)
heading(doc, '4.1 目标检测', 2)
para(doc, '目标检测采用 YOLO。香橙派采集 IMX219 图像后完成缩放、颜色空间转换和推理，输出目标类别、置信度、边界框中心和目标区域。比赛目标可能在颜色、形状或编号上变化，因此模型训练和测试应覆盖不同背景、光照、距离、角度和遮挡。任务状态机使用连续多帧确认，不采用单帧检测结果直接触发动作。')
heading(doc, '4.2 二维码识别', 2)
para(doc, '二维码识别采用 OpenCV。无人车二维码是降落指令来源，方形盒子二维码只用于目标确认，二者在任务逻辑中严格区分。识别流程为：进入二维码搜索区、稳定悬停、提取图像、定位二维码、解码文本、校验内容、连续帧确认、生成降落区指令。二维码文本设置格式校验和白名单，只有连续多帧一致且定位质量满足要求时才允许进入降落阶段。')
heading(doc, '4.3 障碍物识别', 2)
para(doc, '比赛障碍物包括柱子和方形框。视觉模块通过 YOLO 检测已知障碍物，结合边界框大小、图像位置、连续帧稳定性和安全距离估计输出风险等级。对未确认的物体采用保守策略：降低速度或悬停，不把低置信度结果当作“无障碍”。在方形框场景中，根据框的中心和尺寸判断可通过区域，只有通道宽度和定位误差均满足安全余量时才允许穿越。')
heading(doc, '4.4 Micovio 室内定位', 2)
para(doc, 'Micovio 负责无 GNSS 环境下的室内定位。接入 PX4 前需要确定其输出的是位置、速度、姿态还是视觉里程计，并确认坐标系、时间戳、频率、数据质量和异常标志。定位数据进入飞控前进行轴向转换、时间对齐和超时检查。定位质量下降时，系统停止前进并悬停；持续丢失则由 PX4 和任务状态机执行降落或其他预案。')

heading(doc, '五、定位导航与控制策略', 1)
heading(doc, '5.1 导航方式', 2)
para(doc, '比赛场地尺寸有限，系统采用“任务航点 + 视觉修正 + 安全速度限制”的导航策略。起飞后先升至安全高度并确认定位稳定，再前往目标搜索区域。到达区域后使用低速运动和定点悬停完成目标识别；识别无人车二维码后，将二维码指令映射到指定降落区航点，最后采用低速垂直下降完成降落。')
heading(doc, '5.2 避障策略', 2)
table(doc, ['状态','触发条件','控制动作','恢复条件'], [
    ['正常导航','前方风险低、定位有效','按航点和速度上限飞行','持续满足安全条件'],
    ['减速接近','障碍物进入预警距离','降低速度，增加图像确认频率','障碍物远离或路径确认'],
    ['悬停等待','风险高、识别不确定或定位异常','位置保持，停止前进','风险解除或进入降落预案'],
    ['绕行/改道','障碍物阻断当前路径','选择预设侧向航点或重新规划','到达安全通道'],
    ['安全降落','持续定位失败、电量低或严重故障','取消任务，选择安全区域下降','触地并上锁'],
])
heading(doc, '5.3 航迹跟踪', 2)
para(doc, 'PX4 负责底层位置和姿态闭环，香橙派只发送经过限幅和状态检查的高层位置/速度指令。指令包含目标点、速度上限和有效时间；通信超时后不继续使用旧指令。飞行过程中记录定位误差、航点到达时间、速度、姿态和电机输出，用于调节控制参数和判断是否存在振动或动力不足。')

heading(doc, '六、比赛任务执行流程', 1)
heading(doc, '6.1 任务状态机', 2)
table(doc, ['状态','进入条件','主要动作','异常处理'], [
    ['赛前检查','裁判允许准备','检查机体、供电、定位、通信、相机、急停和电池','任一项失败则不上场'],
    ['自主起飞','收到起飞指令且 PX4 允许解锁','解锁、升至安全高度、稳定悬停','起飞异常则悬停/降落'],
    ['避障导航','定位有效、飞行状态稳定','沿预设航线绕开柱子和方形框','风险高则减速、悬停或改道'],
    ['目标搜索','进入目标区域','运行 YOLO，确认指定目标类别/颜色/形状/编号','超时则按搜索策略重试'],
    ['无人车搜索','目标阶段完成或按规则进入','搜索移动无人车，保持安全距离','目标丢失则回到搜索状态'],
    ['二维码扫描','相机对准二维码且位置稳定','OpenCV 解码、校验降落指令','连续失败则调整视角或悬停'],
    ['指定区飞行','二维码指令有效','按指令导航至指定降落区','定位异常则中止降落'],
    ['自主降落','到达降落区、定位稳定','低速下降、检查投影位置、上锁','偏离过大则拉升重试或安全降落'],
])
heading(doc, '6.2 时间和任务管理', 2)
para(doc, '单轮时间上限为 5 分钟。任务程序为每个状态设置超时和重试次数，避免在单个目标或二维码上无限等待。搜索航线采用有限条带或矩形路径，目标一旦连续确认即停止搜索；二维码识别成功后不再重复搜索其他二维码，防止任务逻辑混乱。')
heading(doc, '6.3 自主模式和遥控模式', 2)
para(doc, '自主模式是主要参赛模式，任务中的起飞、避障、目标搜索、二维码识别、降落区选择和降落均由机载系统完成。遥控模式仅用于调试、风险兜底和必要的比赛策略；遥控器接管后该轮按遥控飞行模式计分，因此训练阶段必须尽量把问题暴露在自主模式中解决。')

heading(doc, '七、软件实现与系统集成', 1)
heading(doc, '7.1 软件模块', 2)
table(doc, ['模块','运行平台','功能'], [
    ['相机采集','香橙派','IMX219 驱动、帧时间戳、曝光和图像缓存'],
    ['目标检测','香橙派','YOLO 推理、连续帧确认、目标位置输出'],
    ['二维码识别','香橙派','OpenCV 定位、解码、内容校验和指令解析'],
    ['障碍物判断','香橙派','风险分级、安全距离和速度限制'],
    ['定位桥接','香橙派','Micovio 数据解析、坐标转换、时延和质量检查'],
    ['任务状态机','香橙派','任务推进、超时、重试和异常分支'],
    ['飞控控制','PX4','状态估计、位置姿态控制、混控和故障保护'],
])
heading(doc, '7.2 接口与数据记录', 2)
para(doc, '香橙派与 PX4 通过 MAVLink 或实际可用的串口/网络链路通信。所有外部数据必须包含时间戳和质量信息；桥接程序记录接收时间、发送时间、消息频率、丢包和延迟。PX4 日志保存姿态、位置、估计器状态、电池、电机输出、外部定位输入和故障状态；香橙派保存图像帧号、YOLO 结果、二维码文本、障碍物状态和任务状态转换。')
heading(doc, '7.3 现场部署', 2)
bullets(doc, ['赛前固定相机和 Micovio，禁止临场改变外参和坐标方向。','将模型、参数、任务航点、二维码白名单和安全阈值打包成版本化配置。','准备一套可恢复的 PX4 参数和香橙派启动脚本，现场按固定顺序启动。','指定一名接管人员持遥控器待命，其他队员负责计时、日志和赛场观察。'])

heading(doc, '八、实验验证与参赛准备', 1)
heading(doc, '8.1 验证顺序', 2)
table(doc, ['阶段','验证内容','通过标准'], [
    ['结构检查','质量、轴距、重心、桨叶间隙、紧固件','无松动、无干涉，参数可填写到比赛文档'],
    ['电气检查','供电、信号、PX4 输出、电调响应、急停','无异常复位，四电机顺序和旋向正确'],
    ['动力测试','静态推力、电流、温升、电压压降','满足悬停和机动裕量，温升可控'],
    ['定位测试','Micovio 坐标、频率、延迟、丢失恢复','静止无明显漂移，动作方向与机体一致'],
    ['识别测试','YOLO 目标、二维码、障碍物','记录准确率、漏检率、误检率和延迟'],
    ['控制测试','定高、悬停、航点、降落','无 GNSS 条件下可重复完成基础动作'],
    ['任务测试','按比赛顺序完成全流程','自主模式在 5 分钟内完成并可安全降落'],
])
heading(doc, '8.2 关键实测数据表', 2)
table(doc, ['参数','数值','测试条件/备注'], [
    ['整机质量','待测','含电池、桨叶、香橙派、Micovio 和相机'],
    ['轴距','待测','按比赛参数文档测量照片记录'],
    ['悬停油门和电流','待测','室内悬停，记录电池电压'],
    ['最大推力/推重比','待测','实际电机、电调、桨叶组合'],
    ['Micovio 定位频率和延迟','待测','记录原始输出和 PX4 接收时间'],
    ['YOLO 推理帧率/延迟','待测','香橙派 Zero 3W 实机测试'],
    ['二维码识别成功率','待测','不同距离、角度和光照'],
    ['避障最小安全距离','待测','柱子、方框分别测试'],
    ['完整任务用时','待测','自主模式，目标和二维码按赛场布置'],
])
heading(doc, '8.3 训练重点', 2)
bullets(doc, ['先完成稳定自主起飞和悬停，再加入目标搜索与二维码识别，不在基础飞行不稳定时调试复杂任务。','将移动无人车和静态方形盒子分别训练，明确二维码用途，避免把静态目标二维码误当作降落指令。','在 9 m × 5 m 场地中重复测试搜索航迹和避障航迹，尽量把单轮时间控制在 3 分钟以内，为现场误差和重试留出余量。','至少完成一次连续自主全流程，并保留 PX4 日志、香橙派日志和现场视频。'])

heading(doc, '九、安全与工程可行性', 1)
table(doc, ['风险','预防措施','现场处置'], [
    ['定位丢失或跳变','监测 Micovio 质量、超时和坐标一致性','停止前进，悬停或按预案降落'],
    ['视觉误检/漏检','连续帧确认、置信度阈值、内容校验','不触发降落，回到搜索或悬停'],
    ['链路或进程中断','心跳、超时、PX4 失联保护','由 PX4 执行安全模式，接管人员待命'],
    ['动力不足或过热','推力台测试、温升和电压监测','取消高机动任务，安全降落'],
    ['碰撞和失控','速度上限、障碍物预警、物理急停','立即接管，避免飞出防护网'],
])
para(doc, '本系统的工程可行性取决于三个闭环是否完成：第一是动力和结构闭环，必须知道真实质量、推重比和温升；第二是定位和控制闭环，必须确认 Micovio 数据能被 PX4 正确融合；第三是感知和任务闭环，必须在香橙派实际算力上达到足够的帧率和稳定性。报告中的硬件选型已经覆盖比赛要求，但最终参赛资格和现场表现还需要通过实测参数文档和自主飞行记录证明。')

heading(doc, '十、结论', 1)
para(doc, '本项目针对强国赛道无人机具身智能系统应用赛，提出了一套与比赛任务直接对应的微型四旋翼自主系统。系统以 PX4 负责实时飞控和安全保护，以香橙派 Zero 3W 负责高层智能计算，以 Micovio 实现室内定位，以 IMX219 采集视觉信息，采用 YOLO 完成目标检测、OpenCV 完成二维码识别，并以视觉识别结果驱动避障和任务状态机。')
para(doc, '与比赛评分项对应后，系统的技术主线清晰：定位为导航提供状态，YOLO 为目标搜索提供结果，OpenCV 为无人车降落指令提供依据，视觉障碍物判断为路径和速度提供约束，PX4 将高层任务转化为安全可执行的飞行动作。下一步应优先完成动力测试、Micovio-PX4 接入、目标与二维码数据集、障碍物场地测试和自主全流程训练，并将实测数据补入比赛要求的《无人机参数文档》。')

heading(doc, '附录 比赛参数文档填写清单', 1)
table(doc, ['项目','填写内容','当前状态'], [
    ['整机参数','重量、轴距、外形尺寸、实测照片','待补'],
    ['控制系统','PX4 控制器型号、固件版本','待补'],
    ['计算平台','香橙派 Zero 3W、系统和程序版本','已确定平台，版本待补'],
    ['感知系统','Micovio、IMX219、安装位置和图片','已确定器件，参数待补'],
    ['动力系统','电池型号、1103 电机型号、海科电调型号、桨叶','品牌/规格部分待补'],
    ['安全装置','遥控器、急停方式、接管人员','待补现场信息'],
    ['测试证明','质量、轴距、关键型号和性能测量照片','待测试和整理'],
])

doc.save(OUT)
print(OUT)
