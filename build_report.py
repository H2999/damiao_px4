from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.enum.section import WD_ORIENT

OUT = '/home/hzq/RM/px4-autopilot/无人机结构设计与验证项目实验报告_简洁版.docx'

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), fill)

def set_cell_border(cell, color='000000', size='4'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        tag = 'w:' + edge
        el = borders.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            borders.append(el)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), size)
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)

def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in('w:tcMar')
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)
    for m, v in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        node = tcMar.find(qn('w:' + m))
        if node is None:
            node = OxmlElement('w:' + m)
            tcMar.append(node)
        node.set(qn('w:w'), str(v))
        node.set(qn('w:type'), 'dxa')

def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), 'true')
    trPr.append(tblHeader)

def set_keep_with_next(p):
    pPr = p._p.get_or_add_pPr()
    keep = OxmlElement('w:keepNext')
    pPr.append(keep)

def set_font(run, name='宋体', size=10.5, bold=False, color='000000'):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)

def add_p(doc, text='', style=None, bold_prefix=None, align=None, first_line=True):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    if bold_prefix and text.startswith(bold_prefix):
        r = p.add_run(bold_prefix); set_font(r, bold=True)
        r = p.add_run(text[len(bold_prefix):]); set_font(r)
    else:
        r = p.add_run(text); set_font(r)
    if style is None and first_line:
        p.paragraph_format.first_line_indent = Pt(21)
    return p

def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style='Heading %d' % level)
    r = p.add_run(text); set_font(r, size=16 if level==1 else 13 if level==2 else 11.5, bold=True)
    set_keep_with_next(p)
    return p

def add_table(doc, headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    for i, h in enumerate(headers):
        c = hdr.cells[i]; c.text = ''
        set_cell_shading(c, 'FFFFFF'); set_cell_border(c); set_cell_margins(c)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        r = c.paragraphs[0].add_run(str(h)); set_font(r, size=9.5, bold=True, color='000000')
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if widths: c.width = Inches(widths[i])
    for ridx, row in enumerate(rows):
        cells = table.add_row().cells
        for i, val in enumerate(row):
            c = cells[i]; c.text = str(val)
            set_cell_border(c); set_cell_margins(c)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.05
                for r in p.runs: set_font(r, size=9.2)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table

def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style='List Bullet')
        r = p.add_run(item); set_font(r)
        p.paragraph_format.left_indent = Pt(21)

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.75); sec.bottom_margin = Inches(0.65)
sec.left_margin = Inches(0.85); sec.right_margin = Inches(0.75)
sec.header_distance = Inches(0.3); sec.footer_distance = Inches(0.35)

styles = doc.styles
styles['Normal'].font.name = '宋体'; styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体'); styles['Normal'].font.size = Pt(10.5)
styles['Normal'].paragraph_format.line_spacing = 1.25; styles['Normal'].paragraph_format.space_after = Pt(4)
for level, size in [(1,16),(2,13),(3,11.5)]:
    st = styles['Heading %d' % level]
    st.font.name = '黑体'; st._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); st.font.size = Pt(size); st.font.bold = True; st.font.color.rgb = RGBColor(0,0,0)
    st.paragraph_format.space_before = Pt(9 if level==1 else 6); st.paragraph_format.space_after = Pt(3)
styles['Title'].font.name = '黑体'; styles['Title']._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); styles['Title'].font.size = Pt(22); styles['Title'].font.bold = True; styles['Title'].font.color.rgb = RGBColor(0,0,0)
styles['Subtitle'].font.name = '宋体'; styles['Subtitle']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体'); styles['Subtitle'].font.size = Pt(13); styles['Subtitle'].font.color.rgb = RGBColor(0,0,0)

header = sec.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
rh = header.add_run('无人机结构设计与验证项目实验报告'); set_font(rh, size=8.5, color='666666')
footer = sec.footer.paragraphs[0]; footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
rf = footer.add_run('PX4 室内视觉导航四旋翼验证样机'); set_font(rf, size=8.5, color='666666')

p = doc.add_paragraph(style='Title'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(55)
r = p.add_run('项目实验报告'); set_font(r, name='黑体', size=22, bold=True)
p = doc.add_paragraph(style='Subtitle'); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(10)
r = p.add_run('无人机结构设计与验证'); set_font(r, name='宋体', size=13)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(38)
r = p.add_run('小型四旋翼无人机电控与室内视觉导航系统设计'); set_font(r, size=16, bold=True)
for text in ['实验对象：1103 无刷电机与海科电调四旋翼验证样机', '飞控系统：PX4 开源飞控', '机载计算平台：香橙派 Zero 3W', '报告人：____________    日期：____________']:
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(12)
    r = p.add_run(text); set_font(r, size=11)
doc.add_page_break()

add_heading(doc, '摘要', 1)
add_p(doc, '本项目面向室内环境下的小型四旋翼自主任务，完成机体结构、动力系统、电控系统、机载计算、视觉感知、室内定位与任务流程的一体化设计。样机采用海科电调驱动 1103 无刷电机，配套寸级小桨；飞控采用 PX4 开源飞控，机载计算采用香橙派 Zero 3W，视觉传感器由微空科技 Micovio 模块与 IMX219 摄像头组成。系统以 Micovio 提供的室内定位信息支撑位置闭环，以视觉实现目标检测、二维码识别和障碍物识别，并通过 PX4 完成姿态、位置和任务状态控制。')
add_p(doc, '报告给出从起飞、目标搜索、二维码扫描、障碍物绕行、降落区选择到降落与任务结束的完整流程，建立了动力、结构、通信和安全方面的验证方法。需要特别说明的是，电调具体型号、1103 电机 KV 值、桨叶直径、机体最终质量、Micovio 输出协议及现场测试数据尚未全部提供，因此本文将这部分内容区分为设计方案、理论估算和待实测项目；本文不把未完成的测试写成已验证结论。')
add_p(doc, '关键词：四旋翼；PX4；香橙派 Zero 3W；Micovio；IMX219；YOLO；OpenCV；室内定位；视觉避障')

add_heading(doc, '目录', 1)
for item in ['一、项目概述','二、系统总体方案','三、无人机平台与结构设计','四、电控与动力系统设计','五、感知识别与室内定位方案','六、飞控软件与任务执行流程','七、工程实现与系统集成','八、实验验证方案与结果记录','九、安全设计与风险分析','十、结论与改进方向','附录 A 关键参数表','附录 B 联调检查表']:
    p = doc.add_paragraph(item); p.paragraph_format.left_indent = Pt(21); p.paragraph_format.space_after = Pt(2)
    for r in p.runs: set_font(r)
doc.add_page_break()

add_heading(doc, '一、项目概述', 1)
add_heading(doc, '1.1 项目背景', 2)
add_p(doc, '室内无人机通常缺少 GNSS 信号，且空间狭小、障碍物密集、光照变化明显，对定位连续性、飞行稳定性和感知闭环提出了更高要求。本项目以一台小型四旋翼验证样机为对象，将开源飞控、视觉定位、目标检测、二维码识别和视觉避障集成为一个可重复测试的自主飞行系统。')
add_heading(doc, '1.2 项目目标', 2)
add_bullets(doc, ['完成机体结构和动力系统的总体设计，保证部件安装、重心布置和推重比满足室内试飞要求。','使用 PX4 实现姿态稳定、位置控制、起飞、航点飞行和降落等基础飞行功能。','利用 Micovio 提供室内定位信息，使飞行器在无 GNSS 环境下获得位置估计。','利用 IMX219 摄像头采集图像，使用 YOLO 完成目标检测，使用 OpenCV 完成二维码识别，并通过视觉处理结果实现障碍物识别。','形成从起飞到任务结束的完整流程，以及可执行的安全、验证和故障处置方案。'])
add_heading(doc, '1.3 需求边界与数据说明', 2)
add_p(doc, '用户已确定的硬件和算法选型作为本报告的设计输入。海科电调的具体型号、协议和持续电流能力，1103 电机的 KV 值与额定电流，桨叶的直径和螺距，Micovio 的具体输出接口与坐标系，均需在装机后依据实物标签、数据手册和实测结果补录。报告中凡标注“待实测”的项目，不作为已完成验收的依据。')

add_heading(doc, '二、系统总体方案', 1)
add_heading(doc, '2.1 系统架构', 2)
add_p(doc, '系统采用“PX4 飞控负责实时飞行安全，香橙派负责高层视觉计算与任务决策”的分层架构。PX4 运行在飞控控制器上，直接连接 IMU、气压计、遥控急停链路和四路电调；香橙派 Zero 3W 连接 IMX219 摄像头与 Micovio 模块，运行图像处理、目标检测、二维码识别和任务状态机，并通过 MAVLink 向 PX4 发送位置、速度或航点级控制指令。')
add_table(doc, ['层级','核心部件','主要职责','接口与数据'], [
    ['执行层','1103 电机×4、海科电调×4、螺旋桨','产生推力并执行飞控输出','PWM/DShot 等协议以电调实际支持为准'],
    ['飞控层','PX4 开源飞控','姿态估计、姿态控制、位置控制、失控保护','MAVLink、传感器总线、遥控急停'],
    ['计算层','香橙派 Zero 3W','视觉推理、二维码识别、障碍物判断、任务状态机','摄像头接口、串口/USB/网络链路'],
    ['感知层','Micovio、IMX219','室内定位、图像采集','Micovio 输出协议和相机驱动待确认'],
])
add_heading(doc, '2.2 信息流与控制闭环', 2)
add_p(doc, 'IMX219 输出图像帧，香橙派完成图像预处理后分别送入 YOLO 和 OpenCV QRCodeDetector（或同等 OpenCV 解码接口）。Micovio 输出位置、姿态或视觉里程计信息，经坐标系转换、时间同步和质量检查后注入 PX4 的估计器。PX4 将估计状态与目标位置比较，经过位置控制、姿态控制和混控器计算后输出四路电调控制量。任务状态机只在定位质量、链路状态和电池状态满足条件时推进。')
add_heading(doc, '2.3 设计指标', 2)
add_table(doc, ['指标类别','设计目标','验证方式','状态'], [
    ['基础飞行','稳定完成起飞、悬停、航点飞行和降落','室内无载与带载飞行记录','待实测'],
    ['室内定位','定位输出连续、坐标系一致、无明显跳变','记录 PX4 日志和 Micovio 原始数据','待实测'],
    ['目标检测','在任务距离和光照条件下检测目标并给出置信度','标定数据集与现场视频回放','待实测'],
    ['二维码识别','识别内容、位置和朝向可供任务状态机使用','不同距离、角度和光照测试','待实测'],
    ['避障','检测前方障碍并触发悬停、绕行或返航策略','布置静态障碍物进行闭环测试','待实测'],
    ['安全','失联、定位丢失、低电量时进入预定安全状态','故障注入测试','待实测'],
])

add_heading(doc, '三、无人机平台与结构设计', 1)
add_heading(doc, '3.1 构型与布局', 2)
add_p(doc, '样机采用四旋翼 X 形布局。四个电机位于机臂端部，电池靠近几何中心布置，PX4 飞控安装在重心附近并采用软安装，香橙派与相机布置在机体上层或前向安装位，Micovio 视其视场和安装要求布置在无遮挡位置。结构设计遵循传力路径短、重心集中、线束不干涉桨面和传感器视场的原则。')
add_table(doc, ['部件','设计选型','结构与安装要求'], [
    ['机架','轻量化碳纤维板或等效轻质结构','机臂根部圆角过渡，电机孔周围留足边距'],
    ['电机座','与机臂一体化或刚性连接','保证电机轴线垂直，避免安装面倾斜'],
    ['飞控','PX4 飞控控制器','靠近重心、箭头方向明确、软安装并标定安装旋转'],
    ['机载计算','香橙派 Zero 3W','防振固定，留出散热和供电滤波空间'],
    ['视觉组件','Micovio + IMX219','相机无遮挡，固定焦距，记录外参和坐标系方向'],
    ['防护与起落架','按任务空间选配','不能遮挡相机，也不能与桨叶形成危险间隙'],
])
add_heading(doc, '3.2 结构受力与振动设计', 2)
add_p(doc, '主要载荷路径为“桨叶—电机—电机座—机臂—中心板—电池和设备安装点”。结构分析至少覆盖自重、单电机推力、推力不均匀、着陆冲击和线束/设备安装载荷。电机与机臂连接处、板件开孔处、机臂根部和起落架连接处是应力集中位置。')
add_p(doc, '振动控制采用桨叶动平衡、电机安装面检查、飞控软安装和线束固定相结合的方法。模态分析应以实际质量、紧固件、相机和机载计算平台为输入，先进行自由模态，再根据安装状态补充约束模态。结果应与电机转频及桨叶通过频率对照，避免主要结构模态落入长期运行激励频带。')

add_heading(doc, '四、电控与动力系统设计', 1)
add_heading(doc, '4.1 动力系统组成', 2)
add_p(doc, '动力系统由 1103 无刷电机、海科电调、寸级小桨和电池组成。电调接收 PX4 的四路电机控制信号，完成三相换相和转速执行。最终电调型号、额定电流、输入电压范围、信号协议、是否支持遥测以及电机 KV 值必须以实物和数据手册为准。')
add_table(doc, ['项目','已确定内容','需补充/实测内容'], [
    ['电调','海科品牌电调，四路使用','具体型号、持续/峰值电流、输入电压、信号协议、遥测能力'],
    ['电机','1103 无刷电机，4 台','KV 值、额定电流、最大功率、安装孔距、单机质量'],
    ['螺旋桨','寸级小桨，与 1103 配套','直径、螺距、叶数、正反桨型号、静态推力曲线'],
    ['电源','根据电机和电调额定范围选配','电池节数、容量、放电倍率、低压阈值、峰值电流'],
    ['控制信号','PX4 输出至四路电调','PWM/DShot 具体配置、输出频率、失效状态'],
])
add_heading(doc, '4.2 动力匹配计算', 2)
add_p(doc, '悬停时单电机平均推力可按 T_h = mg/4 估算；推重比可按 R = 4T_max/(mg) 计算。室内自主飞行建议设计推重比不低于 2，并保留姿态机动、抗扰和电压下降裕量。由于用户尚未提供 1103 电机 KV、桨型和电池参数，本文不直接沿用原稿中针对其他桨叶的转速和推力数值；正式结论应以推力台数据或可靠厂商曲线为准。')
add_p(doc, '动力验证至少记录悬停电流、最大电流、单电机静态推力、全机最大推力、ESC 温升、电机温升和电池压降。若满油门推重比不足、悬停油门过高或温升超限，应优先调整桨叶、电池和整机质量，而不是通过提高控制增益弥补动力不足。')
add_heading(doc, '4.3 供电与电磁兼容', 2)
add_p(doc, '动力电源与飞控、香橙派和视觉传感器的供电应分级处理。飞控电源应具备稳定的稳压和滤波，香橙派供电应满足启动和视觉推理峰值电流，Micovio 与相机电源应避免电机换相噪声耦合。所有电源线、信号线和螺旋桨运动区域分离布置，电池主回路设置保险或可控断电措施，地线采用明确的公共参考点并减少大电流回路面积。')

add_heading(doc, '五、感知识别与室内定位方案', 1)
add_heading(doc, '5.1 Micovio 室内定位', 2)
add_p(doc, 'Micovio 模块作为室内定位传感器，为飞行器提供无 GNSS 环境下的位置估计或视觉里程计信息。集成时需要确认输出内容、坐标系定义、频率、时间戳、质量指标和通信接口，并完成传感器安装外参标定。定位数据进入 PX4 前应进行坐标变换、时间对齐和异常值检查；当质量指标低于阈值或数据超时，应禁止继续执行高风险动作。')
add_heading(doc, '5.2 IMX219 图像采集', 2)
add_p(doc, 'IMX219 摄像头连接香橙派 Zero 3W，负责采集目标检测、二维码识别和障碍物分析所需图像。相机安装需固定焦距和姿态，记录分辨率、帧率、曝光、白平衡和畸变参数。实验中应覆盖室内自然光、人工光、逆光和低照度条件，并避免自动曝光变化造成识别阈值漂移。')
add_heading(doc, '5.3 YOLO 目标检测', 2)
add_p(doc, '目标检测采用 YOLO 算法。香橙派端执行图像缩放、颜色空间转换和推理，输出目标类别、置信度和边界框。任务逻辑不应只使用单帧结果，而应采用连续帧确认、置信度阈值和目标跟踪/去抖机制，避免由于瞬时误检导致错误降落或误触发动作。模型部署前应根据香橙派算力选择轻量化模型、输入尺寸和推理后端，并记录平均帧率、端到端延迟和 CPU/内存占用。')
add_heading(doc, '5.4 OpenCV 二维码识别', 2)
add_p(doc, '二维码识别使用 OpenCV 库。处理流程包括图像灰度化、必要的去噪与二值化、二维码定位、解码、内容校验和连续帧确认。识别结果至少包含二维码文本、四角点、中心点、识别时间戳和置信状态。对任务有意义的二维码内容应设置白名单或格式校验，避免把任意图案解码结果当作有效指令。')
add_heading(doc, '5.5 基于视觉的障碍物识别', 2)
add_p(doc, '障碍物识别利用 IMX219 图像完成。第一阶段可采用 YOLO 检测已知障碍物类别；对未知障碍物，可结合图像边缘、占用区域、连续帧运动一致性和安全距离规则进行保守判断。视觉障碍物检测存在遮挡、光照和纹理不足等限制，因此应设置“未知即风险”的策略：当检测置信度不足、图像模糊或定位不可靠时，优先悬停并等待，而不是继续前进。')

add_heading(doc, '六、飞控软件与任务执行流程', 1)
add_heading(doc, '6.1 PX4 飞控配置', 2)
add_p(doc, 'PX4 负责传感器融合、姿态控制、位置控制、混控和故障保护。首次装机应完成机架类型选择、传感器方向校准、水平校准、罗盘与气压计检查、电机顺序检查、遥控器急停检查和电池参数设置。外部定位数据接入后，应在无桨环境中先检查坐标轴、符号、时间戳和延迟，再进行低高度试飞。')
add_table(doc, ['功能','PX4 侧职责','香橙派侧职责','故障处置'], [
    ['姿态稳定','IMU 融合与姿态控制','不直接闭合姿态环','姿态异常则降落/急停'],
    ['室内位置','融合 Micovio 或外部视觉定位','提供带时间戳的定位数据','定位丢失则悬停，超时返航/降落'],
    ['目标搜索','执行位置/速度指令','YOLO 检测并输出目标相对信息','目标丢失则减速、悬停或重新搜索'],
    ['二维码扫描','保持稳定悬停或低速运动','OpenCV 解码与内容校验','连续失败则调整视角，不盲目降落'],
    ['避障','执行速度限制或位置保持','识别障碍并产生安全策略','障碍不确定时悬停'],
])
add_heading(doc, '6.2 完整任务流程', 2)
add_p(doc, '任务状态机建议采用以下顺序：系统自检 → 解锁前检查 → 起飞 → 定高与定位稳定 → 目标搜索 → 目标确认 → 二维码扫描 → 任务结果校验 → 选择降落区 → 避障接近 → 降落 → 上锁与结果记录。每个状态均应定义进入条件、执行动作、超时条件和退出状态，任何关键条件不满足都应进入悬停、返航或降落等安全分支。')
add_table(doc, ['状态','进入条件','核心动作','退出条件'], [
    ['自检','供电正常、链路正常','检查电池、传感器、Micovio、相机和任务参数','全部通过或进入故障'],
    ['起飞','PX4 允许解锁、定位有效','低速起飞至安全高度','高度和定位稳定'],
    ['目标搜索','飞行状态稳定','按搜索航线移动并运行 YOLO','目标连续确认或超时'],
    ['二维码扫描','目标区域已确认','悬停/低速对准二维码并用 OpenCV 解码','内容校验通过或重试上限'],
    ['避障接近','降落区有效且路径安全','按安全速度接近，持续检查视觉障碍物','进入降落判定区或触发风险'],
    ['降落','地面/降落区条件满足','垂直下降并监测位置、速度和障碍','触地、上锁或中止'],
])
add_heading(doc, '6.3 安全分支', 2)
add_bullets(doc, ['Micovio 数据超时或定位质量下降：立即停止前进，进入悬停；持续超时则按预设高度和区域执行降落或返航。','香橙派与 PX4 链路中断：由 PX4 执行通信失联保护，禁止依赖伴随计算机继续自主动作。','视觉图像丢失或严重模糊：停止目标追踪和降落决策，保持安全状态。','低电量：禁止进入新任务阶段，优先返航或降落；电池参数必须通过实际电池和电流计校准。','遥控急停：在无桨测试和低高度测试中确认急停优先级，急停触发时电机进入预定安全状态。'])

add_heading(doc, '七、工程实现与系统集成', 1)
add_heading(doc, '7.1 软件模块划分', 2)
add_table(doc, ['模块','运行位置','输入','输出'], [
    ['相机采集','香橙派','IMX219 视频流','带时间戳图像帧'],
    ['YOLO 检测','香橙派','图像帧','目标类别、框、置信度'],
    ['二维码识别','香橙派','ROI 或全图','二维码文本、角点、有效标志'],
    ['障碍物判断','香橙派','检测结果、连续图像','风险等级、速度限制/悬停指令'],
    ['定位桥接','香橙派','Micovio 输出','PX4 可融合的定位消息'],
    ['任务状态机','香橙派','识别、定位、电池、飞行状态','起飞/航点/悬停/降落指令'],
    ['飞控控制','PX4','传感器和外部数据','姿态与电机输出'],
])
add_heading(doc, '7.2 时间同步与坐标系', 2)
add_p(doc, '视觉和定位数据必须携带采样时间，桥接程序应记录接收时间、发送时间和估计延迟。相机坐标系、Micovio 坐标系、机体坐标系和 PX4 本地坐标系之间应通过明确的轴向约定和外参矩阵转换，禁止仅凭试飞现象修改轴符号。建议在静止、前移、右移、上升和偏航五种动作下分别记录坐标变化，形成坐标系验收表。')
add_heading(doc, '7.3 日志与可追溯性', 2)
add_p(doc, 'PX4 日志应保留姿态、位置、估计器状态、电池、电机输出、故障标志和外部视觉输入；香橙派应保留图像时间戳、YOLO 结果、二维码文本、障碍物状态、任务状态转换和链路延迟。每次试验使用唯一编号，并同时记录硬件版本、软件提交号、参数文件、场地、光照、载荷和异常现象。')

add_heading(doc, '八、实验验证方案与结果记录', 1)
add_heading(doc, '8.1 验证分级', 2)
add_p(doc, '验证遵循“单件—子系统—联调—低风险飞行—完整任务”的顺序。涉及螺旋桨旋转的试验必须完成结构、线束和急停检查；首次动力测试使用固定台架，首次自主飞行使用低高度、空旷区域和人工接管条件。')
add_table(doc, ['阶段','验证内容','主要记录','判定'], [
    ['V1 结构','尺寸、紧固件、重心、桨面间隙','质量、包络、螺钉、裂纹、松动','全部满足设计要求'],
    ['V2 电气','电源、信号、急停、温升','电压、电流、温度、信号协议','无异常复位和过热'],
    ['V3 动力','单电机推力和全机推重比','推力曲线、最大电流、油门','保留悬停与机动裕量'],
    ['V4 定位','Micovio 输出与 PX4 融合','轨迹、延迟、丢包、跳变','坐标一致且连续'],
    ['V5 感知','YOLO、二维码、障碍物','准确率、延迟、误检漏检','满足任务阈值'],
    ['V6 飞行','手动/定高/位置/任务模式','PX4 日志和视频','可安全起飞、悬停、降落'],
    ['V7 全流程','完整任务状态机','全过程日志与异常记录','任务闭环或明确失败原因'],
])
add_heading(doc, '8.2 结构与动力理论校核', 2)
add_p(doc, '结构理论校核使用实际整机质量和设备安装位置。静强度可先采用简化梁模型进行趋势判断，再用实际材料、孔位和紧固方式建立有限元模型。模态分析应输出前若干阶固有频率和振型，并与电机转频、桨叶通过频率以及可能的机体激励频带对照。')
add_p(doc, '动力校核的关键不是单一的额定推力，而是四台电机在实际电池电压、实际桨叶、实际电调和实际安装状态下的总推力。若无法取得厂商曲线，应进行静态推力台测试。只有在记录了整机质量、悬停油门、电流、温升和电池压降之后，才能形成最终动力匹配结论。')
add_heading(doc, '8.3 当前结果与待补数据', 2)
add_table(doc, ['项目','当前可得结论','后续必须补充'], [
    ['总体方案','硬件与算法链路完整，分层架构可实施','确认所有接口、电源和安装空间'],
    ['结构设计','已明确四旋翼布局、受力路径和减振原则','实测尺寸、质量、重心、模态和静强度'],
    ['动力系统','1103 + 海科电调 + 寸级小桨为设计方向','KV、桨型、推力、电流、温升和推重比'],
    ['室内定位','Micovio 可作为室内定位输入','协议、坐标系、频率、延迟、丢失策略'],
    ['视觉算法','YOLO、OpenCV 和视觉避障流程已定义','模型、数据集、准确率、延迟和误检统计'],
    ['完整任务','状态机与安全分支已定义','现场闭环飞行记录和异常复盘'],
])

add_heading(doc, '九、安全设计与风险分析', 1)
add_table(doc, ['风险','原因','预防措施','触发后的动作'], [
    ['桨叶伤害','无桨检查不足、误解锁','无桨先通电，设置物理急停和安全距离','急停、断电、人员撤离'],
    ['定位丢失','纹理不足、光照变化、模块遮挡','检查视场，监控质量和超时','悬停，按策略降落/返航'],
    ['视觉误检','训练数据不足、反光、模糊','连续帧确认、置信度和内容白名单','不执行高风险动作'],
    ['动力不足','桨型或电池不匹配、质量过大','推力台测试、温升和电压监测','限制任务或立即降落'],
    ['链路中断','串口、网络或进程异常','心跳、超时和 PX4 失联保护','交由 PX4 执行失联策略'],
    ['结构振动','桨叶不平衡、安装松动','动平衡、紧固和减振','降落检查并禁止继续高机动'],
])
add_heading(doc, '十、结论与改进方向', 1)
add_p(doc, '本项目完成了小型四旋翼无人机的系统级设计：动力系统采用海科电调和 1103 无刷电机配套寸级小桨，飞控采用 PX4，机载计算采用香橙派 Zero 3W，感知系统采用 Micovio 与 IMX219，相应的软件链路采用 YOLO 目标检测、OpenCV 二维码识别和基于视觉的障碍物识别。系统以 Micovio 室内定位支撑 PX4 位置控制，以香橙派任务状态机组织感知与动作，具备开展室内自主任务验证的架构基础。')
add_p(doc, '当前最重要的工程工作是完成具体器件参数和实测数据闭环。应优先完成电机—电调—桨叶推力测试、Micovio 与 PX4 坐标和时延验证、视觉算法性能测试以及低风险完整任务测试。只有这些数据补齐后，才能对推重比、续航、定位精度、识别率、避障距离和任务成功率给出定量结论。')
add_heading(doc, '改进方向', 2)
add_bullets(doc, ['补充海科电调和 1103 电机的具体型号、KV、额定电流和信号协议，形成可复现的动力参数表。','建立针对室内光照和目标距离的 YOLO 数据集，并在香橙派 Zero 3W 上完成模型量化或轻量化。','补充 Micovio 外参、坐标系和质量指标的自动检查，减少定位接入错误。','增加障碍物安全距离标定、未知障碍物保守策略和视觉失效降落策略。','完成结构有限元、振动和实际飞行日志的对照，形成设计值、仿真值和实测值三列数据。'])

add_heading(doc, '附录 A 关键参数表', 1)
add_table(doc, ['类别','参数','当前记录'], [
    ['平台','构型','四旋翼 X 形布局'],
    ['电机','型号/数量','1103 无刷电机，4 台；KV 待补'],
    ['电调','品牌/数量','海科电调，4 台；具体型号待补'],
    ['螺旋桨','类型','寸级小桨；直径、螺距和叶数待补'],
    ['飞控','系统','PX4 开源飞控'],
    ['机载计算','平台','香橙派 Zero 3W'],
    ['传感器','室内定位','微空科技 Micovio；输出协议待确认'],
    ['摄像头','图像采集','IMX219'],
    ['算法','目标检测','YOLO'],
    ['算法','二维码识别','OpenCV'],
    ['算法','障碍物识别','基于视觉的识别与安全策略'],
    ['通信','飞控链路','MAVLink/串口或网络链路，按实际接口确定'],
])

add_heading(doc, '附录 B 联调检查表', 1)
add_table(doc, ['序号','检查项','结果/备注'], [
    ['1','电机编号、旋向和 PX4 输出通道一致','□通过  □不通过'],
    ['2','螺旋桨型号、正反桨和安装方向正确','□通过  □不通过'],
    ['3','无桨通电、急停和失联保护已验证','□通过  □不通过'],
    ['4','飞控方向、水平、传感器和电池参数已校准','□通过  □不通过'],
    ['5','Micovio 坐标系、时间戳和质量指标已核对','□通过  □不通过'],
    ['6','IMX219 图像曝光、焦距和安装视场已确认','□通过  □不通过'],
    ['7','YOLO 模型推理帧率和延迟已记录','□通过  □不通过'],
    ['8','二维码识别内容白名单和连续帧确认已验证','□通过  □不通过'],
    ['9','障碍物识别触发悬停/绕行策略已验证','□通过  □不通过'],
    ['10','低高度手动接管和完整任务已完成记录','□通过  □不通过'],
])

doc.save(OUT)
print(OUT)
