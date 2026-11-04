from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = '/home/hzq/RM/px4-autopilot/'

def set_font(r, size=10.5, bold=False):
    r.font.name='宋体'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'宋体')
    r.font.size=Pt(size); r.bold=bold; r.font.color.rgb=RGBColor(0,0,0)

def border(c):
    tcPr=c._tc.get_or_add_tcPr(); bs=tcPr.first_child_found_in('w:tcBorders')
    if bs is None: bs=OxmlElement('w:tcBorders'); tcPr.append(bs)
    for e in ('top','left','bottom','right','insideH','insideV'):
        x=bs.find(qn('w:'+e))
        if x is None: x=OxmlElement('w:'+e); bs.append(x)
        x.set(qn('w:val'),'single'); x.set(qn('w:sz'),'4'); x.set(qn('w:color'),'000000')

def cell_margins(c):
    tcPr=c._tc.get_or_add_tcPr(); m=tcPr.first_child_found_in('w:tcMar')
    if m is None: m=OxmlElement('w:tcMar'); tcPr.append(m)
    for s in ('top','start','bottom','end'):
        x=OxmlElement('w:'+s); x.set(qn('w:w'),'100'); x.set(qn('w:type'),'dxa'); m.append(x)

def setup(d):
    s=d.sections[0]; s.top_margin=Inches(.7); s.bottom_margin=Inches(.65); s.left_margin=Inches(.85); s.right_margin=Inches(.75)
    n=d.styles['Normal']; n.font.name='宋体'; n._element.rPr.rFonts.set(qn('w:eastAsia'),'宋体'); n.font.size=Pt(10.5); n.paragraph_format.line_spacing=1.25; n.paragraph_format.space_after=Pt(4)
    for i,sz in [(1,16),(2,12.5)]:
        st=d.styles['Heading %d'%i]; st.font.name='黑体'; st._element.rPr.rFonts.set(qn('w:eastAsia'),'黑体'); st.font.size=Pt(sz); st.font.bold=True; st.font.color.rgb=RGBColor(0,0,0); st.paragraph_format.space_before=Pt(8 if i==1 else 5); st.paragraph_format.space_after=Pt(3)

def p(d,text,indent=True,center=False):
    x=d.add_paragraph(); x.paragraph_format.line_spacing=1.25; x.paragraph_format.space_after=Pt(4)
    if indent: x.paragraph_format.first_line_indent=Pt(21)
    if center: x.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=x.add_run(text); set_font(r); return x

def h(d,text,l=1):
    x=d.add_paragraph(style='Heading %d'%l); r=x.add_run(text); set_font(r,16 if l==1 else 12.5,True); return x

def bullets(d,items):
    for t in items:
        x=d.add_paragraph(style='List Bullet'); x.paragraph_format.left_indent=Pt(21); x.paragraph_format.space_after=Pt(2); r=x.add_run(t); set_font(r)

def table(d,heads,rows):
    t=d.add_table(rows=1,cols=len(heads)); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,v in enumerate(heads):
        c=t.rows[0].cells[i]; c.text=''; border(c); cell_margins(c); q=c.paragraphs[0]; q.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=q.add_run(v); set_font(r,9.3,True)
    for row in rows:
        cs=t.add_row().cells
        for i,v in enumerate(row):
            c=cs[i]; c.text=str(v); border(c); cell_margins(c); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for q in c.paragraphs:
                q.paragraph_format.space_after=Pt(0); q.paragraph_format.line_spacing=1.05
                for r in q.runs: set_font(r,9.2)
    d.add_paragraph().paragraph_format.space_after=Pt(0)

def save_report():
    d=Document(); setup(d)
    p(d,'无人机具身智能系统项目实现报告',False,True).runs[0].font.size=Pt(20)
    p(d,'参赛方向：强国赛道无人机具身智能系统应用赛',False,True)
    p(d,'系统平台：1103 无刷电机四旋翼、PX4、香橙派 Zero 3W、Micovio、IMX219',False,True)
    h(d,'一、无人机平台总体设计')
    p(d,'本项目采用小型四旋翼无人机作为具身智能平台。无人机通过自身传感器获取室内环境信息，利用机载计算机完成目标识别、二维码识别和障碍物判断，再由 PX4 飞控执行起飞、悬停、导航、避障和降落。系统不依赖 GNSS，适用于比赛规定的室内封闭场地。')
    table(d,['系统','选型及作用'],[
        ['机体结构','四旋翼 X 形布局；电池和飞控靠近重心，视觉传感器朝向任务区域。'],
        ['电控系统','海科电调驱动 4 台 1103 无刷电机，PX4 输出四路电机控制信号。'],
        ['飞控系统','PX4 开源飞控，负责姿态稳定、位置控制、电机混控和失控保护。'],
        ['机载计算','香橙派 Zero 3W，运行图像处理、YOLO、OpenCV 和任务状态机。'],
        ['传感器','Micovio 提供室内定位；IMX219 提供目标、二维码和障碍物图像。'],
        ['动力系统','1103 电机、海科电调、寸级桨和锂电池组成电驱动力系统。'],
    ])
    h(d,'二、感知与识别方案')
    h(d,'2.1 目标检测',2); p(d,'使用 YOLO 对 IMX219 图像进行目标检测，输出目标类别、置信度和目标位置。为降低误检，采用连续多帧确认，只有目标类别和位置连续稳定时才进入下一任务。')
    h(d,'2.2 二维码识别',2); p(d,'使用 OpenCV 完成二维码定位和解码。识别结果包括二维码文本、角点和中心位置。二维码内容经过格式校验后转换为降落区域指令；只有连续多帧识别一致时才允许执行降落。')
    h(d,'2.3 障碍物识别',2); p(d,'利用视觉识别柱子、方形框等障碍物。根据障碍物在图像中的位置、大小和连续帧变化判断风险。当障碍物距离过近或识别结果不确定时，飞行器减速或悬停，不继续前进。')
    h(d,'三、定位与导航方案')
    p(d,'Micovio 为无人机提供室内定位或视觉里程计信息。定位数据经过坐标系转换、时间同步和质量检查后送入 PX4，作为位置控制的输入。香橙派根据比赛场地和目标位置生成搜索航线及降落航点，PX4 负责航点跟踪。')
    bullets(d,['起飞后上升至安全高度并确认定位稳定。','沿预设航线避开柱子和方形框，进入目标区域。','使用 YOLO 搜索并确认指定目标。','搜索无人车并使用 OpenCV 识别二维码。','解析二维码中的降落指令，飞往指定降落区。','在降落区悬停、减速并完成安全降落。'])
    h(d,'四、任务执行流程')
    table(d,['阶段','执行内容','异常处理'],[
        ['系统检查','检查电池、PX4、Micovio、相机、通信和急停。','检查失败不得起飞。'],
        ['自主起飞','解锁并上升至安全高度，进入稳定飞行。','起飞异常则悬停或降落。'],
        ['避障导航','按预设航线飞往目标区域。','障碍物风险高时减速、悬停或绕行。'],
        ['目标搜索','YOLO 识别指定目标并确认。','目标丢失则重新搜索。'],
        ['二维码扫描','搜索无人车，悬停并用 OpenCV 解码。','识别失败则调整位置和视角。'],
        ['降落','按二维码指令飞往指定区域并降落。','定位异常时取消降落并悬停。'],
    ])
    h(d,'五、工程实现与安全措施')
    bullets(d,['PX4 负责实时控制，香橙派只发送经过限幅和状态检查的高层指令。','Micovio 数据超时或质量下降时，停止前进并悬停。','视觉图像丢失或识别不确定时，不执行降落动作。','保留遥控器和紧急安全开关，由指定队员负责接管。','完成动力、定位、识别、避障和完整任务测试后再参加正式比赛。'])
    h(d,'六、验证内容')
    table(d,['验证项目','主要内容','结果'],[
        ['平台参数','整机质量、轴距、重心和外形尺寸。','待测'],['动力系统','最大推力、悬停电流、电机/电调温升。','待测'],['室内定位','Micovio 坐标方向、频率、延迟和丢失恢复。','待测'],['目标检测','YOLO 识别率、误检率和推理延迟。','待测'],['二维码识别','不同距离、角度和光照下的识别成功率。','待测'],['完整任务','自主起飞、避障、目标识别、扫码和降落。','待测'],
    ])
    h(d,'七、结论')
    p(d,'本项目完成了面向无人机具身智能系统应用赛的总体设计。系统硬件与比赛要求相匹配，能够围绕自主起飞、避障导航、目标检测、二维码识别和指定区域降落建立完整技术链路。后续重点是补齐实际型号、实测参数和完整自主飞行数据。')
    d.save(ROOT+'无人机具身智能系统项目实现报告_简版.docx')

def save_params():
    d=Document(); setup(d)
    p(d,'无人机参数文档',False,True).runs[0].font.size=Pt(20)
    p(d,'强国赛道无人机具身智能系统应用赛',False,True)
    h(d,'一、整机参数')
    table(d,['参数','内容'],[
        ['无人机类型','小型四旋翼无人机'],['构型','X 形布局'],['整机重量','待测：______ g'],['轴距','待测：______ mm'],['外形尺寸','待测：______ mm × ______ mm × ______ mm'],['飞行高度','比赛常规飞行高度不低于 1.2 m，具体按现场要求执行'],
    ])
    h(d,'二、控制系统')
    table(d,['项目','内容'],[
        ['飞控系统','PX4 开源飞控'],['飞控控制器具体型号','待补：________________'],['固件版本','待补：________________'],['控制功能','姿态控制、位置控制、航点跟踪、失控保护'],['遥控与急停','配备遥控器；指定队员负责紧急接管'],
    ])
    h(d,'三、计算平台')
    table(d,['项目','内容'],[
        ['板载计算机','香橙派 Zero 3W'],['系统版本','待补：________________'],['运行程序','YOLO 目标检测、OpenCV 二维码识别、视觉障碍物识别、任务状态机'],['通信方式','待补：串口 / USB / 网络：________________'],
    ])
    h(d,'四、感知系统')
    table(d,['传感器','型号/用途','备注'],[
        ['室内定位','微空科技 Micovio','输出室内定位或视觉里程计；接口、频率和坐标系待补'],['摄像头','IMX219','用于目标检测、二维码识别和障碍物识别'],['飞控内置传感器','随 PX4 飞控配置','IMU、气压计等具体型号待补'],
    ])
    h(d,'五、动力系统')
    table(d,['项目','型号/参数'],[
        ['电池型号','待补：________________'],['电池电压/容量','待补：________________'],['电机型号','1103 无刷电机，具体 KV 待补：________________'],['电机数量','4 台'],['电调型号','海科电调，具体型号待补：________________'],['螺旋桨','寸级桨，具体直径、螺距和叶数待补：________________'],['动力测试','最大推力：______；悬停电流：______；温升：______'],
    ])
    h(d,'六、提交前检查')
    bullets(d,['整机重量和轴距已经实测，并附测量照片。','PX4 飞控控制器型号和固件版本已经记录。','香橙派 Zero 3W 的系统和程序版本已经记录。','Micovio、IMX219、1103 电机和海科电调的产品型号信息已经整理。','电池、电机和电调的实物照片或型号标签已经整理。','遥控器和紧急安全开关已经完成测试。'])
    d.save(ROOT+'无人机参数文档_简版.docx')

save_report(); save_params()
print('done')
