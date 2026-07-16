from __future__ import annotations

from pathlib import Path
import hashlib
import shutil
import zipfile
import xml.etree.ElementTree as ET

from vsdx import VisioFile, Media, Connect

OUT = Path('output')
VSDX_DIR = OUT / 'Valid_Editable_Visio_Files'
OUT.mkdir(exist_ok=True)
VSDX_DIR.mkdir(parents=True, exist_ok=True)

PAGE_W = 13.333
PAGE_H = 7.5

PALETTE = [
    ('#D9EAF7', '#5B9BD5'),
    ('#E2F0D9', '#70AD47'),
    ('#FFF2CC', '#C9A227'),
    ('#FCE4D6', '#ED7D31'),
]

DIAGRAMS = [
    {
        'file': 'Fig01_Overall_Technical_System.vsdx',
        'title': '图1  项目总体技术体系与创新链',
        'layers': [
            [('need', '战略需求与行业痛点'), ('science', '关键科学技术问题'), ('target', '总体目标与技术指标')],
            [('flight', '多源感知与自主飞行'), ('recognition', '部件识别与缺陷诊断'), ('clearance', '树障空间量测与风险分级'), ('platform', '端—边—云协同平台')],
            [('product', '算法、终端与平台产品'), ('evidence', '检测、查新与应用证据'), ('benefit', '安全、经济与社会效益')],
        ],
        'edges': [('need','flight'),('need','recognition'),('science','recognition'),('science','clearance'),('target','platform'),('flight','product'),('recognition','product'),('clearance','evidence'),('platform','evidence'),('product','benefit'),('evidence','benefit')],
    },
    {
        'file': 'Fig02_Problem_Technology_Indicator_Mapping.vsdx',
        'title': '图2  科学问题—技术突破—工程指标映射',
        'layers': [
            [('p1','异构传感误差耦合与局部失效'),('p2','小目标、遮挡与少样本'),('p3','树障空间动态风险'),('p4','多系统割裂与模型漂移')],
            [('t1','深耦合定位与概率安全规划'),('t2','级联注意力、自监督与长尾学习'),('t3','图像—LiDAR量测与风险预测'),('t4','云边协同、MLOps与业务闭环')],
            [('i1','定位误差、航迹偏差、避障成功率'),('i2','Precision、Recall、F1、mAP'),('i3','净空MAE/RMSE与告警准确率'),('i4','时延、可用度、闭环处置率')],
        ],
        'edges': [('p1','t1'),('p2','t2'),('p3','t3'),('p4','t4'),('t1','i1'),('t2','i2'),('t3','i3'),('t4','i4')],
    },
    {
        'file': 'Fig03_Multisensor_Calibration_and_Fusion.vsdx',
        'title': '图3  多源传感器时空标定与数据融合',
        'layers': [
            [('cam','可见光/红外相机'),('lidar','激光雷达'),('imu','IMU'),('gnss','GNSS/RTK'),('fc','飞控与时间戳')],
            [('time','统一时钟与延迟补偿'),('extrinsic','相机—LiDAR外参标定'),('motion','运动畸变与位姿补偿'),('quality','数据质量评估与异常剔除')],
            [('frame','统一坐标系与时空基准'),('cloud','语义点云与三维环境模型'),('state','高频连续状态估计')],
        ],
        'edges': [('cam','time'),('lidar','time'),('imu','motion'),('gnss','motion'),('fc','time'),('cam','extrinsic'),('lidar','extrinsic'),('time','frame'),('extrinsic','cloud'),('motion','state'),('quality','cloud'),('quality','state')],
    },
    {
        'file': 'Fig04_GNSS_Denied_Autonomous_Flight.vsdx',
        'title': '图4  GNSS受限条件下自主定位与安全飞行闭环',
        'layers': [
            [('sense','点云、视觉、IMU、GNSS观测'),('mission','巡检任务与电子围栏'),('weather','风场、能量与通信状态')],
            [('fusion','滑窗优化/滤波融合定位'),('map','局部三维地图与障碍语义'),('planner','多约束轨迹规划与重规划'),('control','鲁棒控制与失联保护')],
            [('track','导线跟随与稳定成像'),('avoid','动态避障与安全距离保持'),('return','任务完成、返航与数据回传')],
        ],
        'edges': [('sense','fusion'),('sense','map'),('mission','planner'),('weather','planner'),('weather','control'),('fusion','planner'),('map','planner'),('planner','control'),('control','track'),('control','avoid'),('track','return'),('avoid','return')],
    },
    {
        'file': 'Fig05_Perception_Aware_Path_Planning.vsdx',
        'title': '图5  感知质量约束的多目标航迹规划',
        'layers': [
            [('distance','路径长度与任务覆盖'),('risk','碰撞概率与安全裕度'),('energy','能耗与剩余航时'),('view','视角、尺度、清晰度与重叠率')],
            [('objective','加权多目标代价函数'),('constraints','动力学/电子围栏/净空约束'),('solver','MPC、采样规划与在线重规划')],
            [('safe','安全可飞航迹'),('quality','可识别的高质量巡检数据'),('efficiency','任务效率与续航平衡')],
        ],
        'edges': [('distance','objective'),('risk','objective'),('energy','objective'),('view','objective'),('constraints','solver'),('objective','solver'),('solver','safe'),('solver','quality'),('solver','efficiency')],
    },
    {
        'file': 'Fig06_Cascaded_Attention_Recognition.vsdx',
        'title': '图6  关键部件级联注意力鲁棒识别架构',
        'layers': [
            [('image','多视角巡检影像'),('pretrain','自监督预训练表征'),('roi','杆塔/导线/部件候选区域')],
            [('backbone','多尺度卷积骨干'),('channel','通道注意力'),('spatial','空间注意力'),('cross','跨层特征吸收与语义对齐')],
            [('class','部件类别与置信度'),('box','精确定位与边界框'),('feature','可迁移部件特征库')],
        ],
        'edges': [('image','backbone'),('pretrain','backbone'),('roi','spatial'),('backbone','channel'),('channel','spatial'),('spatial','cross'),('cross','class'),('cross','box'),('cross','feature')],
    },
    {
        'file': 'Fig07_Data_Governance_Active_Learning.vsdx',
        'title': '图7  数据治理、主动学习与困难样本闭环',
        'layers': [
            [('collect','多区域/季节/电压等级采集'),('clean','脱敏、去重与质量筛查'),('label','自动预标注与人机复核'),('version','数据版本与谱系记录')],
            [('train','监督+自监督联合训练'),('uncertainty','不确定性与低置信样本挖掘'),('hard','遮挡、逆光、小目标困难样本生成'),('audit','训练日志、评测与偏差审计')],
            [('deploy','模型灰度发布'),('feedback','误检/漏检/漂移样本回流'),('iterate','数据—模型—场景持续迭代')],
        ],
        'edges': [('collect','clean'),('clean','label'),('label','version'),('version','train'),('train','uncertainty'),('uncertainty','hard'),('hard','audit'),('audit','deploy'),('deploy','feedback'),('feedback','iterate'),('iterate','train')],
    },
    {
        'file': 'Fig08_Convolution_Enhanced_Transformer.vsdx',
        'title': '图8  卷积增强Transformer典型缺陷检测网络',
        'layers': [
            [('input','高分辨率巡检图像'),('patch','卷积式Patch Embedding'),('local','局部窗口连接LWC'),('global','跨窗口连接GWC')],
            [('stage1','分层窗口注意力Stage 1'),('stage2','移位窗口与卷积增强Stage 2'),('stage3','全局语义聚合Stage 3'),('fpn','多尺度特征金字塔')],
            [('head','轻量分类/回归检测头'),('defect','裂纹、锈蚀、缺件等缺陷'),('tree','导线与树冠语义分割'),('edge','边缘端量化部署')],
        ],
        'edges': [('input','patch'),('patch','local'),('patch','global'),('local','stage1'),('global','stage2'),('stage1','stage2'),('stage2','stage3'),('stage3','fpn'),('fpn','head'),('head','defect'),('head','tree'),('head','edge')],
    },
    {
        'file': 'Fig09_Small_Target_Long_Tail_Training.vsdx',
        'title': '图9  多尺度小目标检测与长尾训练机制',
        'layers': [
            [('scale','尺度变化与小目标'),('imbalance','类别长尾与正负样本不均衡'),('blur','模糊、遮挡与低对比度'),('rare','稀有缺陷与开放集风险')],
            [('pyramid','高分辨率特征金字塔'),('focal','Focal/质量感知分类损失'),('iou','CIoU/GIoU边界回归'),('augment','Copy-Paste与生成式增广')],
            [('recall','提升小缺陷召回率'),('precision','抑制复杂背景误检'),('general','增强跨场景泛化'),('calibration','置信度校准与人工复核')],
        ],
        'edges': [('scale','pyramid'),('imbalance','focal'),('blur','iou'),('rare','augment'),('pyramid','recall'),('focal','precision'),('iou','general'),('augment','general'),('precision','calibration'),('general','calibration')],
    },
    {
        'file': 'Fig10_Image_LiDAR_Clearance_Measurement.vsdx',
        'title': '图10  图像—LiDAR跨模态树障空间量测',
        'layers': [
            [('semantic','图像导线/树冠语义分割'),('point','LiDAR点云去噪与地物分类'),('pose','时间同步、外参与位姿补偿')],
            [('project','语义投影与跨模态关联'),('wire','导线点集提取与曲线拟合'),('crown','树冠聚类与边界重建'),('uncert','量测误差与不确定度估计')],
            [('distance','最小净空距离'),('location','档距、杆塔与地理位置关联'),('alarm','阈值判定与分级告警')],
        ],
        'edges': [('semantic','project'),('point','project'),('pose','project'),('project','wire'),('project','crown'),('pose','uncert'),('wire','distance'),('crown','distance'),('uncert','distance'),('distance','location'),('distance','alarm'),('location','alarm')],
    },
    {
        'file': 'Fig11_Catenary_Wind_Vegetation_Risk.vsdx',
        'title': '图11  导线悬链线、风偏与树障动态风险预测',
        'layers': [
            [('span','档距、挂点与导线参数'),('weather','温度、风速与覆冰工况'),('growth','树种、生长率与树冠结构'),('measure','当前净空及量测不确定度')],
            [('catenary','悬链线与弧垂估计'),('wind','风偏包络与安全裕度'),('forecast','树冠增长/倒伏趋势预测'),('fusion','多因素风险融合模型')],
            [('risk','提示/一般/严重/紧急等级'),('priority','治理优先级与复查周期'),('route','近距复核航迹与工单生成')],
        ],
        'edges': [('span','catenary'),('weather','catenary'),('weather','wind'),('growth','forecast'),('measure','fusion'),('catenary','fusion'),('wind','fusion'),('forecast','fusion'),('fusion','risk'),('risk','priority'),('priority','route')],
    },
    {
        'file': 'Fig12_Edge_Cloud_MLOps_Closed_Loop.vsdx',
        'title': '图12  端—边—云协同平台与MLOps闭环',
        'layers': [
            [('uav','无人机与机载传感器'),('fixed','固定监测终端'),('edge','边缘计算盒与专网通信'),('security','身份认证、加密与日志审计')],
            [('screen','边缘质量筛查与快速推理'),('cloud','云端精检、GIS与趋势分析'),('model','模型注册、评测、灰度发布与回滚'),('data','数据湖、台账与样本谱系')],
            [('alarm','风险告警与人工确认'),('work','工单派发、处置与销项'),('feedback','结果回填与模型持续学习')],
        ],
        'edges': [('uav','screen'),('fixed','screen'),('edge','screen'),('security','cloud'),('screen','cloud'),('screen','data'),('cloud','model'),('data','model'),('cloud','alarm'),('alarm','work'),('work','feedback'),('feedback','data'),('feedback','model')],
    },
    {
        'file': 'Fig13_Test_Evaluation_Evidence_Chain.vsdx',
        'title': '图13  技术指标、第三方检测与申报证据链',
        'layers': [
            [('algo','算法精度与泛化'),('flight','自主飞行与安全'),('measure','树障量测与风险分级'),('system','平台时延、稳定性与安全')],
            [('lab','标准数据集与实验室对比'),('field','真实线路现场试验'),('third','CMA/CNAS第三方检测'),('novelty','科技查新与成果评价')],
            [('report','检测/验收/鉴定报告'),('proof','专利、论文、软著与标准'),('application','用户应用证明与运行日志'),('audit','合同、发票与专项审计')],
        ],
        'edges': [('algo','lab'),('flight','field'),('measure','field'),('system','third'),('lab','third'),('field','third'),('third','report'),('novelty','report'),('novelty','proof'),('report','application'),('application','audit')],
    },
    {
        'file': 'Fig14_Industry_Academia_StageGate.vsdx',
        'title': '图14  产学研Stage-Gate协同与成果转化机制',
        'layers': [
            [('user','用户需求、场景与评价指标'),('univ','高校基础理论与算法创新'),('institute','科研院所嵌入式与系统集成'),('industry','企业验证、产品化与市场推广')],
            [('g1','G1 需求冻结'),('g2','G2 算法原型'),('g3','G3 平台联调'),('g4','G4 现场验证'),('g5','G5 检测评价与定型')],
            [('ip','知识产权组合与贡献留痕'),('product','平台+终端+SDK产品体系'),('transfer','许可、作价入股、联合投标'),('talent','联合人才培养与持续创新')],
        ],
        'edges': [('user','g1'),('univ','g2'),('institute','g3'),('industry','g4'),('g1','g2'),('g2','g3'),('g3','g4'),('g4','g5'),('g5','ip'),('g5','product'),('ip','transfer'),('product','transfer'),('transfer','talent')],
    },
    {
        'file': 'Fig15_Application_Benefit_Scaling.vsdx',
        'title': '图15  应用场景、效益传导与规模化推广路径',
        'layers': [
            [('transmission','高压输电与跨区通道'),('distribution','城镇配电网'),('forest','山区/林区树障治理'),('mine','矿区与工业园区'),('emergency','灾害应急与特巡')],
            [('capability','自主采集—智能识别—三维量测—工单闭环'),('standard','模块化产品、接口与评价标准'),('service','软件许可、装备销售与巡检服务')],
            [('safety','降低高空与偏远作业风险'),('economic','节约巡检成本与避免停电损失'),('quality','提升缺陷发现和闭环处置效率'),('scale','跨区域、跨电压等级复制推广')],
        ],
        'edges': [('transmission','capability'),('distribution','capability'),('forest','capability'),('mine','capability'),('emergency','capability'),('capability','standard'),('capability','service'),('standard','scale'),('service','economic'),('capability','safety'),('capability','quality'),('economic','scale'),('quality','scale')],
    },
]


def safe_set(obj, name, value):
    try:
        setattr(obj, name, value)
    except Exception:
        pass


def positions(n: int):
    if n == 1:
        return [PAGE_W / 2], 4.2
    margin = 0.85
    usable = PAGE_W - 2 * margin
    xs = [margin + usable * i / (n - 1) for i in range(n)]
    if n >= 5:
        width = 2.15
    elif n == 4:
        width = 2.45
    elif n == 3:
        width = 2.85
    else:
        width = 3.6
    return xs, width


def add_rect(page, media, x, y, w, h, text, fill, line, text_color='#1F2937'):
    shape = media.rectangle.copy(page=page)
    shape.x = x
    shape.y = y
    shape.width = w
    shape.height = h
    shape.text = text
    safe_set(shape, 'fill_color', fill)
    safe_set(shape, 'line_color', line)
    safe_set(shape, 'text_color', text_color)
    return shape


def build_one(diagram: dict):
    media_probe = Media()
    template = Path(media_probe._media_vsdx.filename)
    media_probe._media_vsdx.close_vsdx()
    out_file = VSDX_DIR / diagram['file']

    with VisioFile(str(template)) as vis:
        page = vis.pages[0]
        page.name = diagram['title'][:31]
        page.width = PAGE_W
        page.height = PAGE_H
        for shape in list(page.child_shapes):
            shape.remove()

        media = Media()
        add_rect(page, media, PAGE_W/2, 7.08, 12.65, 0.62, diagram['title'], '#1F4E78', '#1F4E78', '#FFFFFF')
        add_rect(page, media, PAGE_W/2, 0.31, 12.65, 0.34,
                 'Microsoft Visio原生VSDX：所有文本、图形和连接线均为独立可编辑对象',
                 '#F2F2F2', '#BFBFBF', '#595959')

        y_positions = [5.65, 3.62, 1.58]
        nodes = {}
        for layer_index, layer in enumerate(diagram['layers']):
            xs, width = positions(len(layer))
            fill, line = PALETTE[layer_index % len(PALETTE)]
            for x, (key, text) in zip(xs, layer):
                nodes[key] = add_rect(page, media, x, y_positions[layer_index], width, 0.90, text, fill, line)

        for source, target in diagram['edges']:
            if source not in nodes or target not in nodes:
                raise KeyError(f'Unknown edge {source}->{target} in {diagram["file"]}')
            connector = Connect.create(page=page, from_shape=nodes[source], to_shape=nodes[target])
            connector.text = ''
            safe_set(connector, 'line_color', '#7F8C8D')
            safe_set(connector, 'end_arrow', 4)

        vis.save_vsdx(str(out_file))
        media._media_vsdx.close_vsdx()
    return out_file


def validate_vsdx(path: Path):
    required = {
        '[Content_Types].xml',
        '_rels/.rels',
        'visio/document.xml',
        'visio/pages/pages.xml',
        'visio/pages/page1.xml',
    }
    if not zipfile.is_zipfile(path):
        raise RuntimeError(f'Not a ZIP/OPC package: {path}')
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        missing = required - names
        if missing:
            raise RuntimeError(f'Missing required VSDX parts in {path}: {missing}')
        bad = zf.testzip()
        if bad:
            raise RuntimeError(f'CRC error in {path}: {bad}')
        for name in names:
            if name.endswith('.xml') or name.endswith('.rels'):
                ET.fromstring(zf.read(name))
    with VisioFile(str(path)) as vis:
        if len(vis.pages) != 1:
            raise RuntimeError(f'Unexpected page count in {path}')
        count = len(vis.pages[0].all_shapes)
        if count < 8:
            raise RuntimeError(f'Too few editable shapes in {path}: {count}')
        return count


built = []
for diagram in DIAGRAMS:
    path = build_one(diagram)
    shape_count = validate_vsdx(path)
    built.append((diagram, path, shape_count))

readme = OUT / 'README_Visio_Repair.txt'
lines = [
    '输配电线路无人机巡检与树障隐患智能感知——可编辑Visio修复版',
    '',
    '本批文件以有效的原生VSDX模板重建，不再使用手工拼装的OPC包。',
    '每个文本框、矩形和连接线均是Microsoft Visio中的独立对象，可移动、改字、改色和重新连接。',
    '文件名采用英文短名称以避免中文路径兼容问题；图题与申报材料中的图号保持一致。',
    '',
    '文件清单：',
]
for diagram, path, count in built:
    lines.append(f'- {path.name} | {diagram["title"]} | editable shapes: {count}')
lines += [
    '',
    '验证项目：ZIP/OPC完整性、CRC、全部XML/RELS可解析、必要Visio部件存在、vsdx库重新打开成功、可编辑对象数量检查。',
    '建议：先解压总包，再直接用Microsoft Visio 2016/2019/2021/2024或Visio Plan 2打开VSDX文件。',
]
readme.write_text('\n'.join(lines), encoding='utf-8')

checksum = OUT / 'SHA256SUMS.txt'
checksum.write_text('\n'.join(f'{hashlib.sha256(p.read_bytes()).hexdigest()}  {p.name}' for _, p, _ in built), encoding='utf-8')

package = OUT / 'Power_Line_UAV_Valid_Editable_Visio_Files.zip'
with zipfile.ZipFile(package, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.write(readme, readme.name)
    zf.write(checksum, checksum.name)
    for _, path, _ in built:
        zf.write(path, f'Valid_Editable_Visio_Files/{path.name}')

with zipfile.ZipFile(package) as zf:
    if zf.testzip() is not None:
        raise RuntimeError('Package ZIP CRC validation failed')

print('VALID_VISIO_BUILD_OK')
print(f'files={len(built)} package={package} size={package.stat().st_size}')
for diagram, path, count in built:
    print(path.name, count, path.stat().st_size)
