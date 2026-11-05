# -*- coding: utf-8 -*-
"""
通用影视全流程专业文学剧本与横向8列导演分镜生成及8道断言质检引擎 (screenplay_gold_engine.py)
- 严格执行文学剧本全部 8 道硬性断言（零表格、零顶部导读框、零△、零〔〕、零镜头术语与占位括号、标准场标题正则、标准片尾标记、对白与微动作严格交替）
- 支持横向宽幅 8 列专业导演分镜脚本表1（分场预计时长与镜数）与表2（逐镜单镜时长加总）100% 精确算术核验
"""
import re, docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from pub_layout_engine import init_pub_doc, add_pub_p, add_pub_table

STORYBOARD_8_HEADERS = [
    "镜号",
    "单镜时长",
    "景别",
    "镜头焦段",
    "机位与运镜方式",
    "画面表演与走位调度（文字落实为画面）",
    "同期台词／画外音／屏幕文字",
    "光影色温与同期声效"
]
STORYBOARD_8_WIDTHS = [0.65, 0.65, 0.65, 0.70, 1.45, 2.75, 1.85, 1.65]

SCENE_HEADER_RE = re.compile(r"^\d+-\d+\u3000(?:内|外|外/内|内/外)\u3000.+ — (?:日|夜|晨|昏|黄昏|拂晓)$")
FORBIDDEN_SCRIPT_PLACEHOLDERS = [
    "（示例", "（情感填充物", "【角色名】", "【具体", "推镜头", "拉镜头", "摇镜头", "移镜头", "特写镜头"
]

def verify_screenplay_8_assertions(doc_path):
    """对单集或总本文学剧本执行全部 8 道工业级格式与台词微动作交替硬断言"""
    doc = docx.Document(doc_path)
    # 断言1：零表格
    assert len(doc.tables) == 0, f"[断言1失败] 文学剧本严禁包含任何表格: {doc_path}"

    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    assert len(paras) >= 10, f"[剧本内容过短] {doc_path}"
    full_txt = "\n".join(paras)

    # 断言2：零顶部导读框（前3段必须直接为《剧名》、第X集标题与第1场标准场标题）
    assert paras[0].startswith("《") and paras[0].endswith("》"), f"[断言2失败] 首段必须为标准剧名而非导读框: {paras[0]}"
    assert "专业文学剧本" in paras[1], f"[断言2失败] 第二段必须为剧本集标题: {paras[1]}"
    assert SCENE_HEADER_RE.match(paras[2]), f"[断言2失败] 第三段必须直接进入第一场标题，严禁插入顶部导读框: {paras[2]}"

    # 断言3：零 △ 符号
    assert "△" not in full_txt, f"[断言3失败] 文学剧本严禁使用 △ 符号: {doc_path}"

    # 断言4：零 〔〕 方括号
    assert "〔" not in full_txt and "〕" not in full_txt, f"[断言4失败] 文学剧本严禁使用 〔〕 方括号: {doc_path}"

    # 断言5：零镜头术语与零说明性占位括号
    for bad_ph in FORBIDDEN_SCRIPT_PLACEHOLDERS:
        assert bad_ph not in full_txt, f"[断言5失败] 文学剧本包含违规占位符或镜头术语 '{bad_ph}': {doc_path}"

    # 断言6：标准场标题正则校验及紧随人物行校验
    scene_cnt = 0
    for idx, line in enumerate(paras):
        if re.match(r"^\d+-\d+", line):
            assert SCENE_HEADER_RE.match(line), f"[断言6失败] 场标题格式不合规: {line}"
            assert idx + 1 < len(paras) and paras[idx + 1].startswith("人物："), f"[断言6失败] 场标题后缺少标准人物行: {line}"
            scene_cnt += 1
    assert scene_cnt >= 3, f"[断言6失败] 文学剧本有效场次数不足3场: {scene_cnt}"

    # 断言7：标准片尾标记
    assert "*开始片尾音乐：" in full_txt and "切至黑屏，片尾" in full_txt, f"[断言7失败] 缺少标准片尾标记: {doc_path}"

    # 断言8：校验台词与微动作交替律（禁止连续两段纯对白无动作间隔）
    consec_dialogue = 0
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        left_in = p.paragraph_format.left_indent
        if left_in and int(left_in.pt) >= 90:
            consec_dialogue += 1
            assert consec_dialogue <= 1, f"[断言8失败] 出现连续裸对白未穿插可视微动作: {t[:40]}"
        elif p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            consec_dialogue = 0
    return True

def verify_storyboard_math(doc_path):
    """校验横向8列导演分镜脚本：表1分场预计时长、镜数与表2逐镜单镜秒数加总 100% 精确相等"""
    doc = docx.Document(doc_path)
    assert len(doc.tables) >= 2, f"[分镜表缺失] {doc_path}"
    t_overview = doc.tables[0]
    t_shots = doc.tables[1]
    assert len(t_shots.columns) == 8, f"[分镜表列数不为8列] {doc_path}"

    declared_scenes = {}
    for ri in range(1, len(t_overview.rows)):
        sc_cell = t_overview.rows[ri].cells[0].text.strip()
        dur_cell = t_overview.rows[ri].cells[1].text.strip()
        m_sc = re.match(r"^(\d+)-(\d+)", sc_cell)
        assert m_sc, f"[分镜表1场次编号格式错误] {sc_cell}"
        sc_num = int(m_sc.group(2))
        m_cnt = re.search(r"共\s*(\d+)\s*镜", dur_cell)
        m_sec = re.search(r"([\d.]+)\s*秒", dur_cell)
        m_mmss = re.search(r"(\d{2}):(\d{2})", dur_cell)
        assert m_cnt and (m_sec or m_mmss), f"[分镜表1时长或镜数声明格式不规范] {dur_cell}"
        exp_cnt = int(m_cnt.group(1))
        exp_sec = float(m_sec.group(1)) if m_sec else (int(m_mmss.group(1)) * 60 + int(m_mmss.group(2)))
        declared_scenes[sc_num] = {"exp_cnt": exp_cnt, "exp_sec": exp_sec, "actual_cnt": 0, "actual_sec": 0.0}

    for ri in range(1, len(t_shots.rows)):
        shot_id = t_shots.rows[ri].cells[0].text.strip()
        dur_str = t_shots.rows[ri].cells[1].text.strip()
        m_sid = re.match(r"^S(\d+)-(\d+)", shot_id)
        assert m_sid, f"[分镜表2镜号格式错误] {shot_id}"
        sc_num = int(m_sid.group(1))
        assert sc_num in declared_scenes, f"[分镜表2出现未在表1声明的场次] {shot_id}"
        m_d = re.match(r"^([\d.]+)s$", dur_str)
        assert m_d, f"[分镜表2单镜时长格式错误] {dur_str}"
        sec_val = float(m_d.group(1))
        declared_scenes[sc_num]["actual_cnt"] += 1
        declared_scenes[sc_num]["actual_sec"] += sec_val

    for sc_num, info in declared_scenes.items():
        assert info["exp_cnt"] == info["actual_cnt"], (
            f"[分镜镜数算术不符] {doc_path} 第1-{sc_num}场 表1声明 {info['exp_cnt']} 镜 != 表2实际 {info['actual_cnt']} 镜"
        )
        assert abs(info["exp_sec"] - info["actual_sec"]) < 0.01, (
            f"[分镜时长算术不符] {doc_path} 第1-{sc_num}场 表1声明 {info['exp_sec']}s != 表2加总 {info['actual_sec']}s"
        )
    return True

def build_screenplay_doc(series_title, ep_num, ep_title, scenes_list, out_path):
    """根据结构化场次数据生成符合8道硬断言的标准出版级文学剧本 Word 文档"""
    doc = init_pub_doc(landscape=False)
    add_pub_p(doc, f"《{series_title}》", WD_ALIGN_PARAGRAPH.CENTER, True, 18, "黑体", 6, 2, (20, 40, 80))
    add_pub_p(doc, f"第{ep_num:02d}集　《{ep_title}》　专业文学剧本", WD_ALIGN_PARAGRAPH.CENTER, True, 14, "黑体", 2, 12, (25, 45, 85))

    for sc in scenes_list:
        add_pub_p(doc, sc["header"], WD_ALIGN_PARAGRAPH.LEFT, True, 11.5, "黑体", 10, 4, (20, 40, 80))
        add_pub_p(doc, f"人物：{sc['chars']}", WD_ALIGN_PARAGRAPH.LEFT, True, 10.5, "黑体", 0, 4)
        for item in sc["blocks"]:
            if item["type"] == "action":
                add_pub_p(doc, item["text"], WD_ALIGN_PARAGRAPH.LEFT, False, 10.5, "宋体", 2, 4, first_line_indent_pt=21)
            elif item["type"] == "dialogue":
                add_pub_p(doc, item["speaker"], WD_ALIGN_PARAGRAPH.CENTER, True, 10.5, "黑体", 4, 1)
                p = add_pub_p(doc, item["line"].strip("“”\"'"), WD_ALIGN_PARAGRAPH.LEFT, False, 10.5, "宋体", 1, 4)
                p.paragraph_format.left_indent = Pt(108)
                p.paragraph_format.right_indent = Pt(72)

    add_pub_p(doc, f"*开始片尾音乐：《{series_title}》主题片尾曲", WD_ALIGN_PARAGRAPH.LEFT, True, 10.5, "黑体", 12, 4)
    add_pub_p(doc, "切至黑屏，片尾", WD_ALIGN_PARAGRAPH.RIGHT, True, 10.5, "黑体", 2, 6)
    doc.save(out_path)
    verify_screenplay_8_assertions(out_path)
    return out_path
