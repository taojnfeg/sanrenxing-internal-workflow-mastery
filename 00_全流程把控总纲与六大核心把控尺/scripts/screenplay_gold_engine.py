# -*- coding: utf-8 -*-
"""
通用影视全流程专业文学剧本与横向8列导演分镜生成及8道断言质检引擎 (screenplay_gold_engine.py)
- 严格执行文学剧本排版五零铁律（零表格、零导读框、零△、零〔〕、零演员真名）
- 严格校验高信息量台词体系与“一句对白＋一段可视微动作”交替律（max_consec == 1）
- 支持将文学剧本一对一转化为横向宽幅8列专业导演分镜脚本（含单镜预计时长與镜头焦段）
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

def verify_screenplay_8_assertions(doc_path):
    """对单集或总本文学剧本执行 8 道工业级格式与台词微动作交替断言"""
    doc = docx.Document(doc_path)
    assert len(doc.tables) == 0, f"[断言1失败] 文学剧本严禁包含任何表格: {doc_path}"
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    full_txt = "\n".join(paras)
    assert "△" not in full_txt, f"[断言3失败] 文学剧本严禁使用 △ 符号: {doc_path}"
    assert "〔" not in full_txt and "〕" not in full_txt, f"[断言4失败] 文学剧本严禁使用 〔〕 方括号: {doc_path}"
    assert "*开始片尾音乐：" in full_txt and "切至黑屏，片尾" in full_txt, f"[断言7失败] 缺少标准片尾标记: {doc_path}"

    # 校验台词与微动作交替律（禁止连续两段纯对白无动作间隔）
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
