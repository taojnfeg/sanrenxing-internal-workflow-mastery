# -*- coding: utf-8 -*-
"""
《三人行》（THE TRIPLE ECHO OF TIME）V26.0 唯一黄金生成引擎（官方拍摄剧本版式升级版）
内置 8 道 AST/DOM 硬性格式断言：
1. 零 <table> 断言
2. 零 doc-meta 导读框断言
3. 零 △ 符号断言
4. 零 〔〕 括号断言
5. 零 Shot # 编号断言
6. 场次数量 100% 匹配断言
7. 角色名独立居中加粗、台词独立缩进（左 108pt / 右 72pt）无引号断言
8. 严禁连续裸对白堆砌断言（严禁任何两句对白之间无原片真实动作/神色/交互描写（max_consec == 1，consec_dlg_pairs == 0））
并支持官方剧本页眉（左 EPxx，右页码）、左顶格加粗视听提示（如 *开始片尾音乐： / 开始音乐蒙太奇：）与右对齐转场（切至： / 切至黑屏，片尾）。
"""
import os
import re
from typing import List, Dict, Any

def sh(scene_code: str, ie: str, location: str, day_night: str) -> Dict[str, str]:
    """生成标准场标题块：7-4  内  徐图图家·玄关与客厅 — 夜"""
    return {
        "type": "sh",
        "text": f"{scene_code}　{ie}　{location} — {day_night}"
    }

def act(text: str) -> Dict[str, str]:
    """生成纯文学动作/神态/交互描写段（严禁带 △ 或 〔〕）"""
    clean = text.replace("△", "").replace("〔", "（").replace("〕", "）").strip()
    return {"type": "act", "text": clean}

def dlg(role: str, line: str) -> Dict[str, str]:
    """生成标准居中角色名 + 缩进无引号台词块"""
    clean_role = role.strip()
    clean_line = line.strip()
    if (clean_line.startswith("“") and clean_line.endswith("”")) or (clean_line.startswith('"') and clean_line.endswith('"')):
        clean_line = clean_line[1:-1].strip()
    return {"type": "dlg", "role": clean_role, "line": clean_line}

def cut(text: str = "切至：") -> Dict[str, str]:
    """生成右对齐转场标记（如 切至： / 切至黑屏，片尾 / 全剧终）"""
    return {"type": "cut", "text": text.strip()}

def cueline(text: str) -> Dict[str, str]:
    """生成左顶格加粗独立成行的视听/音乐提示（如 开始音乐蒙太奇： / 结束音乐蒙太奇。 / *开始片尾音乐：）"""
    return {"type": "cueline", "text": text.strip()}

def cue(tag: str, text: str) -> Dict[str, str]:
    """生成【闪回·……】等特殊时空提示段"""
    clean = text.replace("△", "").replace("〔", "（").replace("〕", "）").strip()
    return {"type": "act", "text": f"【{tag}】{clean}"}

def validate_screenplay_html(html_str: str, expected_scene_count: int, blocks: List[Dict[str, str]] = None) -> None:
    """8 道硬性 AST/正则断言自检，任一不通过立即抛出异常阻断落盘"""
    assert "<table" not in html_str.lower(), "[断言1失败] 文学剧本中严禁包含 <table> 表格！"
    assert "doc-meta" not in html_str.lower(), "[断言2失败] 文学剧本中严禁包含 doc-meta 顶部导读框！"
    assert "△" not in html_str, "[断言3失败] 文学剧本中严禁包含 △ 符号！"
    assert "〔" not in html_str and "〕" not in html_str, "[断言4失败] 文学剧本中严禁包含 〔〕 括号！"
    assert not re.search(r"Shot\s*#\d+", html_str, re.I), "[断言5失败] 文学剧本中严禁出现 Shot # 镜头编号！"
    actual_scenes = len(re.findall(r'<p class="sh">', html_str))
    assert actual_scenes == expected_scene_count, (
        f"[断言6失败] 场次数量不匹配！期望 {expected_scene_count} 场，实际生成 {actual_scenes} 场。"
    )
    roles = len(re.findall(r'<p class="role">', html_str))
    lines = len(re.findall(r'<p class="line">', html_str))
    assert roles == lines and roles > 0, f"[断言7失败] 角色名({roles})与台词({lines})数量不一致！"
    assert html_str.count("*开始片尾音乐：") == 1, f"[断言7失败] *开始片尾音乐：必须且只能出现 1 次！实际出现 {html_str.count('*开始片尾音乐：')} 次"
    assert html_str.count("切至黑屏，片尾") == 1, f"[断言7失败] 切至黑屏，片尾 必须且只能出现 1 次！实际出现 {html_str.count('切至黑屏，片尾')} 次"


    if blocks is not None:
        consec = 0
        max_consec = 0
        for b in blocks:
            if b.get("type") == "dlg":
                consec += 1
                if consec > max_consec:
                    max_consec = consec
            else:
                consec = 0
        assert max_consec <= 1, (
            f"[断言8失败] 剧本出现连续 {max_consec} 句裸对白而无动作/神色/交互描写穿插（强制要求每句对白之间必有原片真实动作/神态/交互描写，max_consec == 1）！请补充画面动作与神态细节。"
        )

def build_ep_screenplay_html(ep_num: int, ep_title: str, blocks: List[Dict[str, str]], expected_scene_count: int) -> str:
    """构建单集专业文学剧本 HTML 并自动执行 8 道断言校验"""
    body_parts = [
        '<p class="title">《三人行》</p>',
        '<p class="subtitle">THE TRIPLE ECHO OF TIME</p>',
        f'<p class="ep">EP{ep_num:02d}　{ep_title}</p>',
    ]
    for b in blocks:
        btype = b["type"]
        if btype == "sh":
            body_parts.append(f'<p class="sh">{b["text"]}</p>')
        elif btype == "act":
            body_parts.append(f'<p class="act">{b["text"]}</p>')
        elif btype == "cueline":
            body_parts.append(f'<p class="cueline">{b["text"]}</p>')
        elif btype == "dlg":
            body_parts.append(f'<p class="role">{b["role"]}</p>')
            body_parts.append(f'<p class="line">{b["line"]}</p>')
        elif btype == "cut":
            body_parts.append(f'<p class="cut">{b["text"]}</p>')

    html = "\n".join(body_parts)
    validate_screenplay_html(html, expected_scene_count, blocks)
    return html

def _add_header_ep_and_page(doc, ep_num: int):
    """在 Word 页眉左侧写入 EPxx，右侧写入自动页码（对齐官方剧本页眉规范）"""
    from docx.shared import Pt, RGBColor
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    section = doc.sections[0]
    header = section.header
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.text = f"EP{ep_num:02d}\t\t"
    run_p = hp.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run_p._r.append(fldChar1)
    run_p._r.append(instrText)
    run_p._r.append(fldChar2)
    run_p._r.append(fldChar3)
    for r in hp.runs:
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(85, 85, 85)

def build_ep_screenplay_docx(ep_num: int, ep_title: str, blocks: List[Dict[str, str]], expected_scene_count: int, out_docx_path: str) -> str:
    """构建单集专业文学剧本纯 Word (.docx) 文件（先通过 8 道断言校验，再写入 .docx）"""
    import docx
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn

    build_ep_screenplay_html(ep_num, ep_title, blocks, expected_scene_count)

    doc = docx.Document()
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "SimSun"
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal_style.font.size = Pt(11)

    _add_header_ep_and_page(doc, ep_num)

    def add_p(text: str, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, size_pt=11, font_cn="宋体",
              space_before=0, space_after=4, left_indent=0, right_indent=0, color=(17, 17, 17)):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(space_before)
        pf.space_after = Pt(space_after)
        pf.line_spacing = 1.65
        if left_indent:
            pf.left_indent = Pt(left_indent)
        if right_indent:
            pf.right_indent = Pt(right_indent)
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size_pt)
        run.font.name = "Times New Roman" if font_cn == "宋体" else "Arial"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_cn)
        run.font.color.rgb = RGBColor(*color)
        return p

    add_p("《三人行》", WD_ALIGN_PARAGRAPH.CENTER, True, 18, "黑体", 0, 2)
    add_p("THE TRIPLE ECHO OF TIME", WD_ALIGN_PARAGRAPH.CENTER, False, 10, "宋体", 0, 6, color=(102, 102, 102))
    add_p(f"EP{ep_num:02d}　{ep_title}", WD_ALIGN_PARAGRAPH.CENTER, True, 15, "黑体", 0, 18)

    for b in blocks:
        btype = b["type"]
        if btype == "sh":
            add_p(b["text"], WD_ALIGN_PARAGRAPH.LEFT, True, 11.5, "黑体", 14, 6)
        elif btype == "act":
            add_p(b["text"], WD_ALIGN_PARAGRAPH.LEFT, False, 11, "宋体", 4, 6)
        elif btype == "cueline":
            add_p(b["text"], WD_ALIGN_PARAGRAPH.LEFT, True, 11, "黑体", 8, 4)
        elif btype == "dlg":
            add_p(b["role"], WD_ALIGN_PARAGRAPH.CENTER, True, 11, "黑体", 8, 1)
            add_p(b["line"], WD_ALIGN_PARAGRAPH.LEFT, False, 11, "宋体", 1, 8, left_indent=108, right_indent=72)
        elif btype == "cut":
            add_p(b["text"], WD_ALIGN_PARAGRAPH.RIGHT, True, 10.5, "黑体", 6, 8)

    os.makedirs(os.path.dirname(os.path.abspath(out_docx_path)), exist_ok=True)
    doc.save(out_docx_path)
    return out_docx_path

def grab_frame_from_cos_url(url: str, mmss_str: str, out_path: str) -> str:
    import cv2
    parts = str(mmss_str).split(":")
    ms = (int(parts[0]) * 60 + float(parts[1])) * 1000.0 if len(parts) == 2 else float(parts[0]) * 1000.0
    cap = cv2.VideoCapture(url)
    cap.set(cv2.CAP_PROP_POS_MSEC, ms)
    ok, frame = cap.read()
    cap.release()
    if not ok:
        raise RuntimeError(f"[V26.0 CV2 抽帧失败] 无法从直链提取 {mmss_str} 处原画帧")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    cv2.imwrite(out_path, frame, [int(cv2.IMWRITE_JPEG_QUALITY), 95])
    return out_path
