# -*- coding: utf-8 -*-
"""
教案格式转换器 - Streamlit 本地工具
用法: python -m streamlit run app.py
"""
import io, os, re, tempfile
from copy import deepcopy
from docx import Document
from docx.table import Table
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import streamlit as st

# ==================== 解析源教案 ====================

def parse_source(docx_path):
    doc = Document(docx_path)
    tables = doc.tables

    # 找单元分析表（8行x2列，首格含"单元"）
    unit_tbl = None
    for t in tables:
        if len(t.rows) == 8 and len(t.columns) == 2:
            txt = t.rows[0].cells[0].text
            if "单元" in txt or "设计思路" in txt:
                unit_tbl = t
                break

    unit_raw = {}
    if unit_tbl is not None:
        for row in unit_tbl.rows:
            try:
                label = row.cells[0].text.strip()
                content = row.cells[1].text.strip()
                unit_raw[label] = content
            except Exception:
                pass

    def find_unit(kw):
        for k, v in unit_raw.items():
            if kw in k:
                return v
        return ""

    unit_analysis = {
        "教材分析": find_unit("教材分析"),
        "目标要求": find_unit("教学目标"),
        "设计意图": (find_unit("设计思路") + "\n\n" + find_unit("教学方法")).strip(),
        "达成分析": find_unit("教学评价"),
    }

    # 找所有课时起始表
    lesson_starts = []
    for i, t in enumerate(tables):
        try:
            if t.rows[0].cells[0].text.strip().startswith("学校"):
                lesson_starts.append(i)
        except Exception:
            pass

    # 识别单元号
    cn_num = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10}
    unit_num = 1
    for p in doc.paragraphs:
        m = re.search(r'第([一二三四五六七八九十\d]+)单元', p.text)
        if m:
            s = m.group(1)
            unit_num = int(s) if s.isdigit() else cn_num.get(s, 1)
            break

    lessons = []
    for idx, start in enumerate(lesson_starts):
        t_head = tables[start]
        t_obj = tables[start + 1]
        t_proc = tables[start + 2]
        t_board = tables[start + 3]

        r0 = [c.text.strip() for c in t_head.rows[0].cells]
        r1 = [c.text.strip() for c in t_head.rows[1].cells]
        total = next((c.split("：", 1)[-1].strip() for c in r0 if "总课时" in c), "")
        title = next((c.split("：", 1)[-1].strip() for c in r1 if "课题" in c), "")
        leixing = next((c.split("：", 1)[-1].strip() for c in r1 if "教时" in c), "")

        obj_block = t_obj.rows[0].cells[0].text.strip()

        def extract(text, start_kw, end_kws):
            m = re.search(re.escape(start_kw), text)
            if not m:
                return ""
            s = m.end()
            end_pos = len(text)
            for ek in end_kws:
                em = re.search(re.escape(ek), text[s:])
                if em:
                    end_pos = min(end_pos, s + em.start())
            return text[s:end_pos].strip()

        objectives_text = extract(obj_block, "教学目标：", ["教学重点", "教学难点", "教材分析", "学生分析", "教学准备"])
        key_point = extract(obj_block, "教学重点：", ["教学难点", "教材分析", "学生分析"])
        diff_point = extract(obj_block, "教学难点：", ["教材分析", "学生分析", "教学准备"])
        textbook = extract(obj_block, "教材分析：", ["学生分析", "教学准备"])
        student = extract(obj_block, "学生分析：", ["教学准备"])
        objectives = [l.strip() for l in objectives_text.split("\n") if l.strip()]

        segments = []
        for ri in range(2, len(t_proc.rows)):
            row = t_proc.rows[ri]
            cells = row.cells
            if len(cells) < 4:
                continue
            seg_name = cells[0].text.strip()
            if not seg_name:
                continue
            segments.append({
                "name": seg_name,
                "teacher": cells[1].text.strip(),
                "student": cells[2].text.strip(),
                "intent": cells[3].text.strip(),
            })

        board = t_board.rows[0].cells[1].text.strip()
        practice = t_board.rows[1].cells[1].text.strip()
        reflection = t_board.rows[2].cells[1].text.strip()

        lessons.append({
            "total": total, "title": title, "leixing": leixing,
            "objectives": objectives, "key_point": key_point, "diff_point": diff_point,
            "textbook": textbook, "student": student,
            "segments": segments, "board": board, "practice": practice, "reflection": reflection,
        })

    return {"unit_analysis": unit_analysis, "lessons": lessons, "unit_num": unit_num}


# ==================== 填充辅助 ====================

def _get_first_rpr(tc):
    for r in tc.iter(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        if rPr is not None:
            return deepcopy(rPr)
    return None

def _get_first_ppr(tc):
    for p in tc.iter(qn('w:p')):
        pPr = p.find(qn('w:pPr'))
        if pPr is not None:
            return deepcopy(pPr)
    return None

def _strip_bold(rpr):
    if rpr is None: return None
    r = deepcopy(rpr)
    for tag in ('w:b', 'w:bCs'):
        for el in r.findall(qn(tag)):
            r.remove(el)
    return r

def _ensure_bold(rpr):
    if rpr is None: rpr = OxmlElement('w:rPr')
    r = deepcopy(rpr)
    if r.find(qn('w:b')) is None:
        r.insert(0, OxmlElement('w:b'))
    return r

def _make_run(text, rpr):
    r = OxmlElement('w:r')
    if rpr is not None: r.append(deepcopy(rpr))
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    return r

LABEL_RE = re.compile(r'^([^：:\n]{1,12}[:：])\s*(.*)$')

def set_tc_smart(tc, text, bold_labels=True):
    base_rpr = _get_first_rpr(tc)
    ppr = _get_first_ppr(tc)
    normal_rpr = _strip_bold(base_rpr)
    bold_rpr = _ensure_bold(base_rpr) if bold_labels else normal_rpr
    for p in tc.findall(qn('w:p')):
        tc.remove(p)
    for line in text.split('\n'):
        p = OxmlElement('w:p')
        if ppr is not None: p.append(deepcopy(ppr))
        if bold_labels:
            m = LABEL_RE.match(line)
            if m:
                label, rest = m.group(1), m.group(2)
                p.append(_make_run(label, bold_rpr))
                if rest: p.append(_make_run(rest, normal_rpr))
            else:
                p.append(_make_run(line, normal_rpr))
        else:
            p.append(_make_run(line, normal_rpr))
        tc.append(p)

def add_page_break_before(el):
    """在el前插入一个空段落，设置pageBreakBefore属性（不产生多余空白页）"""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pbb = OxmlElement('w:pageBreakBefore')
    pPr.append(pbb)
    p.append(pPr)
    el.addprevious(p)

# ==================== 多单元填充 ====================

def find_unit_blocks(doc):
    """返回 [(title_p_element, analysis_tbl, [lesson_tbls]), ...]"""
    body = doc.element.body
    blocks = []
    current_title = None
    current_analysis = None
    current_lessons = []

    for child in body.iterchildren():
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            text = ''.join(t.text or '' for t in child.iter(qn('w:t'))).strip()
            if re.match(r'^第[一二三四五六七八九十\d]+单元', text):
                # 保存上一个block
                if current_title is not None:
                    blocks.append((current_title, current_analysis, current_lessons))
                current_title = child
                current_analysis = None
                current_lessons = []
        elif tag == 'tbl':
            rows = len(child.findall(qn('w:tr')))
            if rows == 5:
                current_analysis = Table(child, doc)
            elif rows == 13:
                current_lessons.append(Table(child, doc))
    if current_title is not None:
        blocks.append((current_title, current_analysis, current_lessons))
    return blocks


def fill_one_unit(title_p, analysis_tbl, lesson_tbls, parsed, designer):
    ua = parsed["unit_analysis"]
    def fill_unit_row(ri, content):
        tcs = analysis_tbl.rows[ri]._tr.findall(qn('w:tc'))
        set_tc_smart(tcs[-1], content, bold_labels=False)
    fill_unit_row(1, ua["教材分析"])
    fill_unit_row(2, ua["目标要求"])
    fill_unit_row(3, ua["设计意图"])
    fill_unit_row(4, ua["达成分析"])

    lessons = parsed["lessons"]
    # 补/删课时表
    need = len(lessons)
    have = len(lesson_tbls)
    if need > have:
        last_el = lesson_tbls[-1]._tbl
        anchor = last_el
        for _ in range(need - have):
            new_el = deepcopy(last_el)
            anchor.addnext(new_el)
            anchor = new_el
            lesson_tbls.append(Table(new_el, lesson_tbls[-1]._parent))
    elif need < have:
        for extra in lesson_tbls[need:]:
            extra._tbl.getparent().remove(extra._tbl)
        lesson_tbls = lesson_tbls[:need]

    # 修布局
    for tbl in lesson_tbls:
        for ri in (8, 9):
            tr = tbl.rows[ri]._tr
            tcs = tr.findall(qn('w:tc'))
            for ci, span in [(1, '3'), (2, '2')]:
                tcPr = tcs[ci].find(qn('w:tcPr'))
                gs = tcPr.find(qn('w:gridSpan'))
                if gs is not None:
                    gs.set(qn('w:val'), span)
        for ri in range(5, 10):
            tr = tbl.rows[ri]._tr
            tcs = tr.findall(qn('w:tc'))
            tcPr3 = tcs[3].find(qn('w:tcPr'))
            vm = tcPr3.find(qn('w:vMerge'))
            if vm is not None:
                tcPr3.remove(vm)

    # 填每课时
    for i, lesson in enumerate(lessons):
        tbl = lesson_tbls[i]
        tr0 = tbl.rows[0]._tr
        tcs0 = tr0.findall(qn('w:tc'))
        set_tc_smart(tcs0[1], "班级：三（  ）班", bold_labels=True)
        set_tc_smart(tcs0[2], "人数：", bold_labels=True)
        set_tc_smart(tcs0[3], f"总课时：{lesson['total']}", bold_labels=True)

        tr1 = tbl.rows[1]._tr
        tcs1 = tr1.findall(qn('w:tc'))
        set_tc_smart(tcs1[1], f"课题：{lesson['title']}", bold_labels=True)
        set_tc_smart(tcs1[2], f"教时：{lesson['leixing']}", bold_labels=True)
        set_tc_smart(tcs1[3], f"设计者：{designer}", bold_labels=True)

        lines = ["教学目标："] + lesson['objectives']
        lines += ["", "教学重难点：", f"重点：{lesson['key_point']}", f"难点：{lesson['diff_point']}",
                  "", "教材分析：", lesson['textbook'], "", "学生分析：", lesson['student']]
        tr2 = tbl.rows[2]._tr
        set_tc_smart(tr2.findall(qn('w:tc'))[0], "\n".join(lines), bold_labels=True)

        seg_rows = [5, 6, 7, 8, 9]
        for si, seg in enumerate(lesson['segments'][:5]):
            ri = seg_rows[si]
            tr = tbl.rows[ri]._tr
            tcs = tr.findall(qn('w:tc'))
            set_tc_smart(tcs[0], seg['name'], bold_labels=True)
            set_tc_smart(tcs[1], seg['teacher'], bold_labels=False)
            set_tc_smart(tcs[2], seg['student'], bold_labels=False)
            set_tc_smart(tcs[3], seg['intent'], bold_labels=False)

        tr10 = tbl.rows[10]._tr
        set_tc_smart(tr10.findall(qn('w:tc'))[0], "板书设计：\n" + lesson['board'], bold_labels=True)
        tr11 = tbl.rows[11]._tr
        set_tc_smart(tr11.findall(qn('w:tc'))[0], "练习设计：\n" + lesson['practice'], bold_labels=True)
        tr12 = tbl.rows[12]._tr
        set_tc_smart(tr12.findall(qn('w:tc'))[0], "教学反思：\n" + lesson['reflection'], bold_labels=True)

        add_page_break_before(tbl._tbl)


def remove_unit_block(title_p, analysis_tbl, lesson_tbls):
    """删除整个单元块：标题段落+分析表+所有课时表"""
    body = title_p.getparent()
    # 收集要删的元素：从title_p开始，到下一个单元标题之前
    to_remove = []
    el = title_p
    while el is not None:
        nxt = el.getnext()
        to_remove.append(el)
        if nxt is None:
            break
        nxt_tag = nxt.tag.split('}')[-1]
        if nxt_tag == 'p':
            nxt_text = ''.join(t.text or '' for t in nxt.iter(qn('w:t'))).strip()
            if re.match(r'^第[一二三四五六七八九十\d]+单元', nxt_text):
                break
        el = nxt
    for el in to_remove:
        body.remove(el)


def fill_template_multi(template_path, parsed_list, designer, output_path):
    """parsed_list: list of parse_source结果，每个自动识别unit_num"""
    doc = Document(template_path)
    blocks = find_unit_blocks(doc)
    # blocks[i] 对应第(i+1)单元
    # 建立映射: unit_num -> block
    block_map = {i + 1: b for i, b in enumerate(blocks)}
    # 已填的单元号
    filled_nums = set()
    for p in parsed_list:
        un = p.get("unit_num", 1)
        if un in block_map:
            title_p, analysis_tbl, lesson_tbls = block_map[un]
            fill_one_unit(title_p, analysis_tbl, lesson_tbls, p, designer)
            filled_nums.add(un)
            # 在标题段落前加分页符：标题+分析表+课时表一起从新页开始
            add_page_break_before(title_p)
    # 删除所有未填的单元
    for un, block in block_map.items():
        if un not in filled_nums:
            remove_unit_block(block[0], block[1], block[2])
    doc.save(output_path)
    return output_path, len(filled_nums)


# ==================== Streamlit UI ====================

st.set_page_config(page_title="教案格式转换器", page_icon="📝", layout="centered")
st.title("📝 教案格式转换器")
st.caption("把任意格式的教案 Word 转成龙二小标准模板格式，支持一次传多个单元")

DEFAULT_TEMPLATE = os.path.join(os.path.dirname(__file__), "template_source.docx")

with st.sidebar:
    st.header("设置")
    designer = st.text_input("设计者姓名", value="钮梓洁")
    st.divider()
    tpl_file = st.file_uploader("上传你的模板docx（必选）", type=["docx"])

st.write("### 上传源教案（可多选，每个文件一个单元，按上传顺序填第1、2、3…单元）")
src_files = st.file_uploader("选择教案docx", type=["docx"], accept_multiple_files=True, key="src")

if src_files:
    parsed_list = []
    for f in src_files:
        tmp = os.path.join(tempfile.gettempdir(), f"src_{f.name}")
        with open(tmp, "wb") as out:
            out.write(f.getbuffer())
        try:
            p = parse_source(tmp)
            parsed_list.append((f.name, p))
            st.write(f"✅ **{f.name}** → 第{p['unit_num']}单元，{len(p['lessons'])} 课时")
        except Exception as e:
            st.error(f"❌ {f.name} 解析失败: {e}")

    if parsed_list:
        st.write("### 生成")
        if st.button("🚀 生成标准格式教案", type="primary"):
            if tpl_file is None:
                st.warning("请先在左侧上传你的模板docx")
                st.stop()
            tpl_path = os.path.join(tempfile.gettempdir(), "tpl_input.docx")
            with open(tpl_path, "wb") as f:
                f.write(tpl_file.getbuffer())
            out_path = os.path.join(tempfile.gettempdir(), "output_lesson.docx")
            try:
                _, n = fill_template_multi(tpl_path, [p for _, p in parsed_list], designer, out_path)
                with open(out_path, "rb") as f:
                    st.download_button(
                        label="⬇️ 下载生成的教案",
                        data=f.read(),
                        file_name="教案合订.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success(f"完成！已填入 {n} 个单元，未填的空单元已删除，无空白页。")
            except Exception as e:
                st.error(f"生成失败: {e}")
                import traceback
                st.code(traceback.format_exc())
