# -*- coding: utf-8 -*-
"""docx_domain —— 域感知的 Word 整段文本替换。

为什么需要这个模块
------------------
python-docx 里最自然的改文字方式是::

    paragraph.text = "新的整段文字"

这行代码会删掉段落下的**所有** ``w:r``，然后插入一个新的 run。问题在于，
真实文档里的段落并不是纯文本：

* ``[1]`` ``[14]`` 这类引用角标是 Word 的 **REF 交叉引用域**，XML 结构为
  ``fldChar begin → instrText " REF _Ref123 \\r \\h " → fldChar separate
  → 域结果文本 → fldChar end``。域被删后角标变成死文字，全文引用无法再更新。
* ``10⁻³`` 的上标 ``−3`` 是一个带 ``w:vertAlign`` 的独立 run，会被压成普通字符。
* 段落里若嵌了图片（``w:drawing``），会被直接删除。
* ``pPr``（缩进、行距、对齐、样式）虽然不在 run 里，但重建段落时若处理不当也会丢。

肉眼几乎看不出这些损坏，只有在打印或"更新域"时才暴露——属于典型的
**静默数据损坏**。

本模块的做法
------------
把段落内容切成两类：

1. **token**：必须原样保留的结构单元（整个域 / 上下标 run / 图片）
2. **可替换文本**：普通 run 拼起来的文本

替换时，在新文本中按顺序查找每个 token，把 token 之间的文本用普通 run 填充，
token 本身原样深拷贝。这样既换了文字，又保住了结构。

限制（有意为之，不静默失败）
--------------------------
* 段落内含超链接（``w:hyperlink``）包裹的 run 时不处理，直接报错。
* 新文本中 token 顺序/数量与原文不一致时拒绝执行。
* 新旧文本的数字集合必须完全一致，否则拒绝执行（防止改写时改错数据）。
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

MEDIA_TAGS = (qn("w:drawing"), qn("w:pict"), qn("w:object"))
CITE_RE = re.compile(r"\[\d+\]")
NUM_RE = re.compile(r"\d+(?:\.\d+)?")


# --------------------------------------------------------------------------- #
# run 级判定
# --------------------------------------------------------------------------- #
def has_media(run_el) -> bool:
    """run 内是否嵌了图片/图形对象。"""
    return any(run_el.findall(".//" + t) for t in MEDIA_TAGS)


def run_text(run_el) -> str:
    return "".join(t.text or "" for t in run_el.findall(qn("w:t")))


def fld_type(run_el) -> str | None:
    """返回 run 内 fldChar 的类型：begin / separate / end，无则 None。"""
    fc = run_el.findall(qn("w:fldChar"))
    return fc[0].get(qn("w:fldCharType")) if fc else None


def rpr_of(run_el):
    return run_el.find(qn("w:rPr"))


def is_vert_align(run_el) -> bool:
    """是否为上下标 run（10^-3 的 -3 属于此类）。"""
    rpr = rpr_of(run_el)
    return rpr is not None and rpr.find(qn("w:vertAlign")) is not None


# --------------------------------------------------------------------------- #
# 段落结构解析
# --------------------------------------------------------------------------- #
@dataclass
class Token:
    """必须原样保留的结构单元。"""
    kind: str            # 'field' | 'vert' | 'media'
    text: str            # 该 token 在纯文本中对应的字符串
    runs: list = field(default_factory=list)


def _direct_runs(p_el):
    """段落直接子级的 run。超链接里的 run 不计入（见模块文档的限制）。"""
    return p_el.findall(qn("w:r"))


def _field_result_text(grp) -> str:
    """取域结果部分的可见文本（separate 之后、end 之前）。"""
    out, after_sep = [], False
    for r in grp:
        t = fld_type(r)
        if t == "separate":
            after_sep = True
            continue
        if t == "end":
            break
        if after_sep:
            out.append(run_text(r))
    return "".join(out)


def parse_tokens(p_el) -> list[Token]:
    """按文档顺序抽出所有 token。"""
    runs = _direct_runs(p_el)
    tokens: list[Token] = []
    i = 0
    while i < len(runs):
        run = runs[i]
        if fld_type(run) == "begin":
            grp, j = [], i
            while j < len(runs):
                grp.append(runs[j])
                if fld_type(runs[j]) == "end":
                    break
                j += 1
            tokens.append(Token("field", _field_result_text(grp), grp))
            i = j + 1
            continue
        if has_media(run):
            tokens.append(Token("media", "", [run]))
        elif is_vert_align(run) and run_text(run):
            tokens.append(Token("vert", run_text(run), [run]))
        i += 1
    return tokens


def base_rpr(p_el):
    """挑出正文的字符格式：取最长的普通 run 的 rPr 作为模板。"""
    best, best_len = None, -1
    for r in _direct_runs(p_el):
        if has_media(r) or is_vert_align(r) or r.findall(qn("w:fldChar")) or r.findall(qn("w:instrText")):
            continue
        n = len(run_text(r))
        if n > best_len:
            best, best_len = rpr_of(r), n
    return best


def _make_run(text: str, rpr):
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


# --------------------------------------------------------------------------- #
# 核心：安全替换
# --------------------------------------------------------------------------- #
class ReplaceError(Exception):
    """拒绝替换时抛出，附带原因，不做静默降级。"""


def replace_paragraph(p, new_text: str, *, strict_numbers: bool = True) -> str:
    """把段落 ``p`` 的可见文本替换为 ``new_text``，保留全部结构。

    返回一段人类可读的执行说明；无法安全执行时抛 ``ReplaceError``。
    """
    p_el = p._element
    runs = _direct_runs(p_el)

    if any(has_media(r) for r in runs):
        raise ReplaceError("段落内含图片，整段重建会丢失图片；请改用片段替换")
    if p_el.findall(qn("w:hyperlink")):
        raise ReplaceError("段落内含超链接，run 层级不确定；请先降级为普通文本")

    tokens = parse_tokens(p_el)

    # 1) 按 token 在新文本中的出现位置切段
    segments: list[tuple[str, object]] = []
    pos = 0
    for idx, tok in enumerate(tokens):
        if not tok.text:
            continue                      # 图片等无文本 token
        found = new_text.find(tok.text, pos)
        if found < 0:
            raise ReplaceError(f"新文本缺少必须保留的 token {tok.text!r}，拒绝执行")
        if found > pos:
            segments.append(("text", new_text[pos:found]))
        segments.append(("token", idx))
        pos = found + len(tok.text)
    if pos < len(new_text):
        segments.append(("text", new_text[pos:]))

    # 2) 一致性检查：token 不得重复出现、不得新增引用角标
    for kind, val in segments:
        if kind == "text" and CITE_RE.search(val):
            raise ReplaceError("新文本中出现了 token 之外的引用角标，数量与原文不一致")

    old_text = "".join(run_text(r) for r in runs)
    if strict_numbers and sorted(NUM_RE.findall(old_text)) != sorted(NUM_RE.findall(new_text)):
        raise ReplaceError("数字集合发生变化，疑似改写时动到了数据；如确认无误请关闭 strict_numbers")

    # 3) 定位锚点：第一个普通 run
    plain = [r for r in runs if not has_media(r)]
    if not plain:
        raise ReplaceError("段落内没有可放置文本的普通 run")
    anchor = plain[0]
    cursor = anchor

    # 4) 依次插入
    rpr = base_rpr(p_el)
    for kind, val in segments:
        if kind == "text":
            els = [_make_run(val, rpr)]
        else:
            els = [copy.deepcopy(r) for r in tokens[val].runs]
        for el in els:
            cursor.addnext(el)
            cursor = el

    # 5) 删掉旧 run（pPr 从未被触碰）
    for r in plain:
        p_el.remove(r)

    return f"ok：{len(tokens)} 个 token 保留，{len(segments)} 个片段重建"


# --------------------------------------------------------------------------- #
# 校验
# --------------------------------------------------------------------------- #
def inspect_paragraph(p) -> dict:
    """结构化描述段落，用于替换前后对比。"""
    p_el = p._element
    runs = _direct_runs(p_el)
    ppr = p_el.find(qn("w:pPr"))
    return {
        "text": "".join(run_text(r) for r in runs),
        "run_count": len(runs),
        "field_count": sum(1 for r in runs if fld_type(r) == "begin"),
        "fldchar_count": sum(1 for r in runs if r.findall(qn("w:fldChar"))),
        "instr_count": sum(1 for r in runs if r.findall(qn("w:instrText"))),
        "vertalign_count": sum(1 for r in runs if is_vert_align(r)),
        "media_count": sum(1 for r in runs if has_media(r)),
        "has_ppr": ppr is not None,
        "ppr_xml": (ppr.xml if ppr is not None else ""),
        "numbers": sorted(NUM_RE.findall("".join(run_text(r) for r in runs))),
    }


def verify(before: dict, after: dict) -> list[str]:
    """对比两次 inspect 结果，返回问题清单（空列表 = 通过）。"""
    issues = []
    checks = [
        ("field_count", "交叉引用域数量"),
        ("fldchar_count", "域标记数量"),
        ("instr_count", "域指令数量"),
        ("vertalign_count", "上下标 run 数量"),
        ("media_count", "图片数量"),
        ("has_ppr", "段落属性是否存在"),
        ("ppr_xml", "段落属性内容"),
        ("numbers", "数字集合"),
    ]
    for key, label in checks:
        if before[key] != after[key]:
            issues.append(f"{label}发生变化：{before[key]!r} → {after[key]!r}")
    return issues
