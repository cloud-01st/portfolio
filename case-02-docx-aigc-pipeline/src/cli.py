# -*- coding: utf-8 -*-
"""命令行入口：定位 / 批量替换 / 校验 / 自检。

用法
----
    python cli.py selftest
    python cli.py locate report.docx --out flagged.json
    python cli.py apply thesis.docx rewrites.json -o thesis.new.docx
    python cli.py verify thesis.docx thesis.new.docx

``rewrites.json`` 结构（键为段落索引，0 起）::

    {
      "19": {"new": "改写后的整段文字……", "match": "本课题以自然图像为研究对象"},
      "20": {"new": "……"}
    }

``match`` 可选，用于在文档变动后防止错位：给出原文前缀，apply 时逐字比对，
不匹配就拒绝执行这一条，而不是照着索引硬改。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from docx_domain import (
    ReplaceError,
    inspect_paragraph,
    replace_paragraph,
    run_text,
    verify,
)

# 常见 AIGC 检测报告的标注配色
RISK_COLORS = {
    "F12828": "high",     # 红
    "F39800": "medium",   # 橙
    "B0B0B0": "low",      # 灰
}


# --------------------------------------------------------------------------- #
def cmd_locate(args) -> int:
    """扫描文档，按 run 的字体颜色把段落分级。"""
    doc = Document(args.path)
    hits = []
    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue
        colors = {}
        for r in p._element.findall(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                continue
            color = rpr.find(qn("w:color"))
            if color is None:
                continue
            val = (color.get(qn("w:val")) or "").upper()
            if val in RISK_COLORS:
                colors[RISK_COLORS[val]] = colors.get(RISK_COLORS[val], 0) + len(run_text(r))
        if not colors:
            continue
        level = "high" if "high" in colors else ("medium" if "medium" in colors else "low")
        hits.append({
            "index": idx,
            "level": level,
            "chars": sum(colors.values()),
            "breakdown": colors,
            "text": text,
        })

    out = Path(args.out)
    out.write_text(json.dumps(hits, ensure_ascii=False, indent=2), encoding="utf-8")

    by_level = {}
    for h in hits:
        by_level[h["level"]] = by_level.get(h["level"], 0) + 1
    print(f"命中 {len(hits)} 段：{by_level}  → {out}")
    return 0


# --------------------------------------------------------------------------- #
def cmd_apply(args) -> int:
    src = Path(args.src)
    doc = Document(src)
    rules = json.loads(Path(args.rules).read_text(encoding="utf-8"))
    paras = doc.paragraphs

    ok, skipped = [], []
    # 先全部算一遍 before，再统一改
    for key, spec in sorted(rules.items(), key=lambda kv: int(kv[0])):
        idx = int(key)
        new_text = spec["new"] if isinstance(spec, dict) else spec
        match = spec.get("match") if isinstance(spec, dict) else None

        if idx >= len(paras):
            skipped.append((idx, "段落索引越界"))
            continue
        p = paras[idx]

        if match and not p.text.strip().startswith(match[:20]):
            skipped.append((idx, f"原文前缀不匹配，疑似文档已变动：{p.text[:24]!r}"))
            continue

        before = inspect_paragraph(p)
        try:
            msg = replace_paragraph(p, new_text, strict_numbers=not args.allow_number_change)
        except ReplaceError as e:
            skipped.append((idx, str(e)))
            continue

        issues = verify(before, inspect_paragraph(p))
        if issues:
            skipped.append((idx, "结构校验未通过：" + "；".join(issues)))
            continue
        ok.append((idx, msg))

    dst = Path(args.out) if args.out else src.with_suffix(".out.docx")
    doc.save(dst)

    print(f"成功 {len(ok)} 段，跳过 {len(skipped)} 段 → {dst}")
    for idx, reason in skipped:
        print(f"  跳过 段落{idx}: {reason}")
    return 0 if not skipped else 1


# --------------------------------------------------------------------------- #
def cmd_verify(args) -> int:
    a, b = Document(args.a), Document(args.b)
    if len(a.paragraphs) != len(b.paragraphs):
        print(f"段落数不同：{len(a.paragraphs)} → {len(b.paragraphs)}")
        return 1
    bad = 0
    for i, (pa, pb) in enumerate(zip(a.paragraphs, b.paragraphs)):
        issues = verify(inspect_paragraph(pa), inspect_paragraph(pb))
        if issues:
            bad += 1
            print(f"段落{i}: " + "；".join(issues))
    print(f"共 {len(a.paragraphs)} 段，结构异常 {bad} 段")
    return 0 if bad == 0 else 1


# --------------------------------------------------------------------------- #
def _build_fixture(path: Path) -> None:
    """构造一份带域、上标、首行缩进的测试文档。"""
    from docx import Document as D
    from docx.oxml import OxmlElement
    from docx.shared import Pt, Cm

    doc = D()
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0.85)
    pf.line_spacing = Pt(20)

    def add_run(text, sup=False):
        r = p.add_run(text)
        r.font.name = "宋体"
        r.font.size = Pt(12)
        if sup:
            r.font.superscript = True
        return r

    def add_field(result):
        """插入一个 REF 交叉引用域，结构为 begin/instrText/separate/result/end。"""
        el = p._element
        r1 = OxmlElement("w:r")
        fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), "begin")
        r1.append(fc); el.append(r1)

        r2 = OxmlElement("w:r")
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
        it.text = r" REF _Ref123 \r \h "
        r2.append(it); el.append(r2)

        r3 = OxmlElement("w:r")
        fc2 = OxmlElement("w:fldChar"); fc2.set(qn("w:fldCharType"), "separate")
        r3.append(fc2); el.append(r3)

        r4 = OxmlElement("w:r")
        t = OxmlElement("w:t"); t.text = result
        r4.append(t); el.append(r4)

        r5 = OxmlElement("w:r")
        fc3 = OxmlElement("w:fldChar"); fc3.set(qn("w:fldCharType"), "end")
        r5.append(fc3); el.append(r5)

    add_run("传统迭代算法的单次重建耗时可达到分钟级")
    add_field("[1]")
    add_run("，重建误差约为 10")
    add_run("-3", sup=True)
    add_run(" 量级。")

    # 空段 + 普通段，模拟真实文档
    doc.add_paragraph("")
    doc.add_paragraph("这一段没有域，用于对照。")
    doc.save(str(path))


def cmd_selftest(args) -> int:
    tmp = Path(args.workdir)
    tmp.mkdir(parents=True, exist_ok=True)
    src = tmp / "_fixture.docx"
    out = tmp / "_fixture.out.docx"
    _build_fixture(src)

    doc = Document(src)
    p = doc.paragraphs[0]
    before = inspect_paragraph(p)
    print("原文：", before["text"])
    print(f"  域 {before['field_count']} 个 / 上标 {before['vertalign_count']} 个 / "
          f"数字 {before['numbers']} / pPr={'有' if before['has_ppr'] else '无'}")

    new_text = ("传统迭代算法重建一张图像往往要花上分钟级的时间[1]，"
                "而重建误差通常停留在 10-3 量级。")
    msg = replace_paragraph(p, new_text)
    print("替换：", msg)

    after = inspect_paragraph(p)
    print("新文：", after["text"])
    issues = verify(before, after)
    print("结构校验：", "通过" if not issues else issues)

    # 落盘再读回，确认序列化后结构仍在
    doc.save(str(out))
    reloaded = inspect_paragraph(Document(out).paragraphs[0])
    issues2 = verify(before, reloaded)
    print("落盘重读后校验：", "通过" if not issues2 else issues2)

    # 负例 1：数字被改动 → 必须拒绝
    try:
        replace_paragraph(Document(src).paragraphs[0],
                          "传统迭代算法重建一张图像往往要花上分钟级的时间[1]，重建误差约 10-4 量级。")
        neg1 = "未拦截（缺陷）"
    except ReplaceError as e:
        neg1 = f"已拦截：{e}"

    # 负例 2：引用角标丢失 → 必须拒绝
    try:
        replace_paragraph(Document(src).paragraphs[0], "这段文字故意不带角标，重建误差 10-3 量级。")
        neg2 = "未拦截（缺陷）"
    except ReplaceError as e:
        neg2 = f"已拦截：{e}"

    print("负例·数字变动：", neg1)
    print("负例·角标丢失：", neg2)

    # 负例 3：token 齐全但数字集合变了 → 必须由数字校验拦下
    try:
        replace_paragraph(Document(src).paragraphs[0],
                          "传统迭代算法重建一张图像往往要花上 3 分钟[1]，重建误差停留在 10-3 量级。")
        neg3 = "未拦截（缺陷）"
    except ReplaceError as e:
        neg3 = f"已拦截：{e}"

    print("负例·数字新增：", neg3)

    passed = (not issues and not issues2
              and neg1.startswith("已拦截") and neg2.startswith("已拦截")
              and neg3.startswith("已拦截"))
    print("\n自检结果：", "全部通过" if passed else "存在失败项")
    return 0 if passed else 1


# --------------------------------------------------------------------------- #
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="域感知的 Word 整段替换工具")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("selftest", help="跑内置自检")
    s.add_argument("--workdir", default=".selftest")
    s.set_defaults(func=cmd_selftest)

    s = sub.add_parser("locate", help="按字体颜色定位风险段落")
    s.add_argument("path")
    s.add_argument("--out", default="flagged.json")
    s.set_defaults(func=cmd_locate)

    s = sub.add_parser("apply", help="批量安全替换")
    s.add_argument("src")
    s.add_argument("rules")
    s.add_argument("-o", "--out")
    s.add_argument("--allow-number-change", action="store_true",
                   help="允许新文本的数字集合与原文不同（默认禁止）")
    s.set_defaults(func=cmd_apply)

    s = sub.add_parser("verify", help="对比两份文档的结构完整性")
    s.add_argument("a")
    s.add_argument("b")
    s.set_defaults(func=cmd_verify)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
