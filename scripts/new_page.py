#!/usr/bin/env python3
"""architecture-tree / new_page：在当前目录新建一个架构树页面（唯一文件名，可选烘焙 Excel）。

用法：
  python3 new_page.py                                   # 当前目录生成空模板页
  python3 new_page.py --excel 人员.xlsx --primary-field 姓名 --parent-field 主管
  python3 new_page.py --cwd ~/报告 --title "Q4 组织架构" --depth all
  python3 new_page.py --out 汇报树.html                 # 指定文件名（已存在则报错，不覆盖）

带 --excel 时委托同目录的 bake_data.py 注入数据。
"""
import argparse
import datetime
import shutil
import subprocess
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = SKILL_ROOT / "assets" / "org-chart-template.html"
BAKE = SKILL_ROOT / "scripts" / "bake_data.py"


def unique_path(directory, base="汇报树", ext=".html"):
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate = directory / f"{base}-{ts}{ext}"
    n = 2
    while candidate.exists():
        candidate = directory / f"{base}-{ts}-{n}{ext}"
        n += 1
    return candidate


def resolve_input(path_str, cwd):
    """相对路径依次按 当前目录 / cwd 解析。"""
    p = Path(path_str).expanduser()
    if p.is_absolute():
        return p if p.exists() else None
    for base in (Path.cwd(), cwd):
        q = (base / p)
        if q.exists():
            return q.resolve()
    return None


def main():
    ap = argparse.ArgumentParser(description="新建一个架构树页面（可选烘焙 Excel）")
    ap.add_argument("--cwd", default=".", help="输出目录，默认当前目录")
    ap.add_argument("--out", help="输出文件名或完整路径（默认 汇报树-时间戳.html，自动避免覆盖）")
    ap.add_argument("--excel", help="要烘焙进页面的 Excel/CSV 文件")
    ap.add_argument("--title", help="页面标题（烘焙时默认取 Excel 文件名）")
    ap.add_argument("--primary-field", help="主字段（节点名称）")
    ap.add_argument("--parent-field", help="上级字段（上下节点关系）")
    ap.add_argument("--id-field", help="节点标识列（上级值对应的列）")
    ap.add_argument("--level-field", help="层级字段（搜索用）")
    ap.add_argument("--dept-field", help="部门字段（搜索用）")
    ap.add_argument("--depth", help="默认展开层数 1-6 或 all，默认 2")
    ap.add_argument("--hide-fields", help="节点上隐藏的字段，逗号分隔")
    ap.add_argument("--show-fields", help="节点上显示的字段，逗号分隔")
    ap.add_argument("--show-labels", action="store_true", help="字段值前显示字段名")
    ap.add_argument("--agg", action="append", default=[], metavar="方式[:字段][:范围][:名称]",
                    help="聚合统计字段，可重复；如 count、sum:HC")
    args = ap.parse_args()

    if not TEMPLATE.exists() or not BAKE.exists():
        sys.exit(f"未找到模板（{TEMPLATE}）。本 skill 安装不完整，请重新安装 architecture-tree skill。")

    cwd = Path(args.cwd).expanduser().resolve()
    cwd.mkdir(parents=True, exist_ok=True)

    if args.out:
        out = Path(args.out).expanduser()
        if not out.is_absolute():
            out = cwd / out
        out.parent.mkdir(parents=True, exist_ok=True)
        if out.exists():
            sys.exit(f"目标文件已存在：{out}（为避免覆盖已停止；换个文件名或先删除）")
    else:
        out = unique_path(cwd)

    if not args.excel:
        shutil.copyfile(TEMPLATE, out)
        print(f"已新建页面：{out}")
        print("（空模板：用浏览器打开后点“导入 Excel”即可生成树）")
        return

    excel = resolve_input(args.excel, cwd)
    if excel is None:
        sys.exit(f"Excel 文件不存在：{args.excel}")

    cmd = [sys.executable, str(BAKE), "--excel", str(excel), "--out", str(out)]
    if args.title:
        cmd += ["--title", args.title]
    for flag, value in (("--primary-field", args.primary_field), ("--parent-field", args.parent_field),
                        ("--id-field", args.id_field), ("--level-field", args.level_field),
                        ("--dept-field", args.dept_field), ("--depth", args.depth),
                        ("--hide-fields", args.hide_fields), ("--show-fields", args.show_fields)):
        if value:
            cmd += [flag, value]
    if args.show_labels:
        cmd.append("--show-labels")
    for a in args.agg:
        cmd += ["--agg", a]

    ret = subprocess.run(cmd)
    if ret.returncode != 0:
        sys.exit(ret.returncode)
    print(f"已新建页面：{out}")


if __name__ == "__main__":
    main()