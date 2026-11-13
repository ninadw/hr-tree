#!/usr/bin/env python3
"""把 Excel/CSV 数据烘焙进架构树模板，生成可直接打开（或发布到 popo）的单文件 HTML。

用法：
  python3 bake_data.py --excel 数据.xlsx --inspect
  python3 bake_data.py --excel 数据.xlsx --out ./architecture-tree.html \
      --primary-field 姓名 --parent-field 主管 [--id-field 工号] [--level-field 职级] \
      [--dept-field 部门] [--hide-fields 手机,邮箱 | --show-fields 职级,部门] [--depth 2|all] [--title 标题] \
      [--agg count] [--agg sum:HC] [--agg dist:职级:direct:直属职级]

依赖：openpyxl（读取 .xlsx/.xlsm）；.csv 无额外依赖（UTF-8，失败时回退 GBK）。
"""
import argparse
import csv
import datetime as dt
import json
import re
import sys
from pathlib import Path

INJECT_TOKEN = "/*__INJECT__*/ null"
DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "org-chart-template.html"

# 与页面内 autoMap 保持一致的识别规则（页面会再做一次更完整的识别）
RX = {
    "primary": [r"^(姓名|名字|员工姓名|人员姓名|名称|节点名称|组织名称|部门名称|name|full\s*name|employee\s*name)$", r"(姓名|名称|name)"],
    "parent": [r"^(主管|上级|直属上级|直属主管|直接上级|汇报对象|汇报人|上级领导|直属领导|上级部门|父节点|父级|manager|parent|leader|reports?\s*to|supervisor)$",
               r"(主管|上级|汇报|领导|经理|manager|parent|supervisor|reports?\s*to|父)"],
    "level": [r"^(层级|职级|级别|职位级别|level|grade|职位|岗位|title)$", r"(层级|职级|级别|level|grade|职位|岗位)"],
    "dept": [r"^(部门|所属部门|团队|所属团队|组织|所属组织|department|dept|team)$", r"(部门|团队|组织|department|dept|team)"],
}
ROOT_VALUES = {"", "0", "null", "none", "nil", "n/a", "-", "--", "无", "空"}
# 聚合统计：方式 → (名称, 是否需要字段)；范围取值见 AGG_SCOPES（与页面 OrgCore 保持一致）
AGG_OPS = {
    "count": ("下级人数", False), "nonEmpty": ("非空计数", True), "sum": ("求和", True), "avg": ("平均值", True),
    "max": ("最大值", True), "min": ("最小值", True), "distinct": ("去重计数", True), "dist": ("分布", True),
}
AGG_SCOPES = {"desc": "所有下级节点", "descSelf": "所有下级节点 + 自身", "direct": "仅直接下级节点"}
AGG_HELP = ("--agg 格式：方式[:字段][:范围][:显示名称]\n"
            "  方式：" + "，".join(f"{k}={v[0]}" for k, v in AGG_OPS.items()) + "\n"
            "  范围：" + "，".join(f"{k}={v}" for k, v in AGG_SCOPES.items()) + "（默认 desc）\n"
            "  示例：--agg count  --agg sum:HC  --agg dist:职级:direct:直属职级  --agg count::descSelf:团队人数")


def parse_agg(spec, headers):
    parts = spec.split(":")
    op = parts[0].strip()
    if op not in AGG_OPS:
        sys.exit(f"--agg 统计方式无效：{op}\n{AGG_HELP}")
    need_field = AGG_OPS[op][1]
    rest = parts[1:]
    field = ""
    if need_field:
        field = rest.pop(0).strip() if rest else ""
        if not field:
            sys.exit(f"--agg {spec}：方式 {op} 需要指定字段\n{AGG_HELP}")
        if field not in headers:
            sys.exit(f"--agg {spec}：表中不存在列 {field}\n可用列：{', '.join(headers)}")
    elif rest and rest[0].strip() not in AGG_SCOPES:
        rest.pop(0)  # count 允许写成 count::范围（字段位留空或被忽略）
    scope = rest.pop(0).strip() if rest else "desc"
    scope = scope or "desc"
    if scope not in AGG_SCOPES:
        sys.exit(f"--agg {spec}：统计范围无效：{scope}\n{AGG_HELP}")
    label = ":".join(rest).strip()
    return {"op": op, "field": field, "scope": scope, "label": label}


def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).replace("\u00a0", " ").strip()


def read_rows(path, sheet=None):
    """返回二维数组（字符串），首个非空行视为表头。"""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        for enc in ("utf-8-sig", "gbk"):
            try:
                with path.open(newline="", encoding=enc) as f:
                    return [[cell_str(c) for c in row] for row in csv.reader(f)]
            except UnicodeDecodeError:
                continue
        sys.exit("CSV 编码无法识别，请另存为 UTF-8 CSV 或 .xlsx")
    if suffix == ".xls":
        sys.exit("暂不支持 .xls，请在 Excel 中另存为 .xlsx 或 .csv 后重试")
    try:
        import openpyxl
    except ImportError:
        sys.exit("缺少 openpyxl：请先执行 pip install openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet and sheet not in wb.sheetnames:
        sys.exit(f"工作表不存在：{sheet}（可选：{', '.join(wb.sheetnames)}）")
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    rows = [[cell_str(c) for c in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def split_table(aoa):
    start = next((i for i, r in enumerate(aoa) if any(c for c in r)), None)
    if start is None:
        return [], []
    headers = [h or f"列{i + 1}" for i, h in enumerate(aoa[start])]
    body = [r for r in aoa[start + 1:] if any(c for c in r)]
    return headers, body


def suggest(headers):
    def pick(key, exclude):
        for rx in RX[key]:
            for h in headers:
                if h not in exclude and re.search(rx, h, re.I):
                    return h
        return ""
    primary = pick("primary", []) or (headers[0] if headers else "")
    parent = pick("parent", [primary])
    level = pick("level", [primary, parent])
    dept = pick("dept", [primary, parent, level])
    return {"primary": primary, "parent": parent, "level": level, "dept": dept}


def parent_hit_report(headers, body, primary, parent, id_field):
    """统计上级值能否在标识列中找到，帮助判断字段是否选对。"""
    pi, ki = headers.index(parent), headers.index(id_field or primary)
    keys = {r[ki] for r in body if ki < len(r) and r[ki]}
    vals = [r[pi] for r in body if pi < len(r) and r[pi].strip().lower() not in ROOT_VALUES]
    miss = [v for v in vals if v not in keys]
    return len(vals), miss


def detect_id_field(headers, body, primary, parent):
    """在除上级列外的列中，选能命中最多上级值的列作为标识列（同分优先主字段）。"""
    pi = headers.index(parent)
    vals = [r[pi] for r in body if pi < len(r) and r[pi].strip().lower() not in ROOT_VALUES]
    best, best_score = primary, -1
    for h in [primary] + [x for x in headers if x != primary]:
        if h == parent:
            continue
        hi = headers.index(h)
        keys = {r[hi] for r in body if hi < len(r) and r[hi]}
        score = sum(1 for v in vals if v in keys)
        if score > best_score:
            best, best_score = h, score
    return "" if best == primary else best


def csv_list(s):
    return [x.strip() for x in s.split(",") if x.strip()] if s else None


def save_template(path, sheet="人员"):
    """生成与页面"下载模板"一致的 .xlsx（表头 + 填写说明 + 3 行示例）。"""
    try:
        import openpyxl
    except ImportError:
        sys.exit("缺少 openpyxl：请先执行 pip install openpyxl")
    tpl = Path(DEFAULT_TEMPLATE).read_text(encoding="utf-8")
    # 从模板解析：表头(SAMPLE_AOA 第一行) + 填写说明(guides) + 示例(TEMPLATE_SAMPLE，含 3 层)
    ms = re.search(r"const SAMPLE_AOA = \[([\s\S]*?)\]\.map\(line => line\.split\('\|'\)\)", tpl)
    mg = re.search(r"const guides = \[([\s\S]*?)\];", tpl)
    mt = re.search(r"const TEMPLATE_SAMPLE = \[([\s\S]*?)\]\.map\(line => line\.split\('\|'\)\)", tpl)
    if not (ms and mg and mt):
        sys.exit("模板中未找到 SAMPLE_AOA/guides/TEMPLATE_SAMPLE，请检查 org-chart-template.html")
    header_line = next((x.strip().strip("'") for x in ms.group(1).split(",") if x.strip()), "")
    header_row = [c for c in header_line.split("|")]
    guide_row = [m.strip().strip("'") for m in re.findall(r"'((?:[^'\\]|\\.)*)'", mg.group(1))]
    sample_rows = [ln.strip().strip("'").split("|") for ln in re.findall(r"'((?:[^'\\]|\\.)*)'", mt.group(1))]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(header_row)
    ws.append(guide_row)
    for row in sample_rows:
        ws.append([c.replace("\\u00a0", " ") for c in row])
    if len(header_row) != len(guide_row) or len(header_row) < 2:
        sys.exit(f"模板行解析失败：表头 {len(header_row)} 列 / 说明 {len(guide_row)} 列，请检查 org-chart-template.html")
    from openpyxl.styles import Font, PatternFill
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    for c in ws[2]:
        c.font = Font(italic=True, color="808080")
    for i, h in enumerate(header_row, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(10, len(h) * 2 + 2,
            max((len(str(ws.cell(row=r, column=i).value or "")) + 2) for r in range(2, ws.max_row + 1)))
    ws.freeze_panes = "A2"
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="把 Excel 数据烘焙进架构树 HTML 模板")
    ap.add_argument("--excel", help="Excel(.xlsx/.xlsm) 或 CSV 文件路径（--dump-template 时可省略）")
    ap.add_argument("--sheet", help="工作表名称，默认第一个")
    ap.add_argument("--inspect", action="store_true", help="只打印表头、样例行与字段建议")
    ap.add_argument("--out", default="architecture-tree.html", help="输出 HTML 路径")
    ap.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="模板路径（默认 skill 自带模板）")
    ap.add_argument("--primary-field", help="主字段（节点名称）")
    ap.add_argument("--parent-field", help="上级字段（上下节点关系）")
    ap.add_argument("--id-field", help="节点标识列：上级字段的值对应的列，默认与主字段相同")
    ap.add_argument("--level-field", help="层级字段（用于搜索）")
    ap.add_argument("--dept-field", help="部门字段（用于搜索）")
    ap.add_argument("--hide-fields", help="节点上隐藏的字段，逗号分隔（其余全部显示）")
    ap.add_argument("--show-fields", help="节点上显示的字段，逗号分隔（优先于 --hide-fields）")
    ap.add_argument("--show-labels", action="store_true", help="字段值前显示字段名")
    ap.add_argument("--depth", default="2", help="默认展开层数 1-6 或 all，默认 2")
    ap.add_argument("--title", help="页面标题，默认取文件名")
    ap.add_argument("--agg", action="append", default=[], metavar="方式[:字段][:范围][:名称]",
                    help="聚合统计字段，可重复；如 count、sum:HC、dist:职级:direct（--inspect 查看全部取值）")
    ap.add_argument("--dump-template", metavar="路径", help="把页面内下载的导入模板另存为 .xlsx（无需 --excel）")
    args = ap.parse_args()

    if args.dump_template:
        save_template(args.dump_template)
        print(f"已导出导入模板：{Path(args.dump_template).resolve()}")
        return

    if not args.excel:
        sys.exit("缺少 --excel（导出模板请使用 --dump-template 路径）")
    src = Path(args.excel).expanduser()
    if not src.exists():
        sys.exit(f"文件不存在：{src}")
    aoa = read_rows(src, args.sheet)
    headers, body = split_table(aoa)
    if not headers or not body:
        sys.exit("未解析到有效数据：请确认第一行为表头且至少有一行数据")
    auto = suggest(headers)

    if args.inspect:
        print(f"文件：{src.name}　数据行：{len(body)}")
        print("表头：" + " | ".join(headers))
        for r in body[:3]:
            print("样例：" + " | ".join(r[:len(headers)]))
        print("字段建议：" + json.dumps(auto, ensure_ascii=False))
        if auto["primary"] and auto["parent"]:
            id_field = detect_id_field(headers, body, auto["primary"], auto["parent"])
            if id_field:
                print(f"标识列建议：{id_field}（上级字段的值对应此列，烘焙时会自动使用）")
            total, miss = parent_hit_report(headers, body, auto["primary"], auto["parent"], id_field)
            print(f"上级值命中：{total - len(miss)}/{total}" + (f"（未命中示例：{', '.join(miss[:5])}，这些节点会成为顶层）" if miss else ""))
        print(AGG_HELP)
        return

    field_map = {
        "primary": args.primary_field or auto["primary"],
        "parent": args.parent_field or auto["parent"],
        "id": args.id_field or "",
        "level": args.level_field if args.level_field is not None else auto["level"],
        "dept": args.dept_field if args.dept_field is not None else auto["dept"],
    }
    unknown = [v for v in field_map.values() if v and v not in headers]
    for lst in (csv_list(args.hide_fields) or []) + (csv_list(args.show_fields) or []):
        if lst not in headers:
            unknown.append(lst)
    if unknown:
        sys.exit(f"表中不存在列：{', '.join(unknown)}\n可用列：{', '.join(headers)}")
    if not field_map["primary"] or not field_map["parent"]:
        print("提示：未能确定主字段或上级字段，页面打开后会弹出字段设置让用户补齐。", file=sys.stderr)
    elif field_map["primary"] == field_map["parent"]:
        sys.exit("主字段与上级字段不能是同一列")
    elif not field_map["id"]:
        field_map["id"] = detect_id_field(headers, body, field_map["primary"], field_map["parent"])

    depth = args.depth.strip().lower()
    if depth != "all" and not (depth.isdigit() and 1 <= int(depth) <= 6):
        sys.exit("--depth 只能是 1-6 或 all")

    display = {"showLabels": args.show_labels}
    if args.show_fields:
        display["visibleFields"] = csv_list(args.show_fields)
    elif args.hide_fields:
        display["hiddenFields"] = csv_list(args.hide_fields) + [f for f in (field_map["parent"], field_map["id"]) if f]

    payload = {
        "aoa": [headers] + [r[:len(headers)] for r in body],
        "fileName": src.name,
        "title": args.title or src.stem,
        "fieldMap": field_map,
        "displayConfig": display,
        "maxDepth": "all" if depth == "all" else int(depth),
        "aggregates": [parse_agg(s, headers) for s in args.agg],
    }
    tpl_path = Path(args.template).expanduser()
    tpl = tpl_path.read_text(encoding="utf-8")
    if tpl.count(INJECT_TOKEN) != 1:
        sys.exit(f"模板中未找到唯一的注入占位符 {INJECT_TOKEN}：{tpl_path}")
    # 防止数据中的 "</script>" 提前结束脚本块
    data_js = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    out = Path(args.out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(tpl.replace(INJECT_TOKEN, data_js), encoding="utf-8")

    msg = f"已生成：{out.resolve()}（{len(body)} 行，主字段={field_map['primary']}，上级={field_map['parent']}"
    msg += f"，标识列={field_map['id']}）" if field_map["id"] else "）"
    print(msg)
    for a in payload["aggregates"]:
        print(f"聚合统计：{AGG_OPS[a['op']][0]}" + (f"·{a['field']}" if a["field"] else "") + f"（{AGG_SCOPES[a['scope']]}）" + (f" → {a['label']}" if a["label"] else ""))
    if field_map["primary"] and field_map["parent"]:
        total, miss = parent_hit_report(headers, body, field_map["primary"], field_map["parent"], field_map["id"])
        if miss:
            print(f"注意：{len(miss)}/{total} 个上级值在标识列中找不到（如 {', '.join(miss[:5])}），这些节点会成为顶层节点。")


if __name__ == "__main__":
    main()
