# hr-tree · 组织架构树 / 汇报树生成

一个 Comate Skill：把含上下级关系的 Excel/CSV 一键做成**可交互的组织架构树网页**（零构建单文件 HTML，可直接发布分享）。

## 功能
- **Excel 导入**（.xlsx/.xls/.csv）或生成时把数据烘焙进页面；可下载带填写说明的 3 层示例模板。
- **字段设置**：主字段（节点名，恒显示）、上级字段（如"主管"，决定上下级）、节点标识列、层级/部门字段。
- **显示设置**：任意列可显隐；**聚合统计**（下级人数、求和/平均/最大最小、去重计数、分布，如 职级 P5×6 P6×3）。
- **层级控制**：下拉选展开深度，超深出现折叠条，单击局部展开。
- **搜索**：按名称/层级/部门高亮、自动展开路径、滚动定位。
- **交互**：单击展开折叠、双击聚焦、拖拽平移、横/纵布局切换、子节点排序。
- **编辑模式**：按住卡片拖动变更上级、新建/编辑/删除节点、导出含改动的 Excel。
- **保存版本**：存浏览器本地（localStorage），刷新/重开不丢，可恢复/删除；重开自动加载最新版本。
- **标题可改**；**导出 16:9 PPT（零依赖 OOXML+ZIP）/ PNG**。
- **示意图**<img width="2526" height="1472" alt="image" src="https://github.com/user-attachments/assets/6b88279e-6457-4350-8bc6-bcd9b3cf6cee" />


## 安装
Comate 用户把本仓库内容放到技能目录即可：

```bash
# 个人技能
git clone https://github.com/<你的用户名>/hr-tree.git ~/.comate/skills/hr-tree
```

或在 Comate「安装技能」对话框上传打包后的 `.zip`（根目录需含 `SKILL.md`）。

## 使用
- 对 Comate 说「生成汇报树 / 生成组织架构树」→ 在当前目录新建页面。
- 命令行：
  ```bash
  python3 scripts/bake_data.py --excel 数据.xlsx --inspect              # 查看表头与字段建议
  python3 scripts/new_page.py --excel 数据.xlsx --primary-field 姓名 --parent-field 主管
  python3 scripts/bake_data.py --dump-template 导入模板.xlsx            # 导出填写模板
  ```
- 生成的 HTML 双击即可用；首屏需联网加载 React/Babel（jsdelivr，失败自动切 unpkg），离线方案见 `references/data-format.md`。

## 结构
```
SKILL.md                     技能说明与触发规则
assets/org-chart-template.html  核心单文件 Web 应用
scripts/new_page.py          命令式新建页面
scripts/bake_data.py         Excel 烘焙 / 导出模板
references/data-format.md    数据格式、字段映射、离线方案
```

## 依赖
- 运行脚本：Python 3 + `openpyxl`（读 .xlsx）。
- 页面：React 18 / Babel / SheetJS / html2canvas（均走 CDN，按需加载）。

## License
MIT
