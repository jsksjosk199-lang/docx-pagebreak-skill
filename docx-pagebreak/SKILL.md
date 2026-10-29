---
name: docx-pagebreak
description: Use when a Chinese work or student-organization monthly report in Word/WPS needs formatting against supplied rules, real page breaks between sections, or repair of overlapping page-break markers and blank pages.
---

# 月报 Word 排版与分页

处理月报 `.docx` 的格式、表格和分页。先确定用户要“只修分页”还是“按要求完整排版”；仅修分页时保留已有文字、字号、列宽、行高和人员信息。收到聊天记录或多个文件时，先按聊天时间与明确说明找到最新已修改版本，不用旧稿覆盖。

## 工作方式

1. 阅读原稿和当次提供的规范，核对章节、事件数、现有分页、浮动表格，以及姓名、月份和汇报人。不要编造、改写事实或推断项目负责人。只有用户要求沿用本次学生会月报规范时，才读取 [references/student-union-format.md](references/student-union-format.md)；新规范优先。
2. 按任务范围编辑副本。每四项一页属于该示例规范，不自动加到其他单位的月报。明确需要分开的部分之间插入真正的手动分页符，不用连续回车或仅靠自动换页代替。保留已有正确分页，避免重复；不要为了显示标记在正文中输入“分页符”。
3. 对“三张独立表格：本月完成工作／下月工作计划／问题复盘与总结”的已识别版式，可用 [scripts/section_breaks.py](scripts/section_breaks.py) 自动检查并修复两个栏目边界。它不会排版字体或处理一节超过四项的内部分页。其他结构、重名栏目、夹有正文或复杂对象的边界，先分析再定向处理，不能强行套用。
4. 当分页符出现在标题上或标记重影，检查分页符是否嵌在标题单元格、后一张表是否浮动并锚定到分页段落。经检查需要时把该表改成嵌入型，保持表格内容和宽度。用户需要标记居中时，可试 `--wps-center`，这是本机 WPS 的经验值；实际效果必须在目标 WPS 核对。
5. 在 Word/WPS 检查每页内容和两个分页符：无标题遮挡、内容截断或多余空白页；结构检查确认手动分页符所在位置，不能以页数正确代替。若标记段落挤出空白页，在不违背页边距规范时可适度调整底边距，并在交付时说明改动；若有严格尺寸冲突，优先保留全部内容并明确说明未满足的尺寸，不能宣称完全合规。
6. WPS 要看到分页标记，需要在**打开文件的设备**开启“开始 → 显示/隐藏段落标记”；如果用户要求三个开关都勾选，再核实“段落标记、段落布局按钮、段落柄按钮”。这属于本机显示设置，不能承诺会随文件传给别人。检查后是否恢复原设置遵从用户要求。未能在 WPS 实测就明确说明，不声称已验证显示效果。

交付可编辑 Word 成品，说明做了哪些必要调整与尚需核实的人员或规范冲突；不自动发送、上传或提交。

## 辅助脚本

用工作区依赖加载器提供的 Python 运行，确保有 `python-docx`：

```text
python scripts/section_breaks.py inspect "原稿.docx"
python scripts/section_breaks.py repair "原稿.docx" --out "新副本.docx" --boundaries both --wps-center --inline-floating
```

以上路径相对于本 Skill 目录。Windows 控制台出现中文乱码时，可为当前命令设置 `PYTHONIOENCODING=utf-8` 并将控制台输出编码设为 UTF-8，不修改系统设置。

先查看 `inspect` 的栏目、浮动表格及各边界的 `supported`；退出码为 0 不代表所有边界都可自动修复。根据用户范围选择 `--boundaries month-plan`、`plan-review` 或 `both`，不要因修一处而动另一处。`repair` 拒绝覆盖原件或已有输出，碰到不匹配的结构会报错并不写结果。仅当验证发现分页段落挤出空白页且规范允许时，才添加 `--bottom-margin-cm 1.6` 之类的参数；记录原值和新值。完成后仍须在 Word/WPS 进行版面与标记检查。可用 `--help` 查看参数。
