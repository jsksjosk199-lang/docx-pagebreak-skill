# Word 分页修复与月报排版 Skill

一个帮助 Codex 处理中文工作月报的 Skill。把 Word 月报和格式要求交给 Codex，它会根据你的要求检查表格、分页和版式，保留原稿信息，并交付可编辑的副本。

这个项目尤其适合学生组织和部门月报：有时正文已经换页，但 WPS 中的“分页符”标记压在下一张表格标题上；有时需要分开本月工作、下月计划和问题复盘。本 Skill 会引导 Codex 检查真正的分页符、表格浮动方式和页面显示，再在目标软件中核对结果。

## 能做什么

- 根据本次提供的规范检查或调整字体、段落、表格与分页；也可以只修分页，保留其他格式。
- 对能唯一识别的三表结构，使用附带脚本检查和修复栏目之间的手动分页符。
- 协助排查浮动表格、分页标记重影、多余空白页及长文字被固定行高截断的问题。
- 保留原稿，默认另存副本；不自行发送、上传或提交用户文档。

脚本处理的是 **“本月完成工作 / 下月工作计划 / 问题复盘与总结”三张独立表格之间的边界**，不会自动完成所有排版，也不会自动拆分一个栏目内超过四项的内容。识别不到唯一标题、边界夹有正文或复杂对象时，脚本会拒绝自动修改，需由 Codex 检查后采用其他方式处理。提供的学生会规范是一个可选示例，只有用户要求沿用时才使用；不同组织以当次提供的要求为准。

## 安装到 Codex

克隆仓库后，将其中的 `docx-pagebreak` 文件夹复制到 Codex 的技能目录 `~/.codex/skills/`。如果设置了 `CODEX_HOME`，请放到 `$CODEX_HOME/skills/`。重启 Codex 或重新打开会话后，技能可被发现。

Windows PowerShell 示例：

```powershell
git clone https://github.com/jsksjosk199-lang/docx-pagebreak-skill.git
$codexSkills = if ($env:CODEX_HOME) { Join-Path $env:CODEX_HOME 'skills' } else { Join-Path $HOME '.codex\skills' }
New-Item -ItemType Directory -Force -Path $codexSkills | Out-Null
Copy-Item -Recurse 'docx-pagebreak-skill\docx-pagebreak' $codexSkills
```

如目标目录已有同名技能，请先备份或检查差异，再决定是否替换。辅助脚本需要 Python 3 和 `python-docx`；本项目使用 `python-docx` 1.2.0 测试。Codex 可使用其提供的 Python 运行环境，也可以在自己的环境中安装：

```bash
python -m pip install "python-docx>=1.2.0,<1.3.0"
```

## 使用方式

安装后，把自己的月报和格式要求交给 Codex，例如：

> 用 `docx-pagebreak` 处理这份月报，按我给的图片规范排版，并检查真正的分页符。

> 只修复本月工作和下月计划之间的分页，其他格式和内容保持不变。

如果有多个版本，说明哪个是最新稿。人员名单、负责人、月份和活动内容以用户提供的材料为准，技能不会凭空补全。涉及隐私的文档留在使用者自己的工作环境中；是否通过某个服务处理，由使用者选择该服务的配置。

## 手动运行辅助脚本

在仓库目录运行，先检查文档结构：

```bash
python docx-pagebreak/scripts/section_breaks.py inspect "原稿.docx"
```

确认需要处理的边界受支持后，选择 `month-plan`、`plan-review` 或 `both`，并指定一个不存在的输出文件：

```bash
python docx-pagebreak/scripts/section_breaks.py repair "原稿.docx" --out "处理后.docx" --boundaries both
```

需要排查浮动表格遮挡时，可按实际情况使用 `--inline-floating`。`--wps-center` 是针对一个本机 WPS 版本调试的分页标记居中经验值，不保证在其他设备上有效。`--bottom-margin-cm` 会改变页边距，仅在规范允许且确有必要时使用。用 `--help` 查看参数；`inspect` 正常退出并不表示所有边界都支持自动修复，需检查输出中的 `supported`。

处理后请在目标 Word/WPS 中检查页数、分页符位置、文字是否截断、表格是否遮挡以及有无空白页。WPS 的“显示/隐藏段落标记”是设备上的显示设置，不随文档保证同步；打开该设置后才能看到相应标记。

## 测试与适用范围

脚本测试使用合成文档，不包含真实月报或姓名：

```bash
python -m unittest discover -s tests -v
```

本项目的分页处理依据一个实际月报案例开发，并对该案例的版式及结构进行过检查。不同 Word/WPS 版本、模板、字体和打印设置可能导致显示差异；完整排版和其他模板仍应逐页验证。

## 许可

Copyright 2026 jsksjosk199-lang。按 [Apache License 2.0](LICENSE) 开源，版权及归属通知见 [NOTICE](NOTICE)。技能目录也附带同样的许可和归属文件，便于单独复制时保留。该协议允许商业使用和修改；分发时须遵守协议中的版权、许可、NOTICE 和修改说明等要求。`python-docx` 等第三方依赖遵循各自的许可。
