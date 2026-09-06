# 🎯 Math Modeling Skill (NO10JAMAL 增强版)

<div align="center">

**面向数学建模竞赛与建模项目的三阶段工作流**

[![Version](https://img.shields.io/badge/version-1.2.0--c132-blue.svg)](VERSION)
[![Agent Skill](https://img.shields.io/badge/Agent-Skill-4B8BBE.svg)](SKILL.md)

**请关注原作者**

[![CSDN](https://img.shields.io/badge/CSDN-博客-FC5531?logo=csdn&logoColor=white)](https://blog.csdn.net/SJbeITenginner?spm=1010.2135.3001.5343)
[![知乎](https://img.shields.io/badge/知乎-主页-0084FF?logo=zhihu&logoColor=white)](https://www.zhihu.com/people/27-85-7-72-95/posts)
[![小红书](https://img.shields.io/badge/小红书-主页-FF2442)](https://www.xiaohongshu.com/user/profile/6497dd69000000001c02ab98)

</div>

---

### 📌 致谢与上游开源溯源声明

本项目 Fork / 基于 [XiaoMaColtAI/math-modeling-skill](https://github.com/XiaoMaColtAI/math-modeling-skill) 进行深度二次开发与增强。诚挚感谢原作者 **[XiaoMaColtAI](https://github.com/XiaoMaColtAI)** 团队及其为数学建模开源社区贡献的优秀三阶段架构设计与 Agent Skill！

---

### 🚀 本分支（[NO10JAMAL/skill_math-model](https://github.com/NO10JAMAL/skill_math-model)）核心新增与优化内容：

#### 🏆 标杆国一版式全面对齐（以 2025 C132 国一论文为范本）：
- **大十段式自然顺延结构**：规范一级标题中文序号（`一、  问题重述`），各章自然顺延流转，取消机械分页。
- **首页与摘要规范**：三号黑体题目、小三黑体居中 `摘  要`（字间严格双空格）、正文各问加粗首句导引（如“**针对问题一，**”）、关键词双半角空格且彻底清除分号、摘要独立分页。
- **引入 OMML 公式双制表位居中**：基于页宽动态设置双制表位，公式严格居中、编号 `(1)` 紧贴右侧页边距，彻底摒弃隐形表格以消除 OOXML 结构与题注冲突。
- **规范紧凑图表与题注**：表格上方居中加粗黑体题注，五号字紧凑三线表；图片下方居中宋体题注；正文图表双向引用严格校验；附录表格与正文编号智能解耦。
- **附录呈现体系**：规范呈现附录 A（文件列表三线表）与附录 B（等宽 Consolas 8.5pt 带题注与行号代码清单）。

#### 🛡️ AI 风格审计与元数据脱敏工具链：
- **新增自动化审计脚本** [`ai_style_audit.py`](tools/docx/scripts/ai_style_audit.py)：检测文学化隐喻与高危套话禁词库，评估句长标准差（句式突发度检测，消除 AI 平滑特征）。
- **文档核心属性（`core_properties`）底层脱敏**：自动清理脚本及第三方生成工具签名，确保交付物纯净。

#### ⚙️ 工程化与规范文档更新：
- 完善 [`写作规范.md`](references/roles/论文手/references/写作规范.md)、[`章节模板.md`](references/roles/论文手/references/章节模板.md)、[`论文格式规范.md`](references/roles/论文手/references/论文格式规范.md)、[`tools/docx/SKILL.md`](tools/docx/SKILL.md)。
- 适配自动化测试套件 [`tools/docx/scripts/self_check.py`](tools/docx/scripts/self_check.py)，全套自检与 OOXML 验证均 100% 通过（PASSED）。
