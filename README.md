# 创策云 UniPlan AI 课程项目

本仓库用于《大学生创新创业与就业指导》课程小组作业。项目名称为“创策云 UniPlan AI：面向大学生创新创业项目的 AI 协作与商业计划书平台”，目标是完成商业计划书、问卷调研、PPT 汇报、团队分工和前端原型等课程交付物。

## 当前项目状态

项目已进入第三阶段，已从 Markdown 框架升级为可打印、可汇报的正式作业包。当前已完成：

- 商业计划书正文扩写与表格补充
- Word 版商业计划书导出
- PPT 文件导出
- 汇报讲稿导出
- 问卷星可复制版本
- 调研分析模板
- 团队分工与时间计划
- React + Vite 前端展示原型

## 目录结构

```text
.
├── assets/             # 图片、截图等素材占位
├── business-plan/      # 商业计划书大纲与正文
├── deliverables/       # 可提交 Word、PPT、讲稿和提交说明
├── docs/               # 项目简报、团队分工与时间计划
├── ppt-outline/        # PPT 大纲与完整页面内容
├── prototype/          # React + Vite 前端原型
├── research/           # 问卷、问卷星版本、调研分析模板
└── scripts/            # Word 和 PPT 导出脚本
```

## 原型运行方式

进入前端原型目录：

```bash
cd prototype
npm install
npm run dev
```

启动后根据终端提示打开本地地址，通常为：

```text
http://127.0.0.1:5173
```

## 导出 Word 和 PPT

在项目根目录运行：

```bash
python scripts/export_business_plan_docx.py
python scripts/export_pptx.py
```

导出结果位于：

- `deliverables/创策云UniPlanAI商业计划书.docx`
- `deliverables/创策云UniPlanAI课堂汇报.pptx`
- `deliverables/创策云UniPlanAI汇报讲稿.md`

## 最终交付物

- 商业计划书 Markdown：`business-plan/business_plan_draft.md`
- 商业计划书 Word：`deliverables/创策云UniPlanAI商业计划书.docx`
- PPT 内容 Markdown：`ppt-outline/slides_full_content.md`
- PPT 文件：`deliverables/创策云UniPlanAI课堂汇报.pptx`
- 汇报讲稿：`deliverables/创策云UniPlanAI汇报讲稿.md`
- 问卷调研：`research/questionnaire_wjx_ready.md`
- 调研分析模板：`research/survey_analysis_template.md`
- 前端原型：`prototype/`
- 团队分工表：`docs/team_work_plan.md`
- 提交说明：`deliverables/提交说明.md`

## 后续人工补充事项

- 团队成员学号、姓名
- 团队名称
- 指导教师
- 学院专业
- 提交日期
- 问卷星或腾讯问卷链接
- 问卷真实回收结果
- 商业计划书中所有【待问卷调研后补充】位置
- PPT 中可替换为真实原型截图或调研图表的位置

## 注意事项

- 不要编造问卷结果、精确市场规模、收入、利润或用户数。
- 需要真实数据的位置统一保留【待问卷调研后补充】。
- 不提交 `node_modules/` 和 `prototype/dist/`，依赖和构建产物可通过命令重新生成。
