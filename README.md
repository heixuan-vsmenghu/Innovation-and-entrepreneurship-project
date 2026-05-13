# 创策云 UniPlan AI 课程项目

本仓库用于《大学生创新创业与就业指导》课程小组作业。项目名称为“创策云 UniPlan AI：面向大学生创新创业项目的 AI 协作与商业计划书平台”，目标是完成商业计划书、问卷调研、PPT 汇报、团队分工和前端原型等课程交付物。

## 目录结构

```text
.
├── assets/             # 图片、截图等素材占位
├── business-plan/      # 商业计划书大纲与正文
├── docs/               # 项目简报、团队分工与时间计划
├── ppt-outline/        # PPT 大纲与完整页面内容
├── prototype/          # React + Vite 前端原型
├── research/           # 问卷、问卷星版本、调研分析模板
└── scripts/            # 后续脚本占位
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

## 最终交付物

- 商业计划书：`business-plan/business_plan_draft.md`
- PPT 内容：`ppt-outline/slides_full_content.md`
- 问卷调研：`research/questionnaire_wjx_ready.md`
- 调研分析模板：`research/survey_analysis_template.md`
- 前端原型：`prototype/`
- 团队分工表：`docs/team_work_plan.md`

## 注意事项

- 市场规模、用户比例、收入预测等真实数据需在问卷回收后补充。
- 当前文档中的【待问卷调研后补充】为后续替换位置。
- 不提交 `node_modules/` 和 `prototype/dist/`，依赖和构建产物可通过命令重新生成。
