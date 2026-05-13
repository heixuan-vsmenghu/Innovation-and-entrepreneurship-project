import {
  BarChart3,
  BookOpenText,
  BrainCircuit,
  BriefcaseBusiness,
  CalendarCheck,
  CheckCircle2,
  ClipboardList,
  Clock3,
  FilePenLine,
  Files,
  LayoutDashboard,
  ListChecks,
  MessageSquareText,
  Presentation,
  Rocket,
  SearchCheck,
  ShieldCheck,
  Sparkles,
  UsersRound,
} from 'lucide-react';
import { useMemo, useState } from 'react';

const navItems = [
  { id: 'home', label: '首页', icon: LayoutDashboard },
  { id: 'features', label: '功能介绍', icon: Sparkles },
  { id: 'workspace', label: '项目工作台', icon: BriefcaseBusiness },
  { id: 'plan', label: '商业计划书生成', icon: FilePenLine },
  { id: 'tasks', label: '任务看板', icon: ListChecks },
];

const stats = [
  { label: '课程交付物', value: '6 项', note: '计划书 / PPT / 问卷 / 原型', icon: Files },
  { label: '计划书章节', value: '16 章', note: '按课程概要格式整理', icon: BookOpenText },
  { label: '小组任务', value: '12 条', note: '静态模拟进度', icon: CalendarCheck },
  { label: '调研数据', value: '待补充', note: '问卷回收后录入', icon: BarChart3 },
];

const featureCards = [
  {
    title: '选题分析',
    text: '围绕专业方向、课程要求和团队能力生成可讨论的项目方向。',
    icon: BrainCircuit,
    tag: '选题',
  },
  {
    title: '资料整理',
    text: '把政策背景、竞品资料、访谈记录和参考文献集中到项目资料库。',
    icon: SearchCheck,
    tag: '调研',
  },
  {
    title: '问卷调研',
    text: '生成问卷结构、题目建议和结果分析提示，便于后续补充真实数据。',
    icon: ClipboardList,
    tag: '问卷',
  },
  {
    title: '计划书生成',
    text: '按商业计划书章节输出大纲、草稿和待补充数据提醒。',
    icon: FilePenLine,
    tag: '写作',
  },
  {
    title: 'PPT 大纲',
    text: '自动拆分路演页标题、核心内容、图示建议和讲稿提示。',
    icon: Presentation,
    tag: '汇报',
  },
  {
    title: '团队协作',
    text: '用任务看板记录分工、状态、负责人和截止时间。',
    icon: UsersRound,
    tag: '协作',
  },
];

const planSections = [
  { title: '封面', status: '待补充成员信息', progress: 30 },
  { title: '执行概要', status: '已生成草稿', progress: 80 },
  { title: '项目背景与市场痛点', status: '需补充问卷结果', progress: 58 },
  { title: '产品与服务', status: '已完成概要', progress: 76 },
  { title: '市场分析', status: '待补充数据', progress: 42 },
  { title: '商业模式设计', status: '已完成框架', progress: 70 },
  { title: '融资与财务计划', status: '待测算', progress: 36 },
  { title: '风险分析与应对', status: '已完成概要', progress: 72 },
];

const taskColumns = [
  {
    title: '待办',
    tone: 'slate',
    tasks: [
      { title: '补充团队成员学号姓名', owner: '全体', due: '本周', priority: '高' },
      { title: '发放问卷并收集样本', owner: '调研组', due: '待定', priority: '高' },
      { title: '整理竞品分析表', owner: '产品组', due: '待定', priority: '中' },
    ],
  },
  {
    title: '进行中',
    tone: 'blue',
    tasks: [
      { title: '完善商业计划书概要', owner: '写作组', due: '本周', priority: '高' },
      { title: '制作 PPT 视觉页', owner: '设计组', due: '待定', priority: '中' },
      { title: '准备课堂演示流程', owner: '汇报组', due: '待定', priority: '中' },
    ],
  },
  {
    title: '已完成',
    tone: 'green',
    tasks: [
      { title: '确定项目名称与定位', owner: '项目负责人', due: '已完成', priority: '高' },
      { title: '建立项目目录结构', owner: '技术组', due: '已完成', priority: '中' },
      { title: '生成问卷初稿', owner: '调研组', due: '已完成', priority: '中' },
    ],
  },
];

const versions = [
  { name: 'V0.3 原型展示版', time: '课堂展示前', note: '五个页面可演示' },
  { name: 'V0.2 计划书概要版', time: '当前阶段', note: '章节结构已确认' },
  { name: 'V0.1 项目设想版', time: '立项阶段', note: '确定项目方向' },
];

function App() {
  const [activePage, setActivePage] = useState('home');
  const [selectedChapter, setSelectedChapter] = useState('执行概要');

  const activeLabel = useMemo(
    () => navItems.find((item) => item.id === activePage)?.label ?? '首页',
    [activePage],
  );

  return (
    <div className="app-shell">
      <aside className="sidebar" aria-label="主导航">
        <div className="brand">
          <div className="brand-mark">U</div>
          <div>
            <strong>创策云</strong>
            <span>UniPlan AI</span>
          </div>
        </div>

        <nav className="nav-list">
          {navItems.map((item) => {
            const Icon = item.icon;
            return (
              <button
                className={activePage === item.id ? 'nav-item active' : 'nav-item'}
                key={item.id}
                onClick={() => setActivePage(item.id)}
                type="button"
                title={item.label}
              >
                <Icon size={18} />
                <span>{item.label}</span>
              </button>
            );
          })}
        </nav>

        <div className="sidebar-note">
          <ShieldCheck size={18} />
          <span>AI 结果需人工核验，市场数据以问卷和访谈为准。</span>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div>
            <p className="eyebrow">大学生创新创业项目协作平台</p>
            <h1>{activeLabel}</h1>
          </div>
          <div className="top-actions">
            <button className="icon-button" type="button" title="项目消息">
              <MessageSquareText size={18} />
            </button>
            <button className="primary-action" type="button">
              <Rocket size={18} />
              生成汇报材料
            </button>
          </div>
        </header>

        {activePage === 'home' && <HomePage />}
        {activePage === 'features' && <FeaturesPage />}
        {activePage === 'workspace' && <WorkspacePage />}
        {activePage === 'plan' && (
          <PlanPage selectedChapter={selectedChapter} setSelectedChapter={setSelectedChapter} />
        )}
        {activePage === 'tasks' && <TasksPage />}
      </main>
    </div>
  );
}

function HomePage() {
  return (
    <div className="page-stack">
      <section className="hero-panel">
        <div>
          <p className="eyebrow">课程展示原型</p>
          <h2>让小组项目从选题到路演都有清晰流程</h2>
          <p>
            创策云 UniPlan AI 聚焦大学生创新创业课程场景，用结构化工作台连接调研、
            商业计划书、PPT 大纲与任务分工。
          </p>
          <div className="hero-actions">
            <button className="primary-action" type="button">
              <Sparkles size={18} />
              模拟生成计划书
            </button>
            <button className="secondary-action" type="button">
              <ClipboardList size={18} />
              查看交付清单
            </button>
          </div>
        </div>
        <div className="hero-board" aria-label="项目状态概览">
          <div className="board-header">
            <span>项目进度</span>
            <strong>62%</strong>
          </div>
          <div className="progress-track">
            <span style={{ width: '62%' }} />
          </div>
          <div className="mini-grid">
            <div>
              <span>问卷</span>
              <strong>待发放</strong>
            </div>
            <div>
              <span>PPT</span>
              <strong>14 页</strong>
            </div>
            <div>
              <span>计划书</span>
              <strong>概要版</strong>
            </div>
            <div>
              <span>原型</span>
              <strong>可演示</strong>
            </div>
          </div>
        </div>
      </section>

      <section className="stats-grid" aria-label="项目指标">
        {stats.map((item) => {
          const Icon = item.icon;
          return (
            <article className="stat-card" key={item.label}>
              <Icon size={20} />
              <span>{item.label}</span>
              <strong>{item.value}</strong>
              <p>{item.note}</p>
            </article>
          );
        })}
      </section>

      <section className="two-column">
        <div className="panel">
          <div className="section-heading">
            <h3>课程交付物</h3>
            <span>打印与汇报准备</span>
          </div>
          <ul className="deliverable-list">
            {['商业计划书概要', '封面成员信息', '问卷调研设计', 'PPT 汇报大纲', '前端原型', '小组分工表'].map(
              (item) => (
                <li key={item}>
                  <CheckCircle2 size={18} />
                  <span>{item}</span>
                </li>
              ),
            )}
          </ul>
        </div>
        <div className="panel">
          <div className="section-heading">
            <h3>最近动态</h3>
            <span>静态模拟记录</span>
          </div>
          <div className="activity-list">
            <Activity time="09:20" text="生成商业计划书 16 章概要目录" />
            <Activity time="10:05" text="整理问卷题目，控制在 15 题以内" />
            <Activity time="10:40" text="创建课堂展示用前端页面" />
          </div>
        </div>
      </section>
    </div>
  );
}

function FeaturesPage() {
  return (
    <div className="page-stack">
      <section className="intro-band">
        <div>
          <p className="eyebrow">功能地图</p>
          <h2>围绕课程项目全过程设计</h2>
        </div>
        <p>
          平台把 AI 生成能力嵌入项目流程，所有市场规模、调研结论和财务测算都保留人工补充入口。
        </p>
      </section>

      <section className="feature-grid">
        {featureCards.map((item) => {
          const Icon = item.icon;
          return (
            <article className="feature-card" key={item.title}>
              <div className="feature-topline">
                <Icon size={22} />
                <span>{item.tag}</span>
              </div>
              <h3>{item.title}</h3>
              <p>{item.text}</p>
            </article>
          );
        })}
      </section>

      <section className="flow-band">
        {['创建项目', '录入资料', 'AI 生成大纲', '小组协作修改', '导出课堂材料'].map((step, index) => (
          <div className="flow-step" key={step}>
            <span>{index + 1}</span>
            <strong>{step}</strong>
          </div>
        ))}
      </section>
    </div>
  );
}

function WorkspacePage() {
  return (
    <div className="workspace-layout">
      <section className="panel project-summary">
        <div className="section-heading">
          <h3>项目概览</h3>
          <span>创策云 UniPlan AI</span>
        </div>
        <div className="summary-row">
          <span>项目定位</span>
          <strong>高校创新创业项目 AI 协作平台</strong>
        </div>
        <div className="summary-row">
          <span>目标用户</span>
          <strong>课程小组 / 竞赛团队 / 指导教师</strong>
        </div>
        <div className="summary-row">
          <span>当前阶段</span>
          <strong>课程原型与商业计划书概要</strong>
        </div>
        <div className="tag-row">
          <span>AI 写作</span>
          <span>任务看板</span>
          <span>问卷调研</span>
          <span>版本管理</span>
        </div>
      </section>

      <section className="panel">
        <div className="section-heading">
          <h3>资料库</h3>
          <span>待补充真实调研</span>
        </div>
        <div className="resource-list">
          <Resource icon={ClipboardList} title="问卷调研" text="题目已设计，样本数量待补充数据" />
          <Resource icon={SearchCheck} title="竞品分析" text="通用 AI、在线文档、项目管理工具" />
          <Resource icon={BookOpenText} title="课程格式" text="按创业计划书概要章节整理" />
        </div>
      </section>

      <section className="panel wide-panel">
        <div className="section-heading">
          <h3>版本记录</h3>
          <span>关键节点留痕</span>
        </div>
        <div className="version-list">
          {versions.map((version) => (
            <div className="version-item" key={version.name}>
              <Clock3 size={18} />
              <div>
                <strong>{version.name}</strong>
                <p>{version.note}</p>
              </div>
              <span>{version.time}</span>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

function PlanPage({ selectedChapter, setSelectedChapter }) {
  return (
    <div className="plan-layout">
      <section className="chapter-list" aria-label="章节列表">
        {planSections.map((section) => (
          <button
            className={selectedChapter === section.title ? 'chapter-item selected' : 'chapter-item'}
            key={section.title}
            onClick={() => setSelectedChapter(section.title)}
            type="button"
          >
            <span>{section.title}</span>
            <small>{section.status}</small>
          </button>
        ))}
      </section>

      <section className="editor-panel">
        <div className="editor-header">
          <div>
            <p className="eyebrow">AI 草稿生成</p>
            <h2>{selectedChapter}</h2>
          </div>
          <button className="primary-action" type="button">
            <Sparkles size={18} />
            重新生成
          </button>
        </div>

        <div className="prompt-box">
          <span>输入主题</span>
          <p>面向大学生创新创业项目的 AI 协作与商业计划书平台</p>
        </div>

        <div className="generated-copy">
          <h3>生成内容预览</h3>
          <p>
            本章节将围绕创策云 UniPlan AI 的课程项目定位展开，说明平台如何帮助大学生团队完成
            选题分析、资料整理、商业计划书撰写、PPT 大纲生成和任务协作。涉及市场规模、用户数量、
            付费转化率和财务预测的位置统一标记为“待补充数据”，待问卷调研和访谈完成后再补充。
          </p>
          <div className="warning-line">
            <ShieldCheck size={18} />
            <span>课堂提交前请补充真实问卷结果、团队成员信息和教师要求格式。</span>
          </div>
        </div>

        <div className="section-progress-grid">
          {planSections.slice(0, 4).map((section) => (
            <div className="progress-card" key={section.title}>
              <div className="board-header">
                <span>{section.title}</span>
                <strong>{section.progress}%</strong>
              </div>
              <div className="progress-track">
                <span style={{ width: `${section.progress}%` }} />
              </div>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

function TasksPage() {
  return (
    <div className="kanban">
      {taskColumns.map((column) => (
        <section className={`kanban-column ${column.tone}`} key={column.title}>
          <div className="kanban-heading">
            <h3>{column.title}</h3>
            <span>{column.tasks.length}</span>
          </div>
          {column.tasks.map((task) => (
            <article className="task-card" key={task.title}>
              <div className="task-title">
                <strong>{task.title}</strong>
                <span>{task.priority}</span>
              </div>
              <div className="task-meta">
                <span>{task.owner}</span>
                <span>{task.due}</span>
              </div>
            </article>
          ))}
        </section>
      ))}
    </div>
  );
}

function Activity({ time, text }) {
  return (
    <div className="activity-item">
      <span>{time}</span>
      <p>{text}</p>
    </div>
  );
}

function Resource({ icon: Icon, title, text }) {
  return (
    <div className="resource-item">
      <Icon size={20} />
      <div>
        <strong>{title}</strong>
        <p>{text}</p>
      </div>
    </div>
  );
}

export default App;
