# 随访数据看板（网页共享版）

把不同格式的患者随访 Excel 表，在**浏览器内**自动归一化，并可视化用药依从性
（规范 / 不规范 / 脱落 / 其他）的统计、图表与明细。支持**脱敏 / 不脱敏两种网页快照**，
可直接离线打开或外发分享。

- 纯前端、无后端、无网络请求：数据从不离开本机。
- 部署到 GitHub Pages 后，团队成员访问同一个网址即可各自上传文件、本地计算。
- 适合「团队共用 + 外部脱敏网页快照分享」场景。

## 功能

- 多文件上传 → 关键字启发式列映射 → 自动归一化为统一字段
- 全量统计 + 三张图（状态占比环形图 / 药品 TOP5 柱状图 / 随访时间趋势折线图）
- 多维筛选（用药状态卡片点选 + 药品/药店多选 + 时间范围 + 关键词），分面联动
- 明细表 / 患者聚合视图
- **导出快照（脱敏 / 不脱敏）**：生成自包含 HTML，可离线打开、外发

## 本地预览

> 本项目的 `fetch` 已全部移除；建议用本地静态服务器打开（`file://` 双击也可用）。

```bash
cd webapp
python -m http.server 8000
# 浏览器打开 http://127.0.0.1:8000
```

## 部署到 GitHub Pages

1. 推送到仓库（默认分支 `main`，仓库根即 `webapp` 内容）。
2. 仓库 **Settings → Pages → Build and deployment → Source = Deploy from a branch**，
   选择 `main` 分支、目录 `/ (root)`，保存。
3. 等待生效后访问 `https://<用户名>.github.io/followup_dashboard/`。

## 快照（导出）功能说明

- 「脱敏快照」：姓名、电话已脱敏，文件中不含明文个人信息，可安全外发。
- 「不脱敏快照」：含明文个人信息，仅可分享给可信接收方。
- 快照为单文件 HTML，离线双击即可打开；生成时不再依赖网络，
  `https` 与 `file://` 双击打开的本页都能生成。

## 维护（重要）

逻辑代码在三个源文件中，修改后**必须重新生成** `index.html`：

- `mapping.js` —— 字段映射与状态口径配置（移植自 `mapping_config.py`）
- `pipeline.js` —— Excel 解析 / 归一化流程（移植自 `pipeline.py`）
- `app.js` —— 前端编排层（渲染、筛选、快照）

```bash
cd webapp
python build.py      # 读取 index.template.html + 三个 js，生成部署用单文件 index.html
git add -A && git commit -m "..." && git push
```

> `index.html` 是**由 `build.py` 生成的部署产物**，请勿手动直接改它；
> 改 HTML/CSS 请改 `index.template.html`，改逻辑请改上面三个 js，然后重跑 `build.py`。

## 关于本地 Flask 版（不在本仓库）

完整的本地链路版（Flask 后端 `app.py` + `pipeline.py` + `mapping_config.py` +
`dashboard.html`）位于父目录 `followup_dashboard/`，**不纳入本仓库**。
GitHub Pages 是纯静态托管，无法运行 Flask；本仓库的网页版已用浏览器内计算
完整覆盖其功能，并新增了快照导出。

## 目录结构

```
webapp/
├── index.html            # 部署用单文件（build.py 生成，勿手改）
├── index.template.html   # HTML/CSS 模板（手改这里）
├── mapping.js            # 字段映射 / 状态口径（逻辑源）
├── pipeline.js           # 解析 / 归一化（逻辑源）
├── app.js                # 前端编排层（逻辑源）
├── build.py              # 内联生成 index.html
├── vendor/
│   └── xlsx.full.min.js  # SheetJS（浏览器内解析 Excel）
├── README.md
└── .gitignore
```
