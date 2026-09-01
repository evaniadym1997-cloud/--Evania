---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 232f2c8f9e33633f1529ca406dee815a_7d9adf4aa5c911f1a0d9525400826444
    ReservedCode1: PLA3luUfa8KzWiGKbpAqBAR+mdMSPeMvy4lJfCJCQP1JWKQ3cm0N8wcMY4gSauAZ/NiM/lImz5ogHWwLEN7PZrdLhG4nJttBf+1TOWCbk5qhfSb6wVZEY0rSxCCmdsctszdrQ6VCwR+oZp/DYjYjvlCKnrq4swiBb0ChA64vkHRyvu6eY7NOWuFzV/c=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 232f2c8f9e33633f1529ca406dee815a_7d9adf4aa5c911f1a0d9525400826444
    ReservedCode2: PLA3luUfa8KzWiGKbpAqBAR+mdMSPeMvy4lJfCJCQP1JWKQ3cm0N8wcMY4gSauAZ/NiM/lImz5ogHWwLEN7PZrdLhG4nJttBf+1TOWCbk5qhfSb6wVZEY0rSxCCmdsctszdrQ6VCwR+oZp/DYjYjvlCKnrq4swiBb0ChA64vkHRyvu6eY7NOWuFzV/c=
---

# 讲评网页生成器（通用版）

基于漫画 + 讲评结构的可交互 HTML 讲评网页。页面内置「＋ 导入」面板：上传或拖入 Word（.docx）/ PDF / TXT，或直接粘贴文本后，可自动按现有形式生成新的讲评页面。

## 功能特性

- 30 页交互式讲评：18 页漫画 + 12 页讲评
- Tab 切换、逐框点击展开、总览表格逐格点击
- 四段式思维导图、人物关系图、两栏/三栏并列对比
- 可打字文本框、涂鸦工具栏（含放大镜）、页码 xx/30
- info-fold 折叠块，讲评部分以点击动作展开具体内容
- 「＋ 导入」面板：拖入 / 上传 / 粘贴 → 自动提取文本 → 复制指令发送给 Marvis 生成新讲评

## 使用方式

1. 直接打开 `index.html` 即可浏览（无需服务器）。
2. 点击右上角「＋ 导入」：
   - 拖入 `.docx` / `.pdf` / `.txt` / `.md` 文件，或点击选择文件；
   - 或直接在文本框粘贴文章 / 课件原文；
   - 提取出的文本会显示在「提取结果预览」中，可手动编辑补充；
   - 点击「复制文本 · 发送给 Marvis 自动生成讲评」，将复制好的指令粘贴给 Marvis，即可自动生成新的讲评页面。

> 说明：`.docx` 与 `.pdf` 解析依赖 CDN 在线加载（mammoth / pdf.js），离线环境请直接粘贴文本。

## 目录结构

```
anjal-tabla-comic/
├── index.html              # 主页面（含导入面板）
├── assets/                 # 样式与运行时脚本
│   └── runtime.js          # hash 深链、键盘导航等
├── images/                 # 漫画分镜图片
├── docs/
│   └── lesson-template.html # 讲评板块通用模板
└── README.md
```

---

# GitHub 部署详细步骤

## 方式一：GitHub Pages（推荐，免费、无需服务器）

### 1. 创建 GitHub 仓库

1. 登录 [github.com](https://github.com)，点击右上角 `+` → `New repository`；
2. 填写仓库名（如 `lesson-generator`），选择 **Public**（Pages 免费要求公开仓库）；
3. 不要勾选 "Add a README file"（避免冲突），点击 `Create repository`。

### 2. 推送代码到仓库

在项目目录执行（需先安装 Git，并配置好 GitHub 账号）：

```bash
# 进入项目目录
cd anjal-tabla-comic

# 初始化仓库（若尚未初始化）
git init
git add .
git commit -m "feat: 通用版讲评网页，支持导入文档自动生成"

# 关联远程仓库（把下面的地址换成你的仓库地址）
git remote add origin https://github.com/<你的用户名>/lesson-generator.git

# 推送到 GitHub
git branch -M main
git push -u origin main
```

> 若未配置 GitHub 凭据：Windows 可使用 [GitHub Desktop](https://desktop.github.com/) 登录后推送；macOS/Linux 可执行 `gh auth login`（GitHub CLI）或使用 Personal Access Token 作为密码。

### 3. 开启 GitHub Pages

1. 打开仓库页面 → `Settings` → 左侧 `Pages`；
2. `Source` 选择 `Deploy from a branch`，`Branch` 选择 `main`，目录选 `/ (root)`；
3. 点击 `Save`，等待 1~2 分钟，页面顶部会出现访问地址：
   `https://<你的用户名>.github.io/lesson-generator/`

### 4. 后续更新

本地修改后重新推送即可自动重新部署：

```bash
git add .
git commit -m "update"
git push
```

## 方式二：Vercel（可选）

1. 将代码推送到 GitHub 仓库（同上）；
2. 打开 [vercel.com](https://vercel.com)，用 GitHub 账号登录；
3. `Add New Project` → `Import` 选择刚推送的仓库；
4. Framework Preset 选 `Other`，Build Command 留空，Output Directory 留空；
5. 点击 `Deploy`，等待完成后即可获得 `https://<项目名>.vercel.app` 链接。

## 方式三：Cloudflare Pages（可选）

1. 推送代码到 GitHub 仓库；
2. 打开 [dash.cloudflare.com](https://dash.cloudflare.com) → `Workers & Pages` → `Create` → `Pages` → `Connect to Git`；
3. 选择仓库，Build settings 全部留空（纯静态站点），点击 `Save and Deploy`；
4. 部署完成后可获得 `https://<项目名>.pages.dev` 链接。

## 常见问题

| 问题 | 解决 |
|------|------|
| push 时提示认证失败 | 配置 GitHub token：`git remote set-url origin https://<用户名>:<token>@github.com/<用户名>/lesson-generator.git` |
| Pages 页面 404 | 确认仓库为 Public，且 Pages 设置的 Branch/目录正确；推送后等待 1~2 分钟 |
| 图片/资源不显示 | 确认 `images/`、`assets/` 与 `index.html` 位于同一目录层级，路径为相对路径 |
*（内容由AI生成，仅供参考）*
