# 交易数据分析平台

一个基于 React + Vite + Recharts 的本地交易数据分析工具，支持上传 Excel 文件后自动分析回购业务、债券交易、基金、可转债、国债期货等多维度数据。

## 🌐 在线访问

部署后可通过 GitHub Pages 访问：
`https://你的用户名.github.io/仓库名/`

## 📊 功能特性

- **回购业务分析** — 按金额/笔数 Top10 排名，排除超百亿异常交易
- **协议回购画像** — 主要账户的回购类型和交易方向分布
- **长期债券分析** — 10-30年/30年+债券的信用债/利率债深度分析
- **基金/转债/期货分析** — 多品种交易特征画像
- **低活跃度账户** — 交易笔数 ≤2 的账户持仓明细
- **ABS / REITs** — 特殊品种账户梳理
- **可视化** — 月度趋势、账户热力图、雷达图、债券期限结构、异常检测

## 🚀 本地运行

```bash
# 1. 安装依赖
npm install

# 2. 启动开发服务器
npm run dev

# 3. 浏览器打开 http://localhost:5173/
```

## 📦 构建部署

```bash
# 构建生产版本
npm run build

# 构建后的文件在 dist/ 目录
```

## 🌍 部署到 GitHub Pages

本仓库已配置 GitHub Actions 自动部署。只需按以下步骤操作：

1. **创建 GitHub 仓库**
   - 登录 [GitHub](https://github.com)
   - 点击右上角 **+** → **New repository**
   - 输入仓库名（如 `trading-analyzer`）
   - 选择 **Public**（私有仓库 GitHub Pages 需要付费）
   - 点击 **Create repository**

2. **推送代码到 GitHub**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/你的用户名/仓库名.git
   git push -u origin main
   ```

3. **启用 GitHub Pages**
   - 进入仓库页面 → **Settings** → **Pages**
   - **Source** 选择 **GitHub Actions**
   - 等待 Actions 运行完成（约 2-3 分钟）

4. **访问网站**
   - 部署完成后，访问 `https://你的用户名.github.io/仓库名/`

## 📁 项目结构

```
├── .github/workflows/deploy.yml  # GitHub Actions 自动部署
├── src/
│   ├── lib/analyzer.ts           # 核心分析引擎
│   ├── pages/Home.tsx            # 主页面 UI
│   ├── components/ui/            # shadcn/ui 组件
│   └── main.tsx                  # React 入口
├── vite.config.ts                # Vite 构建配置
└── package.json                  # 依赖配置
```

## 🔧 技术栈

- **React 19** + **TypeScript**
- **Vite 7** — 构建工具
- **Tailwind CSS** + **shadcn/ui** — UI 组件
- **Recharts** — 图表可视化
- **SheetJS (xlsx)** — Excel 文件解析
- **React Router** — 前端路由（HashRouter 适配 GitHub Pages）
