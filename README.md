# LAP 学术成果自动化发布系统 (Web of Science 官方直联版)

本系统专为 **LAP 实验室 (Lab of Artificial Photosynthesis - 朱兴旺副教授课题组)** 量身定制。
系统直接对接 **Clarivate Web of Science (WoS) 官方 API**（ResearcherID: `J-8774-2017`），实现学术成果发文与被引频次的自动抓取、双向同步、全网 6 级高可用 CDN 加速、智能多维筛选与响应式排版。

---

## 📁 核心文件清单与分工

| 文件名 | 用途说明 |
| :--- | :--- |
| **`朱兴旺文章_完整整合版.xlsx`** | **论文配图登记总表**：只从中读取【DOI】与【图片链接】，并自动同步提取至 `doi_image_map.json` |
| **`doi_image_map.json`** | **DOI 图片映射库**：由 Excel 自动生成，系统同步时从中匹配每篇论文的配图 |
| **`selected_doi.txt`** | **代表作控制清单**：代表作**严格且仅由此文件读取控制**，绝不自动增减任何文章 |
| **`esi_top_1_percent.txt`** | **ESI TOP 1% 高被引文章清单**：按行存放高被引文章 DOI |
| **`corresponding_titles.txt`** | **通讯作者论文标题库**：高容错算法精准识别通讯作者与第一作者文章 |
| **`crossref_cache.json`** | **卷期页码本地缓存数据库**：自动查询 Crossref API 补全缺失卷期页码并本地持久化 |
| **`sync_wos.py`** | **WoS 官方同步核心引擎**：自动连接 WoS 官方 API、清洗数据、提取 AuthorKeywords、智能分类、匹配配图与指标 |
| **`publications_data.js`** | **编译后的成果数据库**：前端加载的核心数据源，推送到 GitHub 后由全球 CDN 分发 |
| **`fanke_module_code.html`** | **凡科建站【插件代码】模块代码**：嵌入凡科网的完整前端组件（内置 6 级多 CDN 自动容灾、双维排序、多维检索、无横向滚动条） |
| **`preview_publications.html`** | **本地离线全功能预览**：双击直接在浏览器查看最新效果 |
| **`一键同步WoS数据.bat`** | **本地一键同步批处理**：双击即可全自动抓取 WoS 最新成果并生成前台数据 |
| **`.github/workflows/wos_sync.yml`** | **GitHub Actions 定时云同步工作流**：支持定时全自动同步拉取并镜像推送到 Gitee |
| **`.gitignore`** | **Git 忽略配置文件**：确保本地密钥文件（`wos_api_key.txt`）、缓存与临时文件绝不上演上传 |

---

## 🌟 核心运行机制

### 1. 配图更新机制（Excel $\to$ JSON $\to$ 成果页）
- 当有新论文需要添加代表图（Graphical Abstract / TOC）时，只需打开 `朱兴旺文章_完整整合版.xlsx`，在对应行填入【DOI】和凡科图片空间的【图片链接】；
- 运行同步脚本时，系统会自动将 Excel 中的所有配图提取更新到 `doi_image_map.json` 中，并在生成前台卡片时自动关联渲染。

### 2. 代表作管理机制（严格由 `selected_doi.txt` 控制）
- 代表作**完全且严格**由 `selected_doi.txt` 决定，绝不产生非预期的增减；
- 想要新增或减少代表作，只需在 `selected_doi.txt` 中增加或删除一行 DOI；
- 前台展示 Selected Papers 时，严格按预设顶级期刊梯队排序：
  $$\text{NC} \to \text{Angew} \to \text{ACS Catal.} \to \text{AEM} \to \text{Adv. Sci.} \to \text{ACB} \to \text{CEJ} \to \text{Chin. J. Catal.} \to \text{Small} \to \text{其它期刊}$$

### 3. 研究方向智能聚类（Author Keywords + Title 深度匹配）
系统直接整合 Clarivate WoS 官方文章中的作者关键词（Author Keywords）与标题关键词，精准聚类为以下 9 大标准学术方向：
1. **光催化 (Photocatalysis)**
2. **CO2 Reduction**
3. **2D Materials**
4. **H2 Evolution**
5. **N2 Conversion**
6. **Metal Oxides**
7. **Environmental Remediation**
8. **Defect Engineering**
9. **其它主题 (Other Research)**

### 4. 全网 6 级多 CDN 高可用加载架构
为彻底解决第三方平台防盗链（403）和海外/国内访问延迟问题，前端采用 6 级多源自动降级加载策略：
1. **jsDelivr (Fastly Asia/Global)**：亚太与国内秒级响应边缘；
2. **jsDelivr (Gcore China Optimized)**：国内电信、联通、移动网络深度优化节点；
3. **jsDelivr (Main Anycast CDN)**：全球 Anycast 智能调度路由；
4. **jsDelivr (Cloudflare Global)**：欧美及全球海外节点全覆盖；
5. **本地服务器保底 (`zxw.ac.cn`)**：自建域名本地资源保底；
6. **GitHub Raw Direct**：GitHub 官方直链终极防线。

**智能容灾与极速渲染**：
- 任意节点率先返回即刻（约 0.2 秒）渲染完成；
- 若单个节点网络抖动，2.2 秒内自动无感切换至下一节点，保障全球访客 100% 秒开、永不白屏。

---

## 🔒 API Key 安全与权限配置指南

为了防止 Clarivate Web of Science API 密钥泄露，系统采用**本地隔离 + 云端加密**的双重防护机制：

### 1. 本地运行配置（安全且无需反复输入）
- 在本地运行脚本时，系统会自动读取本目录下的 `wos_api_key.txt`；
- 该文件已被写入 `.gitignore`，**永远不会被 Git 提交或推送到 GitHub/Gitee**，确保密钥仅留在您的本地电脑；
- 若需要配置或更换 Key，只需在本地的 `wos_api_key.txt` 中填入您的 Clarivate API Key（支持单 Key 或多个 Key 用英文逗号隔开，脚本已支持多 Key 自动故障切换）。

### 2. GitHub Actions 云端加密配置
若启用了 GitHub Actions 自动云同步，请勿将密钥写在任何代码文件中，只需在 GitHub 仓库中配置加密 Secret：
1. 进入您的 GitHub 仓库页面（如 `https://github.com/zxwiii/lap-publications`）；
2. 依次点击顶部 **Settings** $\to$ 左侧菜单 **Secrets and variables** $\to$ **Actions**；
3. 点击 **New repository secret**：
   - **Name**：填写 `WOS_API_KEY`
   - **Secret**：填入您的 Clarivate WoS 官方 API 密钥（例如 Primary Key）
4. 点击 **Add secret** 保存。

> [!NOTE]
> GitHub Secrets 采用非对称加密存储，任何人（包括仓库所有者和公开访客）在保存后均无法查看明文，仅在 GitHub Actions 运行阶段在隔离内存中注入使用。

---

## 🚀 日常运维与更新流程

### 模式 A：本地一键更新并发布
1. **更新配图或代表作**：
   - 若有新配图，更新 `朱兴旺文章_完整整合版.xlsx`；
   - 若有新代表作，更新 `selected_doi.txt`；
2. **运行同步**：
   - 双击运行 **`一键同步WoS数据.bat`**；
   - 脚本将自动连接 WoS 官方抓取最新发文与被引数据，并生成最新的 `publications_data.js`；
3. **本地预览**：
   - 双击打开 **`preview_publications.html`**，在浏览器中查看最新成果卡片与排序；
4. **推送到 GitHub / 凡科网**：
   - 将更新后的文件提交并推送至 GitHub 仓库，jsDelivr 全球 CDN 将在数分钟内自动完成同步分发；
   - 若凡科模块有样式或逻辑升级，复制 `fanke_module_code.html` 全文覆盖凡科建站后台代码模块即可。

### 模式 B：GitHub Actions 全自动定时云同步
- 仓库内预设的工作流（`.github/workflows/wos_sync.yml`）默认每周一自动执行；
- 自动抓取 WoS 官方最新数据并提交至 `main` 分支，同时镜像推送到 Gitee 仓库，全程无需人工干预。
