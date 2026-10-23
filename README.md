# LAP 学术成果自动化发布系统 (Web of Science 官方直联版)

本系统专为 **LAP 实验室 (Lab of Artificial Photosynthesis - 朱兴旺副教授课题组)** 量身定制。
系统直接对接 **Clarivate Web of Science (WoS) 官方 API**（ResearcherID: `J-8774-2017`），实现学术成果发文与被引频次的自动抓取、双向同步删除、高可用双通道加速（Gitee+本地）与智能排版。

---

## 📁 核心文件清单与分工

| 文件名 | 用途说明 |
| :--- | :--- |
| **`朱兴旺文章_完整整合版.xlsx`** | **论文配图登记总表**：只从中读取【DOI】与【图片链接】，并自动同步提取至 `doi_image_map.json` |
| **`doi_image_map.json`** | **DOI 图片映射库**：由 Excel 自动生成，系统同步时从中匹配每篇论文的配图 |
| **`selected_doi.txt`** | **代表作控制清单**：代表作**严格且仅由此文件读取控制**，绝不自动增减任何文章 |
| **`esi_top_1_percent.txt`** | **ESI TOP 1% 高被引文章清单**：按行存放高被引文章 DOI |
| **`corresponding_titles.txt`** | **通讯作者论文标题库**：高容错算法识别通讯作者文章 |
| **`sync_wos.py`** | **WoS 官方同步核心引擎**：自动连接 WoS 官方、清洗数据、匹配图片与指标 |
| **`publications_data.js`** | **编译后的成果数据库**：更新后上传覆盖至服务器/Gitee 仓库 |
| **`fanke_module_code.html`** | **凡科建站【插件代码】模块代码**：支持 Gitee 3.5秒降级、按引用排序、数据源标注 |
| **`preview_publications.html`** | **本地离线全功能预览**：双击直接在浏览器查看最新效果 |
| **`一键同步WoS数据.bat`** | **本地一键同步**：双击即可全自动同步并生成前台 JS |
| **`.github/workflows/wos_sync.yml`** | **GitHub Actions 每周一定时云同步工作流**：支持自动同步拉取并镜像推送到 Gitee |

---

## 🌟 运行机制说明

1. **配图更新（Excel $\to$ JSON $\to$ 成果页）**：
   - 当有新论文要配图时，只需打开 `朱兴旺文章_完整整合版.xlsx`，填入【DOI】和凡科的【图片链接】；
   - 运行同步脚本时，系统会自动将 Excel 中的所有配图提取更新到 `doi_image_map.json` 中，并在生成前台卡片时自动关联。

2. **代表作管理（严格由 `selected_doi.txt` 控制）**：
   - 代表作**完全且严格**由 `selected_doi.txt` 决定；
   - 想要新增或减少代表作，只需在 `selected_doi.txt` 中增加或删除一行 DOI；
   - 筛选 Selected Papers 时，严格按预设期刊梯队排序：
     $$\text{NC} \to \text{Angew} \to \text{ACS catal} \to \text{AEM} \to \text{Adv sci} \to \text{ACB} \to \text{CEJ} \to \text{Chin. J. Catal.} \to \text{small} \to \text{其它期刊}$$

3. **双线高可用加载（Gitee 主线 + 3.5秒本地自动降级）**：
   - 前端页面优先向 **Gitee CDN** 读取 `publications_data.js`（国内秒级极速响应）；
   - 若 3.5 秒内 Gitee 网络抖动或未返回，系统**无感、秒级自动切换回本地域名**（`https://zxw.ac.cn/publications_data.js`），确保网站 100% 永不白屏。
