# GradeLab · 智能批改台

面向教师的桌面批量批改工具：把学生提交整理成工作区，调用大模型逐份评分，自动生成个人成绩报告（PNG）、成绩单与全班总成绩单（Excel）。

- 界面：Python + PySide6（Qt Widgets），液态玻璃风格、背景缓慢流动、深/浅色主题
- 批改：DeepSeek 默认接口，也可接入任意 OpenAI 兼容接口；模型只接收文本，不接触文件系统
- 打包：PyInstaller（单文件 exe）+ Inno Setup（安装包）

当前版本：**0.9.2**

---

## 项目简介

一次作业往往有几十份提交。逐份打开、对照要求写评语、再汇总成绩，是大量重复劳动。GradeLab 把这条链路固定成一个可复用的流程：

1. 把学生提交按约定放进一个工作区（一人一目录，或 `学号_姓名_题名.py` 平铺）；
2. 导入名单，或让程序自动识别「学生名单」表格；
3. 用一段自然语言描述批改要求，按需附上评分规则、考卷、参考答案；
4. 程序在本地把每份提交转成文本，逐份调用大模型评分；
5. 自动生成个人报告 PNG（写回学生目录）、成绩单 Excel，以及带排名与统计的全班总成绩单。

### 适合场景

- 高校 / 培训机构的软件设计、程序设计类课程作业
- 需要批量批改客观题与主观题的普通作业、测验
- 希望把评分标准沉淀为固定流程的教研场景

### 两种批改模式

| 模式 | 适用 | 评分方式 | 附加材料 |
|---|---|---|---|
| 工程批改 | 软件项目、设计文档 | 功能 40 / 实用性 40 / UI 20 | 成绩单模板、H5 报告模板 |
| 教学批改 | 作业、试卷 | 按题号逐项给分 + 总分等级 | 考卷、参考答案 |

切换模式时，主页背景色调与底栏可选项会一起变化。

### 设计取向

- **本地优先**：文件读取、Excel 生成、报告渲染全部在本机完成；只有整理后的文本会发给模型，模型不接触文件系统
- **结果可追溯**：每次批改都会在工作区 `.smartgrader/result_*.json` 留下完整记录
- **模型可替换**：默认 DeepSeek，也可接任意 OpenAI 兼容接口，自动兼容带/不带 `/v1` 的路径
- **版式可替换**：内置工程版与教学版报告，工程模式还支持自定义 H5 模板
- **开箱可用**：提供免安装 exe 与安装包；安装包按当前用户安装，无需管理员权限，自带卸载程序

### 技术栈

Python 3.11 · PySide6（Qt Widgets / QPainter / Qt WebEngine） · openpyxl · python-docx · pypdf · PyInstaller · Inno Setup

---

## 功能一览

| 模块 | 说明 |
|---|---|
| 双模式 | 主页胶囊按钮切换「工程批改 / 教学批改」，背景色调与底栏选项随之变化 |
| 工程批改 | 按功能 40 / 实用性 40 / UI 20 三维评分，生成软件设计报告 |
| 教学批改 | 可额外提供考卷、参考答案，按题号逐项给分并生成教学版报告 |
| 批量批改 | 每个学生一份，自动识别学号、姓名、班级 |
| 报告 | 个人报告 PNG 写入该学生提交所在目录；报告版式自适应内容高度 |
| 成绩单 | 按模板填充，无模板时生成标准成绩单；另生成含排名与统计的全班总成绩单 |
| 名单 | 从 xlsx / csv 读取；表格无法识别时可调用 API 转写；也可自动发现「学生名单」文件 |
| 可选材料 | 评分规则、考卷、参考答案、成绩单模板、H5 报告模板 |
| 历史项目 | 搜索、排序、打开成绩单/总成绩单/工作区、删除记录 |
| 主题 | 深色 / 浅色 / 跟随系统；字体大小、显示大小、字体颜色、背景颜色、背景图片 |
| 背景图片 | 支持 png / jpg / jpeg / webp / bmp / gif，按窗口比例铺满并居中裁切 |

---

## 支持的文件与目录结构

### 可读取的文件

- 代码：`.py .java .kt .c .h .cpp .cs .js .ts .vue .go .rs .php .rb .swift .sql .html .css .json .xml .yml .yaml .toml .ipynb` 等
- 文档：`.txt .md .docx .pdf`
- 表格：`.xlsx .xlsm .csv`

PDF 依赖文本层；扫描件（纯图片 PDF）无法提取文字，会提示读取失败。

### 目录结构 A：一人一个文件夹

```
工作区/
  2024001_张伟/
    main.py
    utils.py
  2024002_王芳/
    main.py
```

### 目录结构 B：平铺提交（文件名带学号姓名）

```
学生/
  作业要求.txt
  学生名单.xlsx
  作业/
    2021010101_王心怡_成绩统计.py
    2021010102_李梓涵_成绩统计.py
```

程序会自动把 `学号_姓名_题名.ext` 的每个文件识别为一份提交；选择外层「学生」目录也会自动下钻到 `作业` 并拆分。

### 名单自动识别

若未手动导入名单，程序会在工作区及其上级目录寻找文件名含「名单 / 学生」的 `.xlsx/.xlsm/.csv`，读取成功即自动使用（界面「名单」显示人数）。
没有名单时，会从 `学号_姓名_题名` 文件名中解析学号与姓名（班级留空）。

---

## 快速开始（源码运行）

环境：Windows 10/11 x64，Python 3.11+。

```bat
pip install -r requirements.txt
python main.py
```

也可直接双击 `run.bat`。

### 使用流程

1. 主页选择模式（工程 / 教学）。
2. 点「工作区」选择学生提交所在目录（教学模式显示为「作答目录」）。
3. 点「名单」导入学生名单（可跳过，程序会尝试自动识别）。
4. 按需选择底栏可选项：成绩单 / 报告单开关、成绩单模板、评分规则、考卷、参考答案、H5 报告模板。
5. 在输入框描述批改要求（Ctrl + Enter 发送），开始批改。
6. 完成后可打开成绩单、总成绩单或工作区；左上角「新建任务」清空对话重新开始。

---

## API 设置

- 默认：DeepSeek，接口地址 `https://api.deepseek.com`，默认模型 `deepseek-flash`
- 自定义：任意 OpenAI 兼容接口，填写接口地址、模型名称、API Key
- 设置页会显示「实际请求地址」；请求自动兼容带/不带 `/v1` 的路径，并对 429/500/503 自动重试
- 「查看 API 契约」可复制批改 JSON 契约与表格转写提示词，便于对接其他服务

模型只接收程序整理好的文本并返回 JSON，不会直接读取或写入任何文件；本地读写均由程序完成。

---

## 打包与安装包

```bat
build.bat
```

流程：安装依赖 → 清理 `build/`、`dist/` → PyInstaller 打包 → Inno Setup 生成安装包。

产物：

- `dist/GradeLab.exe`：单文件免安装版
- `dist/GradeLab-Setup-<版本>.exe`：安装包（当前用户安装到 `%LOCALAPPDATA%\Programs\GradeLab`，含开始菜单快捷方式，可选桌面快捷方式，自带卸载程序）

生成安装包需要 Inno Setup 6；未安装时可执行：

```bat
winget install JRSoftware.InnoSetup
```

### 版本号位置

升级版本时同步修改：

- `app/config.py` 的 `APP_VERSION`
- `installer.iss` 的 `#define MyAppVersion`
- `build.bat` 末尾提示的文件名

---

## 自定义 H5 报告模板

在底栏选择 `.html` 模板后，工程批改会用 Qt WebEngine 渲染整页并截图。模板中使用 `{{占位符}}`，未提供的占位符会被清空。

可用占位符：

```
student_id  student_name  student_class  total_score  score_color_class  level  comment
project_topic
software_name  software_subtitle  software_type  tech_stack  software_theme
main_features  feature_highlights  feature_improvements
target_scope  target_audience  utility_highlights  utility_improvements
ui_appearance  ui_evaluation
feature_score  utility_score  ui_score
```

`score_color_class` 会按分数返回 `score-red / score-yellow / score-green / score-bright-green / score-gold`。
教学模式使用内置教学版式报告，不使用 H5 模板。

---

## 配置文件与数据

| 路径 | 内容 |
|---|---|
| `%APPDATA%\SmartGrader\config.json` | API、主题、上次工作区等设置 |
| `%APPDATA%\SmartGrader\projects.json` | 历史项目记录 |
| `<工作区>\.smartgrader\result_*.json` | 每次批改的完整结果 |
| `<工作区>\成绩单_*.xlsx` | 成绩单 |
| `<工作区>\全班总成绩单_*.xlsx` | 全班总成绩单（含排名、统计、分数段） |
| `<学生目录>\成绩报告单.png` | 个人报告（多人共用目录时文件名为 `学号_姓名_题名_成绩报告单.png`） |

---

## 常见问题

**提示接口 404**
接口地址填到 `/v1` 为止，不要带 `/chat/completions`；设置页的「实际请求地址」可用于核对。

**只识别到很少的提交**
确认工作区层级：要么每个学生一个子目录，要么使用 `学号_姓名_题名.ext` 的平铺命名。程序在识别过少时会自动下钻一层。

**姓名显示成整串文件名**
说明没有可用名单。导入「学生名单.xlsx」后重新批改即可；或确保文件名符合 `学号_姓名_题名`。

**报告是空白/生成失败**
H5 模板不可用时会自动回退到内置版式；若仍失败，检查模板占位符与图片路径。

**换电脑后背景图片丢失**
背景图以文件路径保存，原图移动或删除后会回落到默认渐变，重新选择即可。

---

## 项目结构

```
GradeLab/
  app/
    config.py           配置读写与默认值
    theme.py            调色板与全局 QSS
    widgets.py          液态玻璃控件、动态背景、模式胶囊、新建任务按钮
    main_window.py      主窗口与侧栏导航
    pages/
      home_page.py      工作台：模式切换、工作区/名单、批改流程
      history_page.py   历史项目
      settings_page.py  设置：API、主题、背景图片
    grading_service.py  扫描分组、读取文本、调用模型、生成成绩单与报告
    excel_service.py    名单读取、成绩单填充、总成绩单与统计
    report_renderer.py  内置报告 PNG 绘制（工程版 / 教学版）
    report_html.py      H5 模板渲染（Qt WebEngine）
    llm_client.py       OpenAI 兼容接口客户端
    workers.py          后台线程
    models.py / contract.py / store.py / table_ai.py
  main.py               程序入口（含 --render-report 自检）
  GradeLab.spec         PyInstaller 配置
  installer.iss         Inno Setup 配置
  build.bat / run.bat
  requirements.txt
  icon.png / icon.ico
  报告模板_软件开发报告.html
```

### 自检命令

打包后可验证 H5 渲染链路：

```bat
GradeLab.exe --render-report 报告模板_软件开发报告.html out.png
```

---

## 版本

- 0.9.2：背景缓慢流动、模式切换改变背景色调、主题支持自定义背景图片（含 GIF）；修复平铺提交的分组、名单自动识别与总成绩单姓名错乱问题
- 0.9.1：支持「学号_姓名_题名」平铺提交
- 0.9.0：双模式（工程/教学）、模式胶囊、新建任务、教学版报告
