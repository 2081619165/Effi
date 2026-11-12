# Effi

Effi 是用于数字证书申请资料处理的 Windows 本地客服辅助工具。它读取 PDF 或图片，使用本地 PaddleOCR 提取信息，由客服核对和修改后生成个人或机构 Word 申请表，并支持人工框选印章或签名填入三个采集格。

## 当前能力

- PDF、PNG、JPG、JPEG 客户信息提取，支持页面定位和人工选页兜底。
- 单位、法定代表人、经办人字段人工编辑，保留识别状态及格式提示。
- 个人与机构 Word 模板填充，使用界面最终值和本机生成日期。
- 印章来源提供当前 PDF、其他 PDF、DOCX 及图片。PDF/DOCX 原图进入待框选候选，独立图片可直接勾选使用。
- 统一的预览、真实像素框选、空格抓手、Ctrl＋滚轮缩放、保存替换和另存相邻新候选。
- 上移、下移交换候选，生成时按列表内勾选顺序填入最多三个采集格。
- 客户信息固定双列，印章候选列表独立滚动。
- 客户结果先交付，自动印章查找只用文字层及缓存。完整印章 OCR 由“提取印章”按需触发，保留手动页码入口。

## 环境与依赖

当前开发环境为 Windows、Python 3.10.11。依赖版本以 `requirements.txt` 为准，包括 CustomTkinter、PaddleOCR/PaddlePaddle、PyMuPDF、OpenCV、Pillow 和 python-docx。打包依赖记录在 `requirements-build.txt`，当前固定 PyInstaller 6.22.3。

## 目录

| 路径 | 职责 |
| --- | --- |
| `application_extractor.py` | 程序入口、GUI、客户字段识别与编辑、任务调度、Word 导出 |
| `ocr_runtime.py` | 本地模型加载、日志抑制、PDF 渲染和图片读取 |
| `ocr_utils.py` | OCR 文本块、坐标和归一化工具 |
| `word_template_filler.py` | 模板结构校验、字段与日期替换、共用三格图片填充 |
| `seals/panel.py` | 来源、统一候选列表、预览与框选窗口 |
| `seals/extraction.py` | 原始图片读取、DOCX 内嵌图片和候选数据结构 |
| `seals/pdf_pages.py` | 印章填报页查找、缓存、按需 OCR、高清渲染 |
| `seals/regions.py` | 采集标签工具及历史区域检测实现，清理前须核对调用 |
| `templates/` | 两份正式空白 Word 模板 |
| `tools/` | 独立技术验证工具，正式业务模块无需依赖该脚本 |
| `effi_spec_common.py` | PyInstaller 共用资源和动态依赖收集 |
| `effi_release_runtime_hook.py` | 正式 onedir 运行时设置 |
| `Effi-onedir.spec` | 正式目录包配置 |
| `docs/` | 进度、问题、决策和打包说明 |

## 本地开发

在受信任开发机安装匹配的 Windows Python 后，使用 PowerShell。

```powershell
Set-Location D:\code\Effi
py -3.10 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe application_extractor.py
```

已有 `.venv` 时直接运行最后一条命令。依赖安装需要开发环境可访问软件源，处理客户资料时使用本地模型。

正常识别和离线打包需要以下目录，每个目录必须具有 `inference.json`、`inference.pdiparams`、`inference.yml`。

```text
.paddlex/official_models/PP-OCRv5_mobile_det/
.paddlex/official_models/PP-OCRv5_mobile_rec/
```

模型目录属于本地资源，由维护者从已核验来源提供。克隆仓库后须检查资源是否完整，不能假定 Git 已包含模型，也不要在处理客户资料时临时下载模型。

## 生成与交付

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv\Scripts\python.exe -m PyInstaller Effi-onedir.spec
```

配置的默认输出为 `dist/Effi/`，入口为 `dist/Effi/Effi.exe`，依赖和业务资源位于 `_internal/`。交付整个目录。实际本地归档中已有 `dist/Effi-v3/`，版本后缀属于归档命名，spec 输出名称仍为 `Effi`。详细步骤见 [打包说明](docs/PACKAGING.md)。

## 使用要点

1. 导入申请资料，点击“开始提取”，人工核对、补充全部业务字段。
2. 选择个人或机构模板。客户信息区始终完整显示，模板决定最终导出映射。
3. 在印章区选取来源。PDF/DOCX 候选先框选保存，独立图片可直接使用或再次框选。
4. 勾选最多三张，按需上移、下移，再生成 Word。无需额外确认按钮。
5. 对照原资料检查字段、日期、印章完整性、来源、顺序和最终分页。

身份证号码以人工填写为主。复杂 Word 绘图、链接图、页眉页脚图片和 Word 图片艺术效果等存在提取限制。原图、裁剪和导出文件均应在受控本地目录保存。

## 维护与安全

修改前阅读 [AGENTS.md](AGENTS.md)。现状见 [项目进度](docs/PROGRESS.md)，待处理事项见 [问题记录](docs/BUGS.md)，设计依据见 [关键决策](docs/DECISIONS.md)。

2026-11-11 仓库交接安全检查发现 Git 历史包含两个未获准公开的测试 PDF，首次推送处于阻断状态。忽略规则无法清除历史对象，须由维护者确认资料性质并决定后续处理方式。目标仓库为公开仓库，处理完该事项前禁止推送当前历史。


## 2026-11-12 安全交接更新

维护者已确认两份历史测试 PDF 为真实客户资料，并授权清理待推送历史。采用独立初始提交重建 master，仅保留当前安全源码、空白模板及维护文档。旧 Git 历史和原始 PDF 备份在本机 .private-backup/，该目录含敏感数据，禁止分享、打包或上传。模型和 IDE 文件保留本机并从新提交排除。此前 2026-11-11 的阻断记录为历史审查记录，最终推送状态以本次交付报告及 Git 远程核对为准。


## 2026-11-12 已完成安全发布

安全初始提交 74987e7 已通过普通推送发布到 origin/master，upstream 已配置，远程 SHA 核对一致。此前的阻断状态已在新公开历史中解决。master 无旧历史父节点，敏感 PDF 不在任何公开提交中。原历史与两份客户 PDF 保存在本机 .private-backup/，该目录仍属敏感备份，禁止上传或分享。模型与 IDE 配置保留本机且不在新索引中。业务代码、依赖、spec 和空白模板与清理前 fa49841 的差异检查通过。本节为最新状态，前文 2026-11-11 记录保留作为历史依据。
