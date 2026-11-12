# Windows onedir 打包与交付

## 当前依据

本地 `.venv/pyvenv.cfg` 记录 Python 3.10.11，基础环境位于 Windows Python 3.10。生产和构建版本分别以 `requirements.txt`、`requirements-build.txt` 为准，PyInstaller 固定为 6.22.3。新机器使用匹配架构的 Windows Python，目标机兼容性须实测。

正式配置为 `Effi-onedir.spec`，入口为 `application_extractor.py`。`effi_spec_common.py` 收集两份 Word 模板、两套 PaddleOCR 模型、动态依赖与发行元数据。正式配置另收集 CustomTkinter assets，并运行 `effi_release_runtime_hook.py`。

## 打包前

1. 检查 Git 状态、拟发布源码提交、模型和模板。使用受控本地资源，排除客户资料及导出文件。
2. 确认两份模板为批准的空白模板。
3. 确认 `.paddlex/official_models/` 下 `PP-OCRv5_mobile_det`、`PP-OCRv5_mobile_rec` 各有 `inference.json`、`inference.pdiparams`、`inference.yml`。spec 会检查文件存在性。
4. 使用已有验证过的 Python 环境。文档命令不会在正常客户处理流程中自动执行。

```powershell
Set-Location D:\code\Effi
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv\Scripts\python.exe -m PyInstaller Effi-onedir.spec
```

本轮交接未执行这些命令。不要将文档中的命令视为以后任务的自动打包授权。

## 输出与资源路径

- 默认工作产物在 `build/Effi-onedir/`。
- spec 的 `EXE` 和 `COLLECT` 名称均为 `Effi`，默认交付目录为 `dist/Effi/`。
- 主程序为 `dist/Effi/Effi.exe`，依赖、模型、模板等在 `_internal/`。
- 源码运行从模块目录寻找资源，冻结运行使用 PyInstaller 的 `_MEIPASS`。
- 本地已有 `dist/Effi-v1`、`Effi-v2`、`Effi-v3` 和 ZIP 归档。`Effi-v3` 中确认存在 EXE 和 `_internal`，名称后缀是归档方式，无法仅凭目录时间证明与某个提交完全一致。

## 交付和验收

交付整个目录或保持其结构的 ZIP。建议归档时记录源码 SHA、依赖版本、构建日期、Windows 环境和验收结果，客户文件与构建目录分开。

在无开发环境的目标 Windows 机离线检查启动、模型加载、资料识别、人工编辑、全部字段可见、来源切换、PDF 手动页码、框选与排序、两种 Word 的 0～3 张印章、日期、长字段和打印排版。记录失败条件和脱敏日志。

## 其他配置

`Effi.spec` 是单文件配置，`Effi-debug.spec` 是带控制台的单文件调试配置。二者保留既有用途，正式 onedir 构建使用本页指定配置，不随意混用。

## 常见问题

| 现象 | 排查方向 |
| --- | --- |
| 模型缺失或无法离线识别 | 检查两个模型目录及三个必要文件，确认归档包含 `_internal` 完整资源 |
| 客服机仅拿到 EXE 后无法启动 | 核对整个 onedir 目录是否完整解压，依赖是否被移动或拦截 |
| 界面字体或图标异常 | 核对 CustomTkinter assets、目标机 DPI 和字体环境 |
| 动态模块缺失 | 根据调试信息核对共用收集规则，避免未经验证裁剪 Paddle 依赖 |
| 启动或识别慢 | 区分冷模型初始化、客户识别、按需印章整页 OCR、锁等待和最终墙钟 |
| Word 找不到模板或结构校验失败 | 检查正式模板文件名、资源位置和实际结构，不用客户填写后的文件覆盖模板 |

保留现有性能日志。正式无控制台模式将库输出重定向到安全输出流，排错可由维护者在源码或受控调试版本中复现；日志不得包含客户原文。


## 2026-11-12 稳定版本归档核对

当前继续使用 `Effi-onedir.spec`，默认输出 `dist/Effi/`，交付完整目录及 `_internal` 资源；不改变已有依赖、模型加载、运行时钩子或构建命令。LibreOffice 实验已取消，不属于 Effi 正式运行或打包必需组件，也不随包分发。本轮只做静态配置和文档核对，没有重新构建 EXE；稳定业务基准为 `17c6192`，现有发行包与该基准的对应关系仍以实际构建记录和人工验收为准。未来固定底图 PNG 方案尚未实现，若引入新的资源或交付要求，届时另行更新本页。
