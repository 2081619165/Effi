"""Effi 正式版与调试版共用的 PyInstaller 收集配置。"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_all, copy_metadata


def collect_build_inputs(
    project_dir: Path,
) -> tuple[list, list, list[str], list[str]]:
    """只显式加入业务资源，并保守保留 OCR 动态依赖。"""
    datas = [
        (
            str(project_dir / "templates" / "机构证章申请表.docx"),
            "templates",
        ),
        (
            str(project_dir / "templates" / "个人证书及电子印章申请表.docx"),
            "templates",
        ),
        (
            str(
                project_dir
                / ".paddlex"
                / "official_models"
                / "PP-OCRv5_mobile_det"
            ),
            ".paddlex/official_models/PP-OCRv5_mobile_det",
        ),
        (
            str(
                project_dir
                / ".paddlex"
                / "official_models"
                / "PP-OCRv5_mobile_rec"
            ),
            ".paddlex/official_models/PP-OCRv5_mobile_rec",
        ),
    ]
    binaries = []
    hiddenimports: list[str] = []

    # PaddleOCR/PaddleX 动态加载较多，第二版仍完整收集核心包。
    # chardet 7.x 含 mypyc 原生扩展，需显式收集其隐式子模块。
    for package in ("paddle", "paddleocr", "paddlex", "cv2", "chardet"):
        package_datas, package_binaries, package_hiddenimports = collect_all(package)
        datas += package_datas
        binaries += package_binaries
        hiddenimports += package_hiddenimports

    # 保留运行时版本查询所需的发行包元数据。
    for distribution in (
        "paddleocr",
        "paddlepaddle",
        "paddlex",
        "PyMuPDF",
        "python-docx",
        "opencv-contrib-python",
        "chardet",
        "imagesize",
        "pyclipper",
        "pypdfium2",
        "python-bidi",
        "shapely",
    ):
        datas += copy_metadata(distribution)

    # 这些入口不参与本地 OCR；其余 Paddle 动态模块暂不裁剪。
    excludes = [
        "paddlex.inference.serving",
        "paddleocr._cli",
        "paddleocr._doc2md",
    ]
    return datas, binaries, hiddenimports, excludes
