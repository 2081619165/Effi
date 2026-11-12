"""仅供 onedir 正式发行版使用，不影响开发运行及调试版。"""

import os


# 正式包不继承开发机或客服电脑上遗留的身份证实验开关。
os.environ["EFFI_DEBUG_ID"] = "0"
os.environ["EFFI_DEBUG_ID_MODE"] = "off"
os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"
