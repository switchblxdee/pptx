import pptx_generator
from pptx_generator.tool import GenerateDigestTool
import pandas as pd

print("откуда грузится пакет:", pptx_generator.__file__)
print("новый код есть:", hasattr(GenerateDigestTool, "_looks_like_overview"))
print("листы файла:", pd.ExcelFile("/vs_code/1.xlsx").sheet_names)
print("определяется как overview:",
      GenerateDigestTool._looks_like_overview("/vs_code/1.xlsx"))