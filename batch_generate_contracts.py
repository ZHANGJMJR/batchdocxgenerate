import pandas as pd
from docxtpl import DocxTemplate
import os
import logging
from datetime import datetime

# ============ 配置部分 ============
EXCEL_FILE = "data.xlsx"        # Excel 数据文件
TEMPLATE_FILE = "合同模板.docx"  # Word 模板文件
OUTPUT_DIR = "生成的合同"         # 输出文件夹
LOG_DIR = "logs"                 # 日志文件夹

# 创建输出和日志目录
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

# ============ 日志配置 ============
log_filename = os.path.join(LOG_DIR, f"{datetime.now().strftime('%Y-%m-%d')}.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_filename, encoding="utf-8"),
        logging.StreamHandler()  # 同时输出到控制台
    ]
)
logger = logging.getLogger(__name__)

# ============ 主程序 ============
def generate_contracts():
    # 1. 读取 Excel
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
    except Exception as e:
        logger.error(f"读取 Excel 失败: {e}")
        return

    logger.info(f"已读取 {len(df)} 条记录")

    # 2. 循环每行数据生成合同
    for idx, row in df.iterrows():
        try:
            # 转为字典，键名对应模板变量
            context = row.to_dict()

            # 加载模板
            tpl = DocxTemplate(TEMPLATE_FILE)

            # 渲染
            tpl.render(context)

            # 文件名：客户名称_合同编号.docx
            filename = f"{context.get('客户名称', '未命名')}_{context.get('合同编号', '')}.docx"
            filename = filename.replace("/", "-").replace("\\", "-")  # 防止非法字符

            # 保存
            output_path = os.path.join(OUTPUT_DIR, filename)
            tpl.save(output_path)

            logger.info(f"生成成功: {output_path}")

        except Exception as e:
            logger.error(f"第 {idx + 1} 条数据生成失败: {e}")

if __name__ == "__main__":
    start = datetime.now()
    generate_contracts()
    logger.info(f"全部完成，用时 {datetime.now() - start}")