import pandas as pd
from docxtpl import DocxTemplate
import os
import logging
from datetime import datetime
import re
import shutil

# ============ 配置部分 ============
EXCEL_FILE = "data.xlsx"        # Excel 数据文件
TEMPLATE_FILE = "合同模板.docx"  # Word 模板文件
OUTPUT_DIR = "生成的合同"         # 输出文件夹
LOG_DIR = "logs"                 # 日志文件夹
FAIL_FILE = "生成失败.xlsx"       # 失败记录文件
MAX_NAME_LEN = 20                # 第一列文件名最大长度

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

# ============ 工具函数 ============
def sanitize_filename(name: str) -> str:
    """去除文件名中的非法字符，并截断"""
    name = re.sub(r'[\/\\\:\*\?\"\<\>\|]', '-', name)  # 去除非法字符
    if len(name) > MAX_NAME_LEN:
        name = name[:MAX_NAME_LEN]  # 截断
    return name

def clear_output_dir():
    """清空目标目录"""
    for f in os.listdir(OUTPUT_DIR):
        file_path = os.path.join(OUTPUT_DIR, f)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            logger.error(f"删除文件失败 {file_path}: {e}")

# ============ 主程序 ============
def generate_contracts():
    # 清空目标目录
    clear_output_dir()
    logger.info(f"已清空目录: {OUTPUT_DIR}")

    # 1. 读取 Excel
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
    except Exception as e:
        logger.error(f"读取 Excel 失败: {e}")
        return

    logger.info(f"已读取 {len(df)} 条记录")

    total_count = len(df)
    success_count = 0
    fail_count = 0
    fail_records = []

    # 2. 循环每行数据生成合同
    for idx, row in df.iterrows():
        try:
            # 转为字典，键名对应模板变量
            context = row.to_dict()

            # 加载模板
            tpl = DocxTemplate(TEMPLATE_FILE)

            # 渲染
            tpl.render(context)

            # 获取第一列的值
            first_col_value = str(row.iloc[0]) if not pd.isna(row.iloc[0]) else "未命名"
            first_col_value = sanitize_filename(first_col_value)

            # 文件名：行索引_第一列的值.docx
            filename = f"{idx + 1}_{first_col_value}.docx"

            # 保存
            output_path = os.path.join(OUTPUT_DIR, filename)
            tpl.save(output_path)

            logger.info(f"生成成功: {output_path}")
            success_count += 1

        except Exception as e:
            logger.error(f"第 {idx + 1} 条数据生成失败: {e}")
            fail_count += 1

            # 记录失败数据，并附上失败原因和原Excel行号
            fail_row = row.copy()
            fail_row["原Excel行号"] = idx + 2  # Excel 第一行是表头，所以 +2
            fail_row["失败原因"] = str(e)
            fail_records.append(fail_row)

    # 3. 导出失败记录
    if fail_records:
        fail_df = pd.DataFrame(fail_records)
        fail_df.to_excel(FAIL_FILE, index=False)
        logger.info(f"失败记录已保存到: {FAIL_FILE}")

    # 4. 统计结果
    logger.info("==== 生成统计 ====")
    logger.info(f"总记录数: {total_count}")
    logger.info(f"成功数量: {success_count}")
    logger.info(f"失败数量: {fail_count}")

if __name__ == "__main__":
    start = datetime.now()
    generate_contracts()
    logger.info(f"全部完成，用时 {datetime.now() - start}")