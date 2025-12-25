"""
Excel文件对比工具
用于对比两个xlsx文件的指定worksheet，支持多列组合键
"""

import argparse
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd


def setup_logging(output_file: str) -> logging.Logger:
    """配置日志输出"""
    logger = logging.getLogger("xlsx_compare")
    logger.setLevel(logging.INFO)
    
    # 文件处理器
    file_handler = logging.FileHandler(output_file, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    
    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # 格式
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger


def load_excel(file_path: str, sheet_name: str) -> pd.DataFrame:
    """加载Excel文件的指定worksheet"""
    return pd.read_excel(file_path, sheet_name=sheet_name, engine="calamine")


def create_composite_key(df: pd.DataFrame, key_columns: list) -> pd.Series:
    """根据多个列创建组合键"""
    return df[key_columns].astype(str).agg("||".join, axis=1)


def compare_xlsx(
    file1: str,
    file2: str,
    sheet1: str,
    sheet2: str,
    key_columns: list,
    logger: logging.Logger
) -> None:
    """对比两个Excel文件"""
    
    logger.info("=" * 60)
    logger.info("Excel文件对比开始")
    logger.info("=" * 60)
    logger.info(f"文件1: {file1} (Sheet: {sheet1})")
    logger.info(f"文件2: {file2} (Sheet: {sheet2})")
    logger.info(f"组合键列: {key_columns}")
    logger.info("-" * 60)
    
    # 加载数据
    try:
        df1 = load_excel(file1, sheet1)
        df2 = load_excel(file2, sheet2)
    except Exception as e:
        logger.error(f"加载文件失败: {e}")
        return
    
    logger.info(f"文件1行数: {len(df1)}, 列数: {len(df1.columns)}")
    logger.info(f"文件2行数: {len(df2)}, 列数: {len(df2.columns)}")
    
    # 验证键列存在
    for col in key_columns:
        if col not in df1.columns:
            logger.error(f"文件1中不存在列: {col}")
            return
        if col not in df2.columns:
            logger.error(f"文件2中不存在列: {col}")
            return
    
    # 创建组合键
    df1["_composite_key"] = create_composite_key(df1, key_columns)
    df2["_composite_key"] = create_composite_key(df2, key_columns)
    
    keys1 = set(df1["_composite_key"])
    keys2 = set(df2["_composite_key"])
    
    # 找出差异
    only_in_file1 = keys1 - keys2
    only_in_file2 = keys2 - keys1
    common_keys = keys1 & keys2
    
    logger.info("-" * 60)
    logger.info("行级别差异统计:")
    logger.info(f"  仅在文件1中存在的行: {len(only_in_file1)}")
    logger.info(f"  仅在文件2中存在的行: {len(only_in_file2)}")
    logger.info(f"  两文件共有的行: {len(common_keys)}")
    
    # 输出仅在文件1中的行
    if only_in_file1:
        logger.info("-" * 60)
        logger.info("仅在文件1中存在的行:")
        for key in sorted(only_in_file1):
            logger.info(f"  键值: {key}")
    
    # 输出仅在文件2中的行
    if only_in_file2:
        logger.info("-" * 60)
        logger.info("仅在文件2中存在的行:")
        for key in sorted(only_in_file2):
            logger.info(f"  键值: {key}")
    
    # 对比共有行的数据差异
    logger.info("-" * 60)
    logger.info("共有行的数据差异:")
    
    # 获取共同列（排除组合键列）
    common_columns = [c for c in df1.columns if c in df2.columns and c != "_composite_key"]
    diff_count = 0
    
    df1_indexed = df1.set_index("_composite_key")
    df2_indexed = df2.set_index("_composite_key")
    
    for key in sorted(common_keys):
        row1 = df1_indexed.loc[key]
        row2 = df2_indexed.loc[key]
        
        row_diffs = []
        for col in common_columns:
            val1 = row1[col] if col in row1.index else None
            val2 = row2[col] if col in row2.index else None
            
            # 处理NaN比较
            val1_is_nan = pd.isna(val1)
            val2_is_nan = pd.isna(val2)
            
            if val1_is_nan and val2_is_nan:
                continue
            elif val1_is_nan != val2_is_nan or val1 != val2:
                row_diffs.append((col, val1, val2))
        
        if row_diffs:
            diff_count += 1
            logger.info(f"  键值: {key}")
            for col, v1, v2 in row_diffs:
                logger.info(f"    列[{col}]: 文件1='{v1}' vs 文件2='{v2}'")
    
    if diff_count == 0:
        logger.info("  无数据差异")
    
    # 列差异
    cols1 = set(df1.columns) - {"_composite_key"}
    cols2 = set(df2.columns) - {"_composite_key"}
    only_cols1 = cols1 - cols2
    only_cols2 = cols2 - cols1
    
    if only_cols1 or only_cols2:
        logger.info("-" * 60)
        logger.info("列级别差异:")
        if only_cols1:
            logger.info(f"  仅在文件1中存在的列: {sorted(only_cols1)}")
        if only_cols2:
            logger.info(f"  仅在文件2中存在的列: {sorted(only_cols2)}")
    
    logger.info("=" * 60)
    logger.info("对比完成")
    logger.info("=" * 60)


def main():
    parser = argparse.ArgumentParser(description="对比两个Excel文件")
    parser.add_argument("file1", help="第一个Excel文件路径")
    parser.add_argument("file2", help="第二个Excel文件路径")
    parser.add_argument("--sheet1", default="Sheet1", help="文件1的worksheet名称")
    parser.add_argument("--sheet2", default="Sheet1", help="文件2的worksheet名称")
    parser.add_argument("--keys", required=True, nargs="+", help="组合键列名（可指定多个）")
    parser.add_argument("--output", default=None, help="输出日志文件路径")
    
    args = parser.parse_args()
    
    # 默认输出文件名
    if args.output is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        args.output = f"compare_result_{timestamp}.log"
    
    logger = setup_logging(args.output)
    
    compare_xlsx(
        args.file1,
        args.file2,
        args.sheet1,
        args.sheet2,
        args.keys,
        logger
    )
    
    print(f"\n对比结果已保存到: {args.output}")


if __name__ == "__main__":
    main()
