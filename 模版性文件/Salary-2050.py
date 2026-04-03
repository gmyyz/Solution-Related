import pandas as pd
from datetime import datetime
import logging

# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 定义文件路径和表名
file_path = r'D:\PY\薪酬\原始数据-2050.xlsx'
salary_sheet_name = '薪酬'
batch_guide_sheet_name = '批导'
output_file_path = r'D:\PY\薪酬\整理后数据-2050.xlsx'

try:
    # 读取薪酬表数据
    salary_df = pd.read_excel(file_path, sheet_name=salary_sheet_name)
    # 读取批导表表头
    batch_guide_df = pd.read_excel(file_path, sheet_name=batch_guide_sheet_name, nrows=0)
    batch_guide_columns = batch_guide_df.columns.tolist()
except FileNotFoundError:
    logging.error(f"未找到文件: {file_path}，请检查文件路径是否正确。")
except Exception as e:
    logging.error(f"读取文件时出现错误: {e}")
else:
    # 定义固定值
    fixed_values = {
        '分组': 'A1',
        '公司代码': '2050',
        '凭证类型': 'ZZ',
        '凭证日期': datetime.now().strftime('%Y%m%d'),
        '过账日期': datetime.now().strftime('%Y%m%d'),
        '币别': 'CNY',
        '记账码': '40',
        '特别总帐标识': '',
        '人员': '',
        '反记账': '',
        '原因代码': '',
        '分配': ''
    }

    # 定义需要处理的薪酬列及其对应的科目
    salary_columns = {
        '求和项:工资': '6601010001',
        '求和项:养老公司': '6601030008',
        '求和项:工伤公司': '6601030003',
        '求和项:失业公司': '6601030002',
        '求和项:医疗公司': '6601030005',
        '求和项:公积金公司': '6601030006'
    }

    # 定义薪酬列名对应的文本描述 实际2026年2月薪酬 计提2026年3月
    text_mapping = {
        '求和项:工资': '实际2026年2月薪酬',
        '求和项:养老公司': '实际2026年2月养老保险',
        '求和项:工伤公司': '实际2026年2月工伤保险',
        '求和项:失业公司': '实际2026年2月失业保险',
        '求和项:医疗公司': '实际2026年2月医疗保险',
        '求和项:公积金公司': '实际2026年2月住房公积金'
    }

    # 存储整理后的数据
    data_rows = []

    # 遍历薪酬数据的每一行
    for index, row in salary_df.iterrows():
        cost_center = row.get('成本中心', row.get('成本中心号'))  # 兼容两种可能的列名
        order = row['内部订单号']
        for salary_column, account in salary_columns.items():
            amount = row[salary_column]
            # 验证金额是否为有效数值
            if pd.notna(amount):
                text = text_mapping.get(salary_column)
                new_row = fixed_values.copy()
                new_row.update({
                    '科目': account,
                    '成本中心': cost_center,
                    '订单': order,
                    '金额（文本类型）': amount,
                    '文本': text
                })
                data_rows.append(new_row)
            else:
                logging.warning(f"行 {index} 的 {salary_column} 金额为无效值，已跳过。")

    # 将数据列表转换为DataFrame
    output_df = pd.DataFrame(data_rows, columns=batch_guide_columns)

    # 按照科目排序
    output_df = output_df.sort_values(by='科目')

    # 金额保留两位小数并转换为文本格式
    output_df['金额（文本类型）'] = output_df['金额（文本类型）'].astype(float).round(2)
    #.astype(str)

    try:
        # 将整理后的数据保存到新的Excel文件中
        output_df.to_excel(output_file_path, index=False)
        logging.info(f'数据已整理并保存到 {output_file_path}')
    except PermissionError:
        logging.error(f"没有权限保存文件到 {output_file_path}，请检查文件权限。")
    except Exception as e:
        logging.error(f"保存文件时出现错误: {e}")