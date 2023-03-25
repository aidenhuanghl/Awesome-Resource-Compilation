#首先导入所需的库   
import sqlite3 

import pandas as pd

#使用pandas读取Excel文件

excel_file = 'C:/Users/Aiden/OneDrive/工作/文档/未分类/ERP 编码库.xlsx'  # 将此替换为您的Excel文件名
sheet_name = 'Sheet1'  # 将此替换为您的工作表名称
df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')

#将数据导入到SQLite数据库

db_file = 'ERP database.db'  # 将此替换为您的数据库文件名
conn = sqlite3.connect(db_file)
df.to_sql(sheet_name, conn, if_exists='replace', index=False)
conn.close()
