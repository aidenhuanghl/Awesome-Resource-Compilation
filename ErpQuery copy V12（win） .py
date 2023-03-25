#无法复制单行
#修改规格栏可以为空
#支持在规格栏内输入10 26 8 来查询轴承,但是规格栏不能为空

#支持回车键查询
#可按输入内容进行查询，并支持多个条件查询
import sqlite3
import tkinter as tk
from tkinter import ttk
import pyperclip
import os

# 定义 query_window 为全局变量
query_window = None

is_topmost = False


def query_database(db_file, query):
    conn = sqlite3.connect(db_file)
    cur = conn.cursor()
    cur.execute(query)
    results = cur.fetchall()
    conn.close()
    return results

def display_results(results):
    root = tk.Tk()
    root.title("查询结果")

    # 创建滚动条
    scrollbar = tk.Scrollbar(root, orient="vertical")
    scrollbar.grid(row=0, column=1, sticky='ns')

    # 创建 Treeview
    tree = ttk.Treeview(root, columns=("行号", "品号", "品名", "规格", "无", "图号"), show="headings", yscrollcommand=scrollbar.set, selectmode="extended")
    tree.grid(row=0, column=0, sticky='nsew')

    # 配置滚动条
    scrollbar.config(command=tree.yview)

    # 设置标题
    tree.heading("行号", text="行号")
    tree.heading("品号", text="品号")
    tree.heading("品名", text="品名")
    tree.heading("规格", text="规格")
    tree.heading("图号", text="图号")
    tree.heading("无", text="无")


    # 设置列宽
    tree.column("行号", width=50, stretch=True)
    tree.column("品号", width=100, stretch=True)
    tree.column("品名", width=200, stretch=True)
    tree.column("规格", width=500, stretch=True)
    tree.column("图号", width=200, stretch=True)
    tree.column("无", width=0, stretch=True)


    for index, result in enumerate(results, start=1):
    # 将图号值添加到 values 元组中
        tree.insert("", tk.END, values=(index, *result[:-1], result[-1]))


    # 绑定单元格单击事件
    tree.bind("<Double-1>", lambda event: copy_cells(tree, event, include_headers=True))

    # 设置权重以便调整窗口大小
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()

    # for index, result in enumerate(results, start=1):
    # # 将图号值添加到 values 元组中
    #     print(f"当前行的结果：{result}") # 打印当前行的结果
    #     tree.insert("", tk.END, values=(index, *result[:-1], result[-1]))


def toggle_topmost():
    global query_window, is_topmost
    if is_topmost:
        query_window.attributes("-topmost", False)
        is_topmost = False
        topmost_button_text.set("切换置顶 (取消置顶)")
    else:
        query_window.attributes("-topmost", True)
        is_topmost = True
        topmost_button_text.set("切换置顶 (置顶)")


def create_query_window():
    global query_window
    query_window = tk.Tk()
    query_window.title("查询条件")
    
    # 创建品号输入框标签
    label_product_code = tk.Label(query_window, text="请输入品号：")
    label_product_code.grid(row=0, column=0, padx=(10, 5), pady=(10, 5), sticky="e")

    # 创建品号输入框
    entry_product_code = tk.Entry(query_window)
    entry_product_code.grid(row=0, column=1, padx=(5, 10), pady=(10, 5), sticky="w")

    # 创建品名输入框标签
    label_product_name = tk.Label(query_window, text="请输入品名：")
    label_product_name.grid(row=1, column=0, padx=(10, 5), pady=(10, 5), sticky="e")

    # 创建品名输入框
    entry_product_name = tk.Entry(query_window)
    entry_product_name.grid(row=1, column=1, padx=(5, 10), pady=(10, 5), sticky="w")

    # 创建规格输入框标签
    label_specification = tk.Label(query_window, text="请输入规格：")
    label_specification.grid(row=2, column=0, padx=(10, 5), pady=(10, 5), sticky="e")

    # 创建规格输入框
    entry_specification = tk.Entry(query_window)
    entry_specification.grid(row=2, column=1, padx=(5, 10), pady=(10, 5), sticky="w")

    # 创建查询按钮
    query_button = tk.Button(query_window, text="查询", command=lambda: execute_query(entry_product_code.get(), entry_product_name.get(), entry_specification.get(), query_window))
    query_button.grid(row=3, columnspan=2, pady=(5, 10))

    # 创建 "切换置顶" 按钮，并将按钮文本绑定到 StringVar
    global topmost_button_text
    topmost_button_text = tk.StringVar()
    topmost_button_text.set("切换置顶 (取消置顶)")
    topmost_button = tk.Button(query_window, textvariable=topmost_button_text, command=toggle_topmost)
    topmost_button.grid(row=4, columnspan=2, pady=(5, 10))

    # 绑定回车键
    query_window.bind("<Return>", lambda event: execute_query(entry_product_code.get(), entry_product_name.get(), entry_specification.get(), query_window))

 

    # 创建并显示查询窗口
create_query_window()


def execute_query(condition_product_code, condition_product_name, condition_specification, query_window):
    db_file = r'C:\Users\Aiden\OneDrive\个人\PythonAutomatic\ERP database.db'

    # 处理规格查询条件
    specification_conditions = condition_specification.split()
    if specification_conditions:
        specification_query = " AND ".join(f"规格 LIKE '%{spec}%' " for spec in specification_conditions)
        query = f"SELECT * FROM Sheet1 WHERE 品号 LIKE '%{condition_product_code}%' AND 品名 LIKE '%{condition_product_name}%' AND {specification_query};"
    else:
        query = f"SELECT * FROM Sheet1 WHERE 品号 LIKE '%{condition_product_code}%' AND 品名 LIKE '%{condition_product_name}%';"

    results = query_database(db_file, query)
    display_results(results)
    query_window.destroy()

def copy_cells(tree, event, include_headers=False):
    # 获取当前焦点单元格的行和列
    row = tree.identify_row(event.y)
    column = tree.identify_column(event.x)

    # 获取所有选中的行
    selected_rows = tree.selection()

    # 如果仅选择了一个单元格，则仅复制该单元格
    if len(selected_rows) == 1 and column != "#0":
        cell_value = tree.item(row)["values"][int(column[1]) - 1]
        pyperclip.copy(str(cell_value))
    else:
        # 初始化一个列表来存储选中行的值
        rows_values = []

        # 如果需要复制表头
        if include_headers:
            headers = [tree.heading(col)["text"] for col in tree["columns"]]
            rows_values.append('\t'.join(headers))

        for selected_row in selected_rows:
            row_values = tree.item(selected_row)["values"]
            rows_values.append('\t'.join(str(val) for val in row_values))

        # 将行值列表连接为一个字符串，用换行符分隔
        rows_str = '\n'.join(rows_values)

        # 复制到剪贴板
        pyperclip.copy(rows_str)


        import tkinter as tk

query_window.mainloop()


