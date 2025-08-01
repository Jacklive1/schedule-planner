import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook
def parse_time_string(t_str):
    #将9:20—10:00转换成（start，end）
    t_str = str(t_str)
    if '-' in t_str:
        start,end=t_str.split('-')
        start2=datetime.strptime(start,'%H:%M').time()
        end2=datetime.strptime(end,'%H:%M').time()
        return start2,end2
    else:
        return None
#d当前时间
now = datetime.now().time()
# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))
# 构造 Excel 文件完整路径（假设文件就在脚本同目录下）
file_path = os.path.join(current_dir, "schedule.xlsx")
# 读取 Excel 文件内容
df = pd.read_excel(file_path)
event="无"
for idx, row in df.iterrows():

    time_str=row['时间']
    result=parse_time_string(time_str)
    if result :
        start,end=result
        if start<=now<=end:
            event=row['时间']+' '+row['事件']
            break
print('当前时间和内容为：\n'+event)

# 询问是否更新完成次数
choice = input("是否更新完成次数？(y/n)：").strip().lower()

if choice == 'y':
    print("\n以下是所有项目：")
    for i, row in df.iterrows():
        count = row.get('完成次数', 0)
        count = int(count) if pd.notna(count) else 0
        print(f"{i}. {row['时间']} {row['事件']}（完成次数：{count}）")

    try:
        selected = int(input("\n请输入你要增加完成次数的序号："))
        if selected==12:
            print("恭喜你度过了充实的一天，你真的很棒！")
        if 0 <= selected < len(df):
            # 用 openpyxl 精确写入对应单元格
            wb = load_workbook(file_path)
            ws = wb.active

            # 找到 "完成次数" 列的列号（从第1行中找）
            header = [cell.value for cell in ws[1]]
            col_index = header.index("完成次数") + 1  # openpyxl 是从1开始数列的
            row_index = selected + 2  # openpyxl 表格从第2行是数据（第1行为表头）

            # 获取当前值，+1后写入
            current_value = ws.cell(row=row_index, column=col_index).value
            current_value = int(current_value) if current_value else 0
            ws.cell(row=row_index, column=col_index).value = current_value + 1

            # 保存
            wb.save(file_path)
            print("✅ 成功更新完成次数，格式未被破坏。")
        else:
            print("❌ 无效序号。")
    except ValueError:
        print("❌ 请输入数字序号。")
else:
    print("操作已取消。")