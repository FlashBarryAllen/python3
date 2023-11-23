import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, colors

font = Font(color="FFFFFF")
border = Border(left=Side(border_style='thin', color='00000000'),
                right=Side(border_style='thin', color='00000000'),
                top=Side(border_style='thin', color='00000000'),
                bottom=Side(border_style='thin', color='00000000')
                )
tc = PatternFill(fill_type='solid', start_color='00000000')

# 创建一个工作簿
wb = openpyxl.Workbook()


# 获取工作表
ws = wb.active

pattern  = r"x(\d+) y(\d+)"
new_line = False
row_num  = 4
col_num  = 3
count    = 0
step     = 5

xy_coordinates = []

with open("log.log", "r") as f:
    # 读取 .log 文件中的每一行
    for line in f:
        # 使用正则表达式匹配每一行
        match = re.search(pattern, line)
        if match:
            x = match.group(1)
            y = match.group(2)
           
            if int(x) == 0:
                step = 5
                col_num = 3
                
                ws.insert_rows(1, row_num + 1)

                step = step * count
                count = count + 1

                col_num = col_num + 1
                xy_coordinates.append((row_num + step, col_num))
                ws.cell(row=row_num, column=col_num).value = int(x)

                col_num = col_num + 1
                xy_coordinates.append((row_num + step, col_num))
                ws.cell(row=row_num, column=col_num).value = int(y)
            else:
                col_num = col_num + 4
                xy_coordinates.append((row_num + step, col_num))
                ws.cell(row=row_num, column=col_num).value = int(x)

                col_num = col_num + 1
                xy_coordinates.append((row_num + step, col_num))
                ws.cell(row=row_num, column=col_num).value = int(y)

for row in xy_coordinates:
    x = row[0]
    y = row[1]

    print(x, y)

    cell1 = ws.cell(row=x, column=y)
    cell2 = ws.cell(row=x+1, column=y)
    ws.merge_cells(cell1.coordinate + ":" + cell2.coordinate)

    ws.cell(row=x, column=y).alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
    ws.cell(row=x, column=y).fill = tc
    ws.cell(row=x, column=y).border = border
    ws.cell(row=x, column=y).font = font

# 保存工作簿
wb.save("route.xlsx")