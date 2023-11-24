import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

font      = Font(color="FFFFFF")
fill      = PatternFill(fill_type='solid', start_color='00000000')
alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
border    = Border(left=Side(border_style='thin', color='00000000'), right=Side(border_style='thin', color='00000000'),
                   top=Side(border_style='thin', color='00000000'),bottom=Side(border_style='thin', color='00000000'))

pattern  = r"x(\d+) y(\d+)"
new_line = False
row_num  = 4
col_num  = 3
step     = 5

xy_coordinates = []

wb = openpyxl.Workbook()
ws = wb.active

def get_xp():
    count = 0
    with open("log.log", "r") as f:
        for line in f:
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

def pait_xp():
    for row in xy_coordinates:
        x = row[0]
        y = row[1]

        #print(x, y)

        cell1 = ws.cell(row=x, column=y)
        cell2 = ws.cell(row=x+1, column=y)
        ws.merge_cells(cell1.coordinate + ":" + cell2.coordinate)

        ws.cell(row=x, column=y).alignment = alignment
        ws.cell(row=x, column=y).fill      = fill
        ws.cell(row=x, column=y).border    = border
        ws.cell(row=x, column=y).font      = font

def main():
    get_xp()
    pait_xp()
    wb.save("route.xlsx")

main()