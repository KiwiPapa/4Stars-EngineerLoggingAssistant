

# 打开想要更改的excel文件
from xlutils.copy import copy
import xlrd
import xlwings as xw

path1 = '.\\蓬莱9.解释成果表-1单.xls'
path2 = '.\\蓬莱9.解释成果表-1单-new.xls'
old_excel = xlrd.open_workbook(path1, formatting_info=True)
row_num = old_excel.sheets()[0].nrows
col_num = old_excel.sheets()[0].ncols
print(row_num, ' ', col_num)


# 将操作文件对象拷贝，变成可写的workbook对象
new_excel = copy(old_excel)
# 获得第一个sheet的对象
ws = new_excel.get_sheet(0)
# 写入数据
ws.write(2, 0, '解释序号')
ws.write(2, 1, '井段(m)')
ws.write(2, 2, '厚度(m)')
ws.write(2, 3, '最大声幅(%)')
ws.write(2, 4, '最小声幅(%)')
ws.write(2, 5, '平均声幅(%)')
ws.write(2, 6, '结论')

# 写入第一大列
for row in range(3, row_num):
    for col in range(0, col_num):
        item = old_excel.sheets()[0].cell_value(row, col)
        ws.write(row_num + row - 3, col, item)
# 写入第二大列
for row in range(3, row_num):
    for col in range(8, col_num):
        item = old_excel.sheets()[0].cell_value(row, col)
        ws.write(row_num + row - 3, col - 8, item)

# 另存为excel文件，并将文件命名
new_excel.save(path2)

# 删除无用列
# 创建app，打开工作表
app = xw.App(visible=False, add_book=False)
app.screen_updating = False
app.display_alerts = False
load_wb = app.books.open(path2)
load_ws = load_wb.sheets.active
# 处理列，将指定列从大到小删除（避免先删除小列导致后续列号变动）
load_ws.api.columns('O').delete
load_ws.api.columns('N').delete
load_ws.api.columns('M').delete
load_ws.api.columns('L').delete
load_ws.api.columns('K').delete
load_ws.api.columns('J').delete
load_ws.api.columns('I').delete
load_ws.api.columns('H').delete

load_ws.range('A1:G1000').columns.autofit()
# 设置边框
last_column = load_ws.range(1, 1).end('right').get_address(0, 0)[0]  # 获取最后一列
last_row = load_ws.range(1, 1).end('down').row  # 获取最后一行
a_range = f'A1:{last_column}{last_row}'  # 生成表格的数据范围
load_ws.range(a_range).api.Borders(8).LineStyle = 0  # 上边框
load_ws.range(a_range).api.Borders(9).LineStyle = 0  # 下边框
load_ws.range(a_range).api.Borders(7).LineStyle = 0  # 左边框
load_ws.range(a_range).api.Borders(10).LineStyle = 0  # 右边框
load_ws.range(a_range).api.Borders(12).LineStyle = 0  # 内横边框
load_ws.range(a_range).api.Borders(11).LineStyle = 0  # 内纵边框

# 处理完毕，保存、关闭、退出Excel
load_wb.save()
load_wb.close()
app.quit()