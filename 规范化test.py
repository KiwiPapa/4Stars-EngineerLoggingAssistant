

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

#
for row in range(3, row_num):
    for col in range(8, col_num):
        item = old_excel.sheets()[0].cell_value(row, col)
        ws.write(row_num + row - 3, col - 8, item)

# 另存为excel文件，并将文件命名
new_excel.save(path2)