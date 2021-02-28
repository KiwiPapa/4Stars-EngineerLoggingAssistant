
# 打开想要更改的excel文件
import os

from xlutils.copy import copy
import xlrd
import xlwings as xw

path1 = '.\\蓬莱9.解释成果表-1单.xls'
path2 = ''.join([path1.replace('.' + path1.split('.')[-1], ''), '(已规范化).xls'])
old_excel = xlrd.open_workbook(path1, formatting_info=True)
row_num = old_excel.sheets()[0].nrows
col_num = old_excel.sheets()[0].ncols
print(row_num, ' ', col_num)

############################################### 若为两大列则进行单列规范化
if int(col_num) > 10:
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

    ######################## 删除无用列
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

    # 获取行数
    info = load_ws.used_range
    last_row = info.last_cell.row
    alpha_dict = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K', 12: 'L',
                  13: 'M', 14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T', 21: 'U', 22: 'V', 23: 'W',
                  24: 'X', 25: 'Y', 26: 'Z'}
    last_column_num = info.last_cell.column
    last_column = alpha_dict[last_column_num]
    print(last_row, ' ', last_column)

    # 删除最后的空行，单数行情况
    last_row_value = load_ws.range(last_row, last_column_num).value
    if last_row_value == None:
        # load_ws.api.rows(last_row).delete
        load_ws.range(f'{last_column}{last_row}').api.EntireRow.Delete()
        last_row = str(int(last_row) - 1)

    # 设置边框
    a_range = f'A1:{last_column}{last_row}'  # 生成表格的数据范围
    load_ws.range(a_range).api.Borders(8).LineStyle = 1  # 上边框
    load_ws.range(a_range).api.Borders(9).LineStyle = 1  # 下边框
    load_ws.range(a_range).api.Borders(7).LineStyle = 1  # 左边框
    load_ws.range(a_range).api.Borders(10).LineStyle = 1  # 右边框
    load_ws.range(a_range).api.Borders(12).LineStyle = 1  # 内横边框
    load_ws.range(a_range).api.Borders(11).LineStyle = 1  # 内纵边框
    load_ws.range(a_range).api.Font.Name = 'Times New Roman'
    load_ws.range(a_range).api.RowHeight = 20

    load_ws.range(a_range).columns.autofit()

    b_range = load_ws.range(a_range)
    # 设置单元格 字体格式
    b_range.color = 255, 255, 255       # 设置单元格的填充颜色
    b_range.api.Font.ColorIndex = 1     # 设置字体的颜色，具体颜色索引见下方。
    b_range.api.Font.Size = 11          # 设置字体的大小。
    b_range.api.Font.Bold = False       # 设置为粗体。
    b_range.api.HorizontalAlignment = -4108    # -4108 水平居中。 -4131 靠左，-4152 靠右。
    b_range.api.VerticalAlignment = -4108      # -4108 垂直居中（默认）。 -4160 靠上，-4107 靠下， -4130 自动换行对齐。
    # b_range.api.NumberFormat = "0.00"          # 设置单元格的数字格式。

    # 处理完毕，保存、关闭、退出Excel
    load_wb.save()
    load_wb.close()
    app.quit()

    if os.path.exists(path1):
        # 删除文件，可使用以下两种方法
        os.remove(path1)
        # os.unlink(my_file)
    else:
        print('no such file:%s' % path1)

############################################### 若为一大列则进行直接进行规范化
elif int(col_num) < 10:
    # 将操作文件对象拷贝，变成可写的workbook对象
    new_excel = copy(old_excel)
    # 获得第一个sheet的对象
    ws = new_excel.get_sheet(0)
    # 写入数据
    ws.write(2, 0, '解释序号')
    ws.write(2, 1, '井段(m)')
    ws.write(2, 2, '厚度(m)')
    ws.write(2, 3, '最大指数')
    ws.write(2, 4, '最小指数')
    ws.write(2, 5, '平均指数')
    ws.write(2, 6, '结论')
    # 另存为excel文件，并将文件命名
    new_excel.save(path2)

    ######################## 格式进一步规范
    # 创建app，打开工作表
    app = xw.App(visible=False, add_book=False)
    app.screen_updating = False
    app.display_alerts = False
    load_wb = app.books.open(path2)
    load_ws = load_wb.sheets.active

    # 获取行数
    info = load_ws.used_range
    last_row = info.last_cell.row
    alpha_dict = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K', 12: 'L',
                  13: 'M', 14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T', 21: 'U', 22: 'V', 23: 'W',
                  24: 'X', 25: 'Y', 26: 'Z'}
    last_column_num = info.last_cell.column
    last_column = alpha_dict[last_column_num]
    print(last_row, ' ', last_column)

    # 删除最后的空行，单数行情况
    last_row_value = load_ws.range(last_row, last_column_num).value
    if last_row_value == None:
        # load_ws.api.rows(last_row).delete
        load_ws.range(f'{last_column}{last_row}').api.EntireRow.Delete()
        last_row = str(int(last_row) - 1)

    # 设置边框
    a_range = f'A1:{last_column}{last_row}'  # 生成表格的数据范围
    load_ws.range(a_range).api.Borders(8).LineStyle = 1  # 上边框
    load_ws.range(a_range).api.Borders(9).LineStyle = 1  # 下边框
    load_ws.range(a_range).api.Borders(7).LineStyle = 1  # 左边框
    load_ws.range(a_range).api.Borders(10).LineStyle = 1  # 右边框
    load_ws.range(a_range).api.Borders(12).LineStyle = 1  # 内横边框
    load_ws.range(a_range).api.Borders(11).LineStyle = 1  # 内纵边框
    load_ws.range(a_range).api.Font.Name = 'Times New Roman'
    load_ws.range(a_range).api.RowHeight = 20

    load_ws.range(a_range).columns.autofit()

    b_range = load_ws.range(a_range)
    # 设置单元格 字体格式
    b_range.color = 255, 255, 255  # 设置单元格的填充颜色
    b_range.api.Font.ColorIndex = 1  # 设置字体的颜色，具体颜色索引见下方。
    b_range.api.Font.Size = 11  # 设置字体的大小。
    b_range.api.Font.Bold = False  # 设置为粗体。
    b_range.api.HorizontalAlignment = -4108  # -4108 水平居中。 -4131 靠左，-4152 靠右。
    b_range.api.VerticalAlignment = -4108  # -4108 垂直居中（默认）。 -4160 靠上，-4107 靠下， -4130 自动换行对齐。
    # b_range.api.NumberFormat = "0.00"          # 设置单元格的数字格式。

    # 处理完毕，保存、关闭、退出Excel
    load_wb.save()
    load_wb.close()
    app.quit()

else:
    pass