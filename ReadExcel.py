import xlrd
# 打开 xls 文件
book = xlrd.open_workbook("test.xlsx")
print("表单数量:", book.nsheets)
print("表单名称:", book.sheet_names())
# 获取第1个表单
sh = book.sheet_by_index(0)
print(u"表单 %s 共 %d 行 %d 列" % (sh.name, sh.nrows, sh.ncols))
print("第二行第三列:", sh.cell_value(1, 2))
# 遍历所有表单
for s in book.sheets():
    for r in range(s.nrows):
        # 输出指定行
        print(s.row(r))

# 时间格式修正
# new_date = xlrd.xldate.xldate_as_datetime(date, book.datemode)
# xlrd.xldate.xldate_from_datetime_tuple
# style = xlwt.easyxf(num_format_str='D-MMM-YY')
# ws.write(1, 0, datetime.now(), style)