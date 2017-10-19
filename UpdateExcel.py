# 采用这些类库没有直接的修改方法，只能读取后复制、修改、写回
from xlrd import open_workbook
from xlutils.copy import copy
# 打开文件
rb = open_workbook("example.xls")
# 复制
wb = copy(rb)
# 选取表单
s = wb.get_sheet(0)
# 写入数据
s.write(0, 1, 'new data')
# 保存
wb.save('example.xls')