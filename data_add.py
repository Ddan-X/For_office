import os
import xlrd

fl = []

# user 输入文件目录
dic = input('信息是否添加，显示所有信息脚本，输入文件目录（如 F:\python_DanL\ZoeOffice\登记表）：')

#获取文件
def get_file(filepath):
    for i in os.listdir(filepath):
        path = os.path.join(filepath, i)
        if os.path.isdir(path):
            get_file(path)
        if path.endswith(".xls"):
            #print(path) #(输出文件夹所有文件)
            fl.append(path) #文件存入fl(list)
get_file(dic) #存放文件的文件夹路径，根据不同进行修改

for f in fl:
    workbook = xlrd.open_workbook(f, formatting_info=True)
    sheet = workbook.sheet_by_index(1)

    sheet.cell_value(0,0)

    #print(sheet.row_values(2))

    rows, cols = sheet.nrows, sheet.ncols
    print(sheet.row_values(1))
    print("=====================================================================================================================================")
    for row in range(2,rows):

        print(sheet.row_values(row))