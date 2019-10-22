import os
import xlrd

fl = []

# user 输入文件目录
dic = input('信息是否修改脚本，输入文件目录（如 F:\python_DanL\ZoeOffice\登记表）：')

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


def get_front_color(xf):
    font = workbook.font_list[xf.font_index]
    if not font:
        return None
    return get_color(font.colour_index)
def get_color(color_index):
    return workbook.colour_map.get(color_index)


for f in fl:
    workbook = xlrd.open_workbook(f, formatting_info=True)
    sheet = workbook.sheet_by_index(0)

    sheet.cell_value(0,0)

    print(sheet.row_values(2))

    rows, cols = sheet.nrows, sheet.ncols
    for col in range(cols):
        c = sheet.cell(2, col)
        xf = workbook.xf_list[c.xf_index]

        if get_front_color(xf) != (0,0,0):
            print(get_front_color(xf),c.value,"有变动;  ")











