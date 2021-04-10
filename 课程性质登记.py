"""
2019.09.05，处理课程性质修改登记表单。
1. 读取问卷信息表（xlrd），将信息转换到输出表（openpyxl）中。
2. 读取名册（xlrd），检查已填写数量，输出没填的人的名单。
"""
import openpyxl,xlrd

data_file = 'data/45278222_0_现工院课程性质错误登记表_135_133.xls'
output_file = 'data/output.xlsx'
register_file = 'data/现代工学院2017级花名册.xlsx'


def read_register_dict(filename)->dict:
    """
    读取名册文件。返回格式：dict<str,tuple>
    学号：(姓名，专业)。
    学号用str。
    """
    wb = xlrd.open_workbook(filename)
    ws = wb.sheets()[0]
    dct = {}
    for row in range(1,ws.nrows):
        values = list(ws.row_values(row))[:3]
        dct[str(int(values[0]))]=(values[1],values[2])
    # wb.close()
    return dct

def trans_data(register_file,data_file,output_file):
    """
    读取并转换数据，同时比对。
    以学号为key比对。用dict登记所有姓名出现的次数。输出以下提示：
    1. 没填的人。
    2. 一人多份的。
    3. 学号姓名不匹配的。
    4. 不存在的学号，直接拒绝录入。
    """
    register_dct = read_register_dict(register_file)
    data_wb = xlrd.open_workbook(data_file)
    data_sheet = data_wb.sheets()[0]
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.append(('学号','姓名','课程号','课程名','改为何种性质课程'))
    count_dct = {}  # 数据结构：学号-次数。
    for row in range(1,data_sheet.nrows):
        values = list(data_sheet.row_values(row))
        number = values[6]
        name = values[7]
        # 拒绝录入不存在的学号
        if register_dct.get(number,None) is None:
            print("不存在的学号，拒绝录入！",number,name)
            continue
        elif register_dct[number][0] != name:
            print(f"警告：姓名与学号不匹配，以学号为准。{number}，表单姓名是{name}，名册姓名是"
                  f"{register_dct[number][0]}")
        count_dct[number]=count_dct.get(number,0)+1
        for i in range(5):
            id_col = i*3+8
            course_col = i*3+9
            type_col = i*3+10
            cid = values[id_col]
            course = values[course_col]
            ctype = values[type_col]
            if cid == '(空)' or cid == '(跳过)':
                continue
            out_ws.append((number,name,cid,course,ctype))
    print("填写表单总数：",len(count_dct))
    print("未填写表单总数：",len(register_dct)-len(count_dct))
    # 输出一人填写多份表单的
    print("一人多份表格：")
    for name,num in count_dct.items():
        if num > 1:
            print(name,register_dct[name][0],num)
    # 未填写表单的
    no_numbers = list(sorted(set(register_dct.keys()) - set(count_dct.keys())))
    print("未填写的：")
    major_count = {}
    for number in no_numbers:
        print(number,register_dct[number])
        major=register_dct[number][1]
        major_count[major]=major_count.get(major,0)+1
    print("未填写专业统计：")
    for major,count in major_count.items():
        print(major,count)
    # data_wb.close()
    out_wb.save(output_file)
    out_wb.close()


if __name__ == '__main__':
    trans_data(register_file,data_file,output_file)
