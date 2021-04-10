"""
2019.09.25。
读取data/courses中的选课数据，生成选课全表，保存在data/选课全表.xlsx中。
"""

source_dir = 'data/courses'
output_file = 'data/选课全表.xlsx'

import xlrd,openpyxl,os

def read(source_dir)->(dict,dict):
    """
    读取表格内容，返回结构1：
    学号:{
        课程名1:选课类型,
        课程名2:选课类型,
        ...,
        姓名: //这两项每次更新一次
        专业:
    }
    学号，姓名，年级，学院，专业，课程1，课程2，。。。
    返回结构2：各个课程对应的列标号，且从第5列开始。
    """
    os.chdir(source_dir)
    data = {}
    courses = set()
    for file in os.scandir('.'):
        if not file.is_dir():
            # 这个判断其实没用
            wb = xlrd.open_workbook(file.name)
            course_name = file.name.split('.')[0]
            courses.add(course_name)
            ws = wb.sheets()[0]
            for r in range(1,ws.nrows):
                number = ws.cell(r,0).value
                name = ws.cell(r,1).value
                selecttype = ws.cell(r,2).value
                selecttype = '未知' if not selecttype else selecttype
                major = ws.cell(r,3).value
                data.setdefault(number,{})["name"]=name
                data[number]["major"]=major
                data[number][course_name]=selecttype
    course_col = {}
    for i,c in enumerate(courses):
        course_col[c]=i+5
    os.chdir('..');os.chdir('..')
    return data,course_col

def write(data:dict,course_col:dict,output_file:str):
    wb = openpyxl.Workbook()
    ws = wb.active
    titles = ['学号','姓名','年级','学院','专业',]
    titles.extend(course_col.keys())
    ws.append(titles)
    n = len(course_col.keys())
    for number,dct in sorted(data.items()):
        try:
            grade,institute,major = dct['major'].split('-')
        except ValueError:
            # 2015-材料-材料类,材料-生物医学工程
            print(dct['major'])
            grade=dct['major'].split('-')[0]
            institute,major = dct['major'].split(',')[-1].split('-')

        line = [number,dct['name'],grade,institute,major]+['']*n
        del dct['major']
        del dct['name']
        for course,tp in dct.items():
            line[course_col[course]]=tp
        ws.append(line)
    wb.save(output_file)

if __name__ == '__main__':
    data,course_col = read(source_dir)
    write(data,course_col,output_file)
