
import os

import re
import jieba

import collections
import codecs
from util.operate_mysql import *
from util.operate_mysql import clear_table as clear_table
import docx
import doc
import xlrd
import xlwt
from xlutils.copy import copy

docx_path = 'F:\\pythoncode\\score\\dst_docx\\' #要分析的文件目录
student_idname = 'F:\\pythoncode\\score\\jk152078.xls'
score_fiel = 'F:\\pythoncode\\score\\jk152078score.xls'


def get_max_min_size():
    min = 0
    max = 0
    file_list = os.listdir(docx_path)
    for file_name in file_list:
        file_name_list = file_name.split(".")
        if file_name_list[1] == "docx":
            document = docx.Document(docx_path+file_name)
        #elif file_name_list[1] == "doc":
        #    document = doc.Document(docx_path+file_name)
        contents_list = []
        # 取文件内容
        for paragraph in document.paragraphs:
            contents_list.append(paragraph.text.encode('utf-8'))
        # 分词
        words_list = []
        file_len = 0
        for text in contents_list:
            file_len = file_len + len(text)

        if file_len > max:
            max = file_len
        if 0 == min:
            min = file_len
        if file_len<min:
            min = file_len
    return max, min


def get_score_from_file(file_name, max, min):
    document = docx.Document(file_name)
    contents_list = []
    # 取文件内容
    for paragraph in document.paragraphs:
        contents_list.append(paragraph.text.encode('utf-8'))
    # 按文件长度计分
    words_list = []
    file_len = 0
    for text in contents_list:
        file_len = file_len + len(text)

    score = round(file_len/max * 15) + 80
    return score



def parase_all_and_record_score():
    #打开学生名单excel，copy一份另存并写入成绩
    source_excel_file = xlrd.open_workbook(student_idname)
    dest_excel_file = copy(source_excel_file)
    read_table = source_excel_file.sheet_by_name(u'Sheet1')  # 通过名称获取
    write_table = dest_excel_file.get_sheet(0)

    max, min = get_max_min_size()
    file_list = os.listdir(docx_path)

    for file_name in file_list:
        file_name_list = []
        file_name_list = file_name.split('-')
        student_tmp = file_name_list[2].split(".")
        student_name = student_tmp[0]
        score = 0
        for i in range(read_table.nrows):
            get_name_fromcell = read_table.cell(i,2).value
            if(get_name_fromcell == student_name):
                print(student_tmp)
                if student_tmp[1] == "docx":
                    score = get_score_from_file(docx_path+file_name, max, min)
                #elif student_tmp == "doc":
                #    document = doc.Document(file_name)
                #    score = get_score_from_file(document, max, min)
                write_table.write(i, 3, score)

    dest_excel_file.save(score_fiel)

def parase_all_and_insert_db():
    clear_table('words_temp')  # 清除数据
    operatMySQl = OperateMySQL()

    jieba.load_userdict("dict.txt")

    # 获取文本内容
    contents_list = []
    file_list = os.listdir(docx_path)
    for filename in file_list:
        filepathname= docx_path + filename
        print(filepathname)
        contents_list.clear()
        document = docx.Document(filepathname)
        #取文件内容
        for paragraph in document.paragraphs:
            contents_list.append(paragraph.text.encode('utf-8'))
        # 分词
        words_list = []
        for text in contents_list:
            splits = jieba.cut(text)
            for word in splits:
                words_list.append(word.lower())

    # 统计分词频率，计入数据库
    counter = collections.Counter(words_list)

    # 插入数据库
    for collect in counter:
        sqli = "insert into words_temp values ('{0}',{1});"
        sqlm = sqli.format(collect, counter[collect])
        operatMySQl.execute(sqlm)
    operatMySQl.commit()


    # 清洗无用数据
    clean_date_table = ['的', ' ', '\t', '\n', '、', '.', '，', '/', '；', '（', '）']
    for clean_item in clean_date_table:
        sqli = "delete from words_temp  where words_content = '{0}';"
        sqlm = sqli.format(clean_item)
        operatMySQl.execute(sqlm)
    operatMySQl.commit()


if __name__ == '__main__':
    starttime = datetime.datetime.now()
    print('Start time is %s.' % (str(datetime.datetime.now())))

    parase_all_and_insert_db()

    #parase_all_and_record_score()

    # 程序结束时间 及 耗时
    timedelta = datetime.datetime.now() - starttime
    print('End time is %s.' % (str(datetime.datetime.now())))
    print('Total test execution duration is %s.' % (timedelta.__str__()))



