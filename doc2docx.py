# coding=gb2312


from win32com import client as wc
import os
import shutil
import datetime

word = wc.Dispatch('Word.Application')

src_path_56 = r"F:\pythoncode\score\src_doc_56"
dst_path_56 = r'F:\pythoncode\score\dst_docx_56'
src_path_78 = r"F:\pythoncode\score\src_doc_78"
dst_path_78 = r'F:\pythoncode\score\dst_docx_78'
dst_path_all = r'F:\pythoncode\score\dst_docx_all'
src_doc_all = r'F:\pythoncode\score\src_doc_all'

def doc2docx(src_dir, dst_dir):
    '''
    if(os.path.exists(dst_dir)):
        print("delete ", dst_dir)
        shutil.rmtree(dst_dir)
    os.mkdir(dst_dir)
    '''

    i = 0
    j = 0
    for path, subdirs, files in os.walk(src_dir):
        for wordFile in files:
            wordFullName = os.path.join(path, wordFile)
            dotIndex = wordFile.rfind(".")

            if (dotIndex != -1):
                try:
                    fileSuffix = wordFile[(dotIndex + 1):]
                    if (fileSuffix == "doc"):
                        fileName = wordFile[:dotIndex]
                        docxName = fileName + ".docx"

                        docxFullName = os.path.join(dst_dir, docxName)
                        print('正在转化：' + wordFullName)
                        doc = word.Documents.Open(wordFullName)
                        i += 1
                        doc.SaveAs(docxFullName,12)
                        doc.Close()
                    elif(fileSuffix == "docx"):
                        dstdocxFullName = os.path.join(dst_dir, wordFile)
                        shutil.copyfile(wordFullName, dstdocxFullName)
                except Exception:
                    j += 1
                    print(wordFullName + ':该文件保存失败****************************************')

    print('尝试转换' + str(i) + '个docx')
    print('其中成功的有：' + str(i - j) + '个')
    print('失败的共有:' + str(j) + '个')



# word.Visible = True #是否可见
# word.DisplayAlerts = 0
def docx2doc(src_dir, dst_dir):

    if(os.path.exists(dst_dir)):
        shutil.rmtree(dst_dir)
    os.mkdir(dst_dir)

    i = 0
    j = 0
    for path, subdirs, files in os.walk(src_dir):
        for wordFile in files:

            wordFullName = os.path.join(path, wordFile)

            dotIndex = wordFile.rfind(".")

            if (dotIndex != -1):
                try:
                    fileSuffix = wordFile[(dotIndex + 1):]
                    if (fileSuffix == "docx"):
                        fileName = wordFile[:dotIndex]
                        docxName = fileName + ".doc"

                        docxFullName = os.path.join(dst_dir, docxName)
                        print
                        '正在转化：' + wordFullName
                        doc = word.Documents.Open(wordFullName)
                        i += 1
                        doc.SaveAs(docxFullName, 1)
                        doc.Close()
                except Exception:
                    j += 1
                    print
                    wordFullName + ':该文件保存失败****************************************'

    print('尝试转换' + str(i) + '个docx')
    print('其中成功的有：' + str(i - j) + '个')
    print('失败的共有:' + str(j) + '个')

def copy_file_from_src_to_dst(src_dir,dst_dir):
    for wordFile in os.listdir(src_dir):
        # 若是文件
        dotIndex = wordFile.rfind(".")
        if (dotIndex != -1):
            try:
                fileSuffix = wordFile[(dotIndex + 1):]
                if (fileSuffix == "docx"):
                    source_file = os.path.join(src_dir, wordFile)
                    dst_file = os.path.join(dst_dir, wordFile)
                    shutil.copy(source_file, dst_file)
            except Exception:
                print("Error copy file")

if __name__ == '__main__':
    starttime = datetime.datetime.now()
    print('Start time is %s.' % (str(datetime.datetime.now())))

    doc2docx(src_doc_all, dst_path_all)
    #docx2doc(r"H:\My\stock\doc")
    #转换5,6班
    #doc2docx(src_path_56, dst_path_56)

    # 转换7,8班
    #doc2docx(src_path_78, dst_path_78)

    if 0:
        # 转换5,6，7,8班
        if (os.path.exists(dst_path_all)):
            shutil.rmtree(dst_path_all)
        os.mkdir(dst_path_all)
        copy_file_from_src_to_dst(src_path_56, dst_path_all)
        copy_file_from_src_to_dst(src_path_78, dst_path_all)

    word.Quit()

    # 程序结束时间 及 耗时
    timedelta = datetime.datetime.now() - starttime
    print('End time is %s.' % (str(datetime.datetime.now())))
    print('Total test execution duration is %s.' % (timedelta.__str__()))