# coding=gb2312


from win32com import client as wc
import os
import shutil

word = wc.Dispatch('Word.Application')

src_path = r"F:\py\src_doc"
dst_path = r'F:\py\dst_docx'

def doc2docx(src_dir, dst_dir):
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

    word.Quit()
    print('尝试转换' + str(i) + '个docx')
    print('其中成功的有：' + str(i - j) + '个')
    print('失败的共有:' + str(j) + '个')



# word.Visible = True #是否可见
# word.DisplayAlerts = 0
def docx2doc(src_dir, dst_dir):
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

    word.Quit()
    print('尝试转换' + str(i) + '个docx')
    print('其中成功的有：' + str(i - j) + '个')
    print('失败的共有:' + str(j) + '个')


if __name__ == '__main__':
    #docx2doc(r"H:\My\stock\doc")
    doc2docx(src_path, dst_path)
