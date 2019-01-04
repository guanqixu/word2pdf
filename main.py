# -*- encoding: utf-8 -*-
import os
from win32com import client
# pip instatll win32com
import pdfmerge


def doc2pdf(doc_name, pdf_name):
    """ word文件转pdf

    :param doc_name: word文件名称
    :param pdf_name: 转换后pdf文件名称

    """
    try:
        word = client.DispatchEx("Word.Application")
        if os.path.exists(pdf_name):
            os.remove(pdf_name)
        worddoc = word.Documents.Open(doc_name, ReadOnly=1)
        worddoc.SaveAs(pdf_name, FileFormat=17)
        worddoc.Close()
        word.Quit()
        return pdf_name
    except Exception as e:
        print(e)
        return 1


def doc2pdf(folder_name):
    """ 将文件夹中的 word 全部转 pdf

    :param folder_name: 文件夹名称
    :return: 转换后 pdf 所在文件夹
    """
    items = [x for x in os.listdir(folder_name)]
    input = folder_name.split('\\')[-1]
    output = "output_pdf"
    output_path = folder_name.replace(input, output)
    os.makedirs(output_path, exist_ok=True)

    try:
        for file in items:
            # 只遍历 word 文件
            if os.path.splitext(file)[1] == ".docx":
                file_path = "%s\\%s" % (folder_name, file)
                pdf_path = file_path.replace(
                    "docx", "pdf").replace(input, output)
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                word = client.DispatchEx("Word.Application")
                print(file_path)
                worddoc = word.Documents.Open(file_path, ReadOnly=1)
                worddoc.SaveAs(pdf_path, FileFormat=17)
                worddoc.Close()
                word.Quit()
        return output_path

    except Exception as e:
        print(e)
        return ""


def mergepdf(foldername):
    ''' 合并当前文件夹下的所有 pdf 文件

    :param foldername: 文件夹路径

    '''

    files = ["%s\\%s" % (foldername, file) for file in os.listdir(
        foldername) if os.path.splitext(file)[1] == '.pdf']

    for file in files:
        pdfmerge.add_blank_page(file)

    pdfmerge.merge(files)


if __name__ == '__main__':
    '''
    1 为 word 转 pdf
    2 为 merge 所有 pdf
    '''

    prompt = '''
    1 word to pdf
    2 merge pdf
    please enter a number:
    '''
    while True:
        num = input(prompt)

        if num == '1':
            folder = input("Please enter the folder path:")
            print(doc2pdf(folder))

        elif num == '2':
            folder = input("Please enter the folder path:")
            mergepdf(folder)

        elif num == "":
            break

        else:
            print("error input")