# -*- encoding: utf-8 -*-
import os
from win32com import client
# pip instatll win32com
import pdfmerge


def doc2pdf(doc_name, pdf_name):
    """
    :word文件转pdf
    :param doc_name word文件名称
    :param pdf_name 转换后pdf文件名称
    """
    try:
        word = client.DispatchEx("Word.Application")
        if os.path.exists(pdf_name):
            os.remove(pdf_name)
        worddoc = word.Documents.Open(doc_name, ReadOnly=1)
        worddoc.SaveAs(pdf_name, FileFormat=17)
        worddoc.Close()
        return pdf_name
    except Exception as e:
        print(e)
        return 1


def doc2pdf(folder_name):
    """
    haha 
    """
    items = [x for x in os.listdir(folder_name)]
    try:
        for file in items:
            if os.path.splitext(file)[1] == ".docx":

                file_path = "%s\\%s" % (folder_name, file)
                pdf_path = file_path.replace("docx", "pdf", 1)
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                word = client.DispatchEx("Word.Application")
                print(file_path)
                worddoc = word.Documents.Open(file_path, ReadOnly=1)
                worddoc.SaveAs(pdf_path, FileFormat=17)
                worddoc.Close()

    except Exception as e:
        print(e)
        return 1


if __name__ == '__main__':
    # foldname = r"F:\Git\mijigenerator\MijiGenerator\bin\Debug\output_doc"
    # doc2pdf(foldname)
    file = r"F:\Git\mijigenerator\MijiGenerator\bin\Debug\output_doc\20180416 过敏体质渐露.pdf"
    folder = r"F:\Git\mijigenerator\MijiGenerator\bin\Debug\output_doc"
    for file in os.listdir(folder):
        if os.path.splitext(file)[1] == ".pdf":
            pdfmerge.add_blank_page(folder + "\\" + file)