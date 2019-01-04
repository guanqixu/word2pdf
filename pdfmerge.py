from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter


def get_reader(filename):
    ''' 获取 pdfreader

    :param filename: 文件名称
    :return: pdfreader
    :rtype: PdfFileReader

    '''
    try:
        pdf_reader = PdfFileReader(filename)

    except Exception as e:
        print(e)

    return pdf_reader


def add_blank_page(filename):

    ''' 添加一个空白页

    只有在页数不是偶数的情况下，才添加
    :param filename： 文件名
    
    '''
    reader = get_reader(filename)

    writer = PdfFileWriter()
    writer.appendPagesFromReader(reader)

    pagecount = writer.getNumPages()
    if pagecount % 2 != 0:
        writer.addBlankPage()

    writer.write(open(filename, 'wb'))


def merge(filenames):

    ''' 合并多个pdf文件

    :param filenames: 文件集合
    :return: 是否成功合并
    
    '''

    merger = PdfFileMerger()

    try:
        for file in filenames:
            merger.append(file)
        merger.write("merge.pdf")

    except Exception as e:
        print(e)
        return False

    return True
