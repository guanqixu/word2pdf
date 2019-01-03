from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter


def get_reader(filename):
    try:
        pdf_reader = PdfFileReader(filename)

    except Exception as e:
        print(e)

    return pdf_reader


def add_blank_page(filename):
    reader = get_reader(filename)

    writer = PdfFileWriter()
    writer.appendPagesFromReader(reader)

    pagecount = writer.getNumPages()
    if pagecount % 2 != 0:
        writer.addBlankPage()

    writer.write(open(filename, 'wb'))
