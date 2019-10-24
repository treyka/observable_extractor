#!/usr/bin/env python3

from docopt import docopt
from mimetypes import guess_type
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import textract
import xlrd 
import re



__doc__ = """observable_extractor.py :: extract observables from various file formats

Usage:
    observable_extractor.py --input=INPUT
"""


def pdf_to_txt(file_):
    '''convert a pdf to text'''
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(file_, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching, check_extractable=True):
        interpreter.process_page(page)
    txt = retstr.getvalue()
    fp.close()
    device.close()
    retstr.close()
    if type(txt) == 'bytes':
        return(txt.decode("utf-8"))
    else:
        return(txt)

def excel_to_txt(file_):
    '''convert an excel to text'''
    book = xlrd.open_workbook(file_)
    txt = str()
    i = 0
    while i < book.nsheets:
        sheet = book.sheet_by_index(i)
        row = 0
        while row < sheet.nrows:
            line = str()
            col = 0
            while col < sheet.ncols:
                line += sheet.cell(row, col).value + ' '
                col += 1
            txt += line + '\n'
            row += 1
        i += 1
    if type(txt) == 'bytes':
        return(txt.decode("utf-8"))
    else:
        return(txt)


def docx_to_txt(file_):
    '''convert .docx to text'''
    txt = textract.process(file_, extension='docx')
    # docx = docx_doc(file_)
    # txt = str()
    # for i in docx.paragraphs:
    #     txt += ' ' + i.text.encode('utf-8')
    if type(txt) == 'bytes':
        return(txt.decode("utf-8"))
    else:
        return(txt)


def doc_to_txt(file_):
    '''convert .doc to text'''
    txt = textract.process(file_, extension='doc')
    # doc = doc_doc(file_)
    # # fix carriage returns
    # txt = re.sub(r'\r', r'\r\n', doc.read())
    if type(txt) == 'bytes':
        return(txt.decode("utf-8"))
    else:
        return(txt)


def observables_from_txt(txt):
    '''extract observables from text'''
    observables = {}
    observables['sha512'] = set()
    for sha512 in re.findall(r'\b[0-9a-f]{128}\b', txt, re.IGNORECASE):
        observables['sha512'].add(sha512)
    observables['sha256'] = set()
    for sha256 in re.findall(r'\b[0-9a-f]{64}\b', txt, re.IGNORECASE):
        observables['sha256'].add(sha256)
    observables['sha1'] = set()
    for sha1 in re.findall(r'\b[0-9a-f]{40}\b', txt, re.IGNORECASE):
        observables['sha1'].add(sha1)
    observables['md5'] = set()
    for md5 in re.findall(r'\b[0-9a-f]{32}\b', txt, re.IGNORECASE):
        observables['md5'].add(md5)
    observables['ip'] = set()
    for ip in re.findall(r'\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b', txt, re.IGNORECASE):
        observables['ip'].add(ip)
    observables['url'] = set()
    for url in re.findall(r'\b(http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+)\b', txt, re.IGNORECASE):
        observables['url'].add(url)
    return(observables)



if __name__ == '__main__':
    args = docopt(__doc__)
    file_ = args['--input']
    filetype = guess_type(file_)[0]
    if filetype in ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
        observables = observables_from_txt(excel_to_txt(file_))
    elif filetype in ['application/vnd.openxmlformats-officedocument.wordprocessingml.document']:
        observables = observables_from_txt(docx_to_txt(file_))
    elif filetype in ['application/msword']:
        observables = observables_from_txt(doc_to_txt(file_))
    elif filetype in ['application/pdf']:
        observables = observables_from_txt(pdf_to_txt(file_))
    elif filetype in ['text/plain', 'text/csv']:
        txt_file = open(file_)
        txt = txt_file.read()
        txt_file.close
        observables = observables_from_txt(txt)
    else: 
        observables = None
        print(filetype)
    # for key in observables.keys():
    #     print('%s: %i' % (key, len(observables[key])))
    for key in observables.keys():
        for i in observables[key]:
            print('%s: %s' % (key, i))
