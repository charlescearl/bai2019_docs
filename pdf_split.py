import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import pandas as pd


def pdf_splitter(path):
    fname = os.path.splitext(os.path.basename(path))[0]
    df = pd.read_csv('AllAwardees.tsv', index_col=None, sep='\t').fillna(value=' ')
    df.index = df.index.rename("I")
    user_names = [f'submission_{rec["I"]}.pdf' for rec in df.to_records()]

    pdf = PdfFileReader(path)
    for page in range(pdf.getNumPages()-1):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))

        output_filename = user_names[page]

        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)

        print('Created: {}'.format(output_filename))

def _fname(rec):
    return '_'.join(str(rec['First Name']).split()+str(rec['Last Name']).split())


if __name__ == '__main__':
    path = './AwardsWithAbstracts.pdf'
    pdf_splitter(path)
