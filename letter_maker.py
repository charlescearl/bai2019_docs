"""Generate BAI acceptance letters from a template.

This assumes that you have Python 3.7 installed.

Python package are
- pandas==0.25
- pdfkit==0.6
- defopt==5.1
- tika

These can be added with pip,

  pip install pandas==0.25 pdfkit==0.6 defopt==5.1 tika

The library wkhtmltopdf is also required.
On OSX this call be installed as:

   brew install caskroom/cask/wkhtmltopdf

Run this by invoking
   python ./letter_maker.py main -a <HTML ABSTRACT TEMPLATE FILE> -t <TRAVEL GRANT TEMPLATE> -a <CSV OF AUTHORS>

This assumes that the images subdirectory from the template file is in the
directory from which you are running this script.

By installing pylint, you can also run the test file test_letter_maker.py by running

   pytest ./test_letter_maker.py
"""
import re
import pandas as pd
import pdfkit
import defopt
import tempfile

_DEFAULT_TRAVEL_GRANT = 'yoshua_travel_grant_template.html'
_DEFAULT_ABSTRACT = 'yoshua_abstract_template.html'
_DEFAULT_AUTHOR = 'info.csv'

_PAPER_TITLE_COL = 'Paper Title'
_PAPER_SUBMIT_COL = 'PassportRequired'
_HOME_ADDR_COL = 'Residential Address'
_WORK_ADDR_COL = 'Work/School Address'
_WORK_TITLE_COL = 'Job Title'
_EMAIL_COL = 'Email'
_COUNTRY_COL = 'Passport Issuing Country'
_PASSPORT_COL = 'Passport Number'
_NAME_COL = 'Name'


def main(template_abstract=_DEFAULT_ABSTRACT, template_grant=_DEFAULT_TRAVEL_GRANT, author_info=_DEFAULT_AUTHOR):
    """Build letters from html template and csv of author info.

    :param str template_abstract: Path to the html template file for authors who submitted an abstract.
    :param str template_grant: Path to the html template file for authors who asked for travel grant.
    :param str author_info: CSV file containing the information for accepted authors.
    :return:
    """

    with open(template_abstract) as input_handle:
        abstract_file = input_handle.read()
    with open(template_grant) as input_handle:
        travel_grant_file = input_handle.read()
    cols = [_NAME_COL, _PASSPORT_COL, _COUNTRY_COL, _EMAIL_COL, _WORK_TITLE_COL,
            _WORK_ADDR_COL, _HOME_ADDR_COL, _PAPER_SUBMIT_COL, _PAPER_TITLE_COL]
    df = pd.read_csv(author_info, na_values=['-'], usecols=cols, sep='\t')  # pylint: disable=invalid-name
    #empty_data = get_incomplete_records(df)
    #empty_data.to_csv('./missing_data_entries.tsv', header=True, index=False, sep='\t')
    # Just process the ones for which these columns have values
    # df = clean_records(df)  # pylint: disable=invalid-name
    df = df.fillna(value=' ')
    emails_to_send = []
    for rec in df.to_records(index=None):
        if rec[_PAPER_SUBMIT_COL] == 'No':
            print(f'{rec["Name"]} did not require a passport')
            #continue
        # Check that the Passport is a valid number -- no whitespace and some numbers present
        if not valid_passport(rec):
            print(f'Passport {rec["Passport Number"]} is not valid.')
        abstract = None
        if not has_abstract(rec):
            template_file = travel_grant_file
            print(f'{rec["Name"]} did not submit an abstract')
        else:
            template_file = abstract_file
            abstract = rec[_PAPER_TITLE_COL]
        doc_text = modify_template(rec, template_file, abstract)
        doc_name = f'{rec["Name"].replace(" ", "_")}'
        with open('./scratch_file.html', mode='w') as file_handle:
            file_handle.write(doc_text)
            pdfkit.from_url('./scratch_file.html', f'./pdfs/{doc_name}.pdf')
            emails_to_send.append({'email': rec['Email'], 'file': f'./pdfs/{doc_name}.pdf'})
    pd.DataFrame.from_dict(emails_to_send).to_csv('./emails_to_send.csv', index=None)



def has_abstract(rec):
    """Check that we have an abstract."""
    return rec[_PAPER_TITLE_COL] is not None and rec[_PAPER_TITLE_COL] != ' '


def valid_passport(rec):
    """Check that there is a valid passport field."""
    return re.match(r'[A-Za-z]{0,2}[0-9]+', rec[_PASSPORT_COL])


def get_incomplete_records(df):  # pylint: disable=invalid-name
    """Return data frame having records with empty Name, Passport, Nationality, Email or Address fields."""
    return df[df[[_NAME_COL, _PASSPORT_COL, _COUNTRY_COL, _EMAIL_COL, _HOME_ADDR_COL, ]].isnull().any(axis=1)]


def clean_records(df):  # pylint: disable=invalid-name
    """Filter the records."""
    df = df.dropna(subset=[_NAME_COL, _PASSPORT_COL, _COUNTRY_COL, _EMAIL_COL, _HOME_ADDR_COL, ])
    return df.fillna(value=' ')


def modify_template(rec, template_file, abstract):
    """Create a new document using the template and if needed abstract.

    Parameters
    ----------
    rec : object
        The df record containing the data.
    template_file : str
        The template to modify.
    abstract : str
        If present the abstract to use.

    Returns
    -------
    str
        The document to use.
    """
    doc_text = template_file.replace('##Name##', rec[_NAME_COL])
    doc_text = doc_text.replace('##Name of Applicant##', rec[_NAME_COL])
    doc_text = doc_text.replace('##Email Address##', rec[_EMAIL_COL])
    doc_text = doc_text.replace('##Company Address##', rec[_WORK_ADDR_COL])
    doc_text = doc_text.replace('##Occupation##', rec[_WORK_TITLE_COL])
    doc_text = doc_text.replace('##Passport Number##', rec[_PASSPORT_COL])
    doc_text = doc_text.replace('##Passport Issuing Country##', rec[_COUNTRY_COL])
    doc_text = doc_text.replace('##Residential Address##', rec[_HOME_ADDR_COL])
    if abstract and '##Abstract Title##' in doc_text:
        doc_text = doc_text.replace('##Abstract Title##', abstract)
    #doc_text = doc_text.replace('##Flight cost##', f'all roundtrip costs')
    return doc_text


if __name__ == '__main__':
    defopt.run(main, short={'template_abstract': 'a', 'template_grant': 'T', 'author_info': 'i'}, strict_kwonly=True)
