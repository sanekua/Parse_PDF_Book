import PyPDF2
import re
import pandas as pd

number_re = re.compile('(?P<pv_size>P{1}\d{3})')
intro_re = re.compile('(Introduction:?)')
title_re = re.compile('(P{1}\d{3}\s?([A-Z]{1,}[\s,-:;]*\n?)+)')
author_re = re.compile('([ ]?\w* [\w-]+[-,\n\a\b,\t,\v]? [\w-]+ \d{1})')
univer_re = re.compile('(\d{1}[A-Za-z,]{2,}[ \w]*[,]?)')
session_name_list = []
session_title_list = []
session_intro_list = []
session_author_list = []
session_univer_list = []


def author(topic_author):
    data = dict()
    s = re.findall(r'([A-Z]{1}[a-z]+)', topic_author)
    topic_author = ' '.join(s)
    ans2 = re.compile('(\w{2,} [\w.,]*)').findall(topic_author)
    try:
        data[ans2[0]] = ans2[-1]
    except:
        data['None'] = 'None'
    return data


def parse_author_affiliations(file):
    """
    This function extract data
    (many authors - many affiliations)
    :param file:
    :return:
    """
    topic_author = [author_re.findall(file)][0]
    topic_affiliation = [univer_re.findall(file)][0]
    affiliation_list = dict()
    author_list = dict()
    for i in topic_affiliation:
        if i[0] in '123456789':
            affiliation_list[i[0]] = (i[1:])
    for i in topic_author:
        if i[-1] in affiliation_list.keys():
            author_list[i[:-1]] = affiliation_list[i[-1]]
    if not author_list:
        author_list = author(file)
    return author_list


def parse_article(article_content, start, end):
    """
    Function parse and extract data from the article
    :param article_content:
    :param start: start of the article
    :param end: end of the article
    :return: all data information from the all articles
    """
    file = article_content[start:end]
    topic_number = [file[m.start():m.end()] for m in number_re.finditer(file)][0]
    topic_title = [m.end() for m in title_re.finditer(file)][0]
    topic_intro = [m.start() for m in intro_re.finditer(file)]
    topic_intro = topic_intro[0] if topic_intro else topic_title + 150
    authors = parse_author_affiliations(file[topic_title-1:topic_intro])
    for i in authors:
        session_author_list.append(i)
        session_univer_list.append(authors[i])
        session_name_list.append(topic_number)
        session_title_list.append(file[4:topic_title-1])
        session_intro_list.append(file[topic_intro:end])
    data_from_artickles = {
        'Session Name ': session_name_list,
        'Topic Title':  session_title_list,
        'Name (incl. titles if)': session_author_list,
        'Affiliation(s) Name(s)': session_univer_list,
        'Presentation abstract': session_intro_list,
    }
    return data_from_artickles


def parse_pdf_book(book):
    """
    This function return pdf book content.
    start at page 6, end at page 63
    """
    with open(book, 'rb') as pdfFileObj:
        pdf_reader = PyPDF2.PdfFileReader(pdfFileObj)
        pdf_content = ''
        for page in range(6, 64):
            pdf_content += pdf_reader.getPage(page).extractText()
        return pdf_content


if __name__ == '__main__':
    file = parse_pdf_book('Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf')
    file1 = file.split()
    file22 = ' '.join(file1)
    file21 = file22.replace('™', '')
    file2 = file21.replace('P1032', 'p"1032"').replace('P05495', 'p"05495"')
    topic_number1 = [m.start() for m in number_re.finditer(file2)]
    start = topic_number1[0]
    for i in topic_number1[1:]:
        all_data = parse_article(file2, start, i)
        start = i
    df = pd.DataFrame(all_data)
    df.to_excel('program_output.xlsx')
