import PyPDF2
import re
import pandas as pd

number_re = re.compile('(?P<pv_size>P{1}\d{3})')
intro_re = re.compile('(Introduction:?)')
background_re = re.compile('(Background:?)')
objective_re = re.compile('(Objective:?)')
title_re = re.compile('(P{1}\d{3}\s?([A-Z]{1,}[\s,-:;\?()\*@]*\n?)+)')
univer_re = re.compile('(\d{1}[A-Za-z,]{2,}[ \w]*[,]?)')
session_name_list = []
session_title_list = []
session_intro_list = []
session_author_list = []
session_univer_list = []


def author_list(file):
    """
    This function returns all authors published in this book
    from the Index topic of the pdf file
    (page 64 and page 65)
    :param file:
    :return: authors_list
    Example
    ['Florence Abdallah', 'Jill Abell', 'William Abramovits', 'Ken Abrams',]
    """
    file_split = file.split()
    file_text = ' '.join(file_split)
    topic_authors = re.findall(r'([A-Z]{1}[a-z]+, ?[A-Za-z-]* ?[A-Za-z-]* )', file_text)
    authors_list = []
    for s in topic_authors:
        authors_list.append(' '.join(s.replace(' ', '').split(',')[::-1]))
    return authors_list


def parse_author_affiliations(author_par):
    """
    This function extract data
    (many authors - many affiliations)
    :return: dict {author: affiliations}
    Example:
    {'Herbert Pang' : 'Yale University'}
    """
    compare_author = author_list(authors_dict)
    checked_author_list = []
    data = dict()
    s = re.findall(r'([A-Z]{1}[a-z]* [A-Z]{1}[a-z-]*)', author_par)
    for i in s:
        if i in compare_author:
            checked_author_list.append(i)
            data[i] = s[-1]
    if not checked_author_list:
        data[s[0]] = s[-1]
    return data


def without_abstract(topic_author):
    """
    For Topic without Introduction, Objective, Background
    only text in Presentation Abstract
    """
    global last_word
    s1 = re.findall(r'(, ?[A-Z]{1}[a-z]* [A-Za-z ]*)', topic_author)
    last_word = re.search(s1[-1], topic_author).end() + 1
    return last_word


def presentation_abstract(file, topic_title):
    """
    Parsing Abstract with
    Introduction, Objective, Background
    """
    topic_intro = [m.start() for m in intro_re.finditer(file)]
    topic_background = [m.start() for m in background_re.finditer(file)]
    topic_objective = [m.start() for m in objective_re.finditer(file)]
    article_text = None
    if topic_intro:
        article_text = [m.start() for m in intro_re.finditer(file)]
    elif topic_background:
        article_text = [m.start() for m in background_re.finditer(file)]
    elif topic_objective:
        article_text = [m.start() for m in objective_re.finditer(file)]
    elif not article_text:
        article_text = [topic_title +
                        without_abstract(file[topic_title:topic_title + 120])]
    return article_text[0]


def parse_article(article_content, start, end):
    """
    Function parse and extract data from the article
    :param article_content:
    :param Start: start of the article
    :param end: end of the article
    :return: all data information from the all articles
    """
    file = article_content[start:end]
    topic_number = [file[m.start():m.end()] for m in number_re.finditer(file)][0]
    topic_title = [m.end() for m in title_re.finditer(file)][0]
    topic_intro = presentation_abstract(file, topic_title)
    authors = parse_author_affiliations(file[topic_title-1:topic_intro])
    for i in authors:
        session_author_list.append(i)
        session_univer_list.append(authors[i])
        session_name_list.append(topic_number)
        session_title_list.append(file[4:topic_title - 1])
        session_intro_list.append(file[topic_intro:end])
    data_from_artickles = {
        'Session Name ': session_name_list,
        'Topic Title': session_title_list,
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
        global authors_dict
        authors_dict = pdf_reader.getPage(64).extractText() + \
                       pdf_reader.getPage(65).extractText()
        return pdf_content


if __name__ == '__main__':
    file = parse_pdf_book('Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf')
    file1 = file.split()
    file22 = ' '.join(file1)
    file21 = file22.replace('™', '').replace('Œ', '').replace('ﬂ', '').replace('ﬁ', '')
    file2 = file21.replace('P1032', 'p10320').replace('P05495', 'p"05495"') \
        .replace('Sarajevowith ', 'Sarajevo Introduction: Many ancient')
    topic_number1 = [m.start() for m in number_re.finditer(file2)]
    topic_number1.append(383556)
    start = topic_number1[0]
    for i in topic_number1[1:]:
        data_to_excel = parse_article(file2, start, i)
        start = i
    df = pd.DataFrame(data_to_excel)
    df.to_excel('program_output.xlsx')

