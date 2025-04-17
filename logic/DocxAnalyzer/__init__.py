""
from enum import Enum
from typing import NamedTuple
import zipfile
import os
import glob
import xmltodict
import shutil
import logging

logger = logging.getLogger(__name__)
logging.basicConfig(filename='docxAnalyzer.log', encoding='utf-8', level=logging.DEBUG)


# creates a dir with files extracted from .docx file
def extract_file(filename: str) -> str:
    logger.info("entered extract_file")
    output_dir = filename[:len(filename)-5] + "_extracted"
    if os.path.isdir(output_dir):
        shutil.rmtree(output_dir)
    logger.info("creating directory", output_dir, "to unpack to")
    os.mkdir(output_dir)

    with zipfile.ZipFile(filename, 'r') as zip_ref:
        zip_ref.extractall(output_dir)
    return output_dir


def dict_extract_kvs(d):
    logger.info("entered dict_extract_kvs")
    stack = [("initial_dict", d)]
    kvs = []

    while stack:
        curr = stack.pop()
        if isinstance(curr, tuple):
            if len(curr) == 2:
                if isinstance(curr[1], dict):
                    stack.extend(curr[1].items())
                elif isinstance(curr[1], list):
                    stack.extend(curr[::-1])
                else:
                    kvs.append(curr)
        elif isinstance(curr, dict):
            stack.extend(curr.items())
        elif isinstance(curr, list):
            stack.extend(curr[::-1])
    return kvs


def check_fonts(xmlfile) -> bool:
    logger.info("entered check_fonts")
    doc = xmltodict.parse(open(xmlfile, "r", errors="replace").read())
    kvs = dict_extract_kvs(doc)
    for kv in kvs:
        if kv[0] == "@w:ascii" and \
                not (kv[1] == "Times New Roman" or kv[1] == "Cambria Math"):
            return False
    return True


def check_A4(xmlfile) -> bool:
    logger.info("entered check_A4")
    f = open(xmlfile, "r", errors="replace")
    a = f.read().split("<")
    for line in a:
        if "w:pgSz w:w=\"11906\" w:h=\"16838\"" in line:
            return True
        if "w:pgSz w:w=\"11907\" w:h=\"16840\"" in line:
            return True
    return False


def has_centered_footer(xmlfile) -> bool:
    logger.info("entered has_centered_footer")
    return 'w:jc w:val="center"' in open(xmlfile, "r", errors="replace").read()


def has_key_words(xmlfile) -> dict:
    logger.info("entered has_key_words")
    contents = False
    intro = False
    outro = False
    shortens = False
    termins = False
    appendix = False
    sources = False

    t = open(xmlfile, "r", encoding="utf-8").read().upper()

    if "СОДЕРЖАНИЕ" in t:
        contents = True
    if "ОГЛАВЛЕНИЕ" in t:
        contents = True
    if "ТЕРМИН" in t:
        termins = True
    if "ВВЕДЕНИЕ" in t:
        intro = True
    if "ЗАКЛЮЧЕНИЕ" in t:
        outro = True
    if "СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ" in t:
        shortens = True
    if "СОКРАЩЕНИЯ" in t:
        shortens = True
    if "ПРИЛОЖЕНИЕ" in t:
        appendix = True
    if "ИСТОЧНИКОВ" in t:
        if "СПИСОК" in t:
            sources = True

    res = dict()
    res["СОДЕРЖАНИЕ"] = contents
    res["СПИСОК ТЕРМИНОВ И ОПРЕДЕЛЕНИЙ"] = termins
    res["ВВЕДЕНИЕ"] = intro
    res["ЗАКЛЮЧЕНИЕ"] = outro
    res["СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ"] = shortens
    res["ПРИЛОЖЕНИЕ"] = appendix
    res["СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"] = sources
    res["все"] = (
        contents and termins and intro and outro and shortens and appendix and sources)
    logger.info("Results:")
    logger.info(res)
    return res


def check_links(xmlfile):
    logger.info("entered check_links")
    anchors = []
    words = []
    result = dict()
    result["has_custom_par"] = False
    result["has_numbering"] = True
    result["extra_contents"] = False
    document = xmltodict.parse(open(xmlfile, "r", errors="replace").read())
    content = document
    for p in content:
        cur_text = ""
        vk = dict_extract_kvs(p)
        for el, txt in vk:
            if el == 'w:t':
                cur_text += txt.upper()
        words.append(cur_text)

        for k, v in vk:
            if k == '@w:anchor':
                anchors.append(v)

    if not ("СОДЕРЖАНИЕ" in words[0]) or ("ОГЛАВЛЕНИЕ" in words[0]):
        result["extra_contents"] = True
    for t in words:
        if "СОДЕРЖАНИЕ" in t:
            continue
        if "ОГЛАВЛЕНИЕ" in t:
            continue
        if "СПИСОК ТЕРМИНОВ И ОПРЕДЕЛЕНИЙ" in t:
            continue
        if "ВВЕДЕНИЕ" in t:
            continue
        if "ЗАКЛЮЧЕНИЕ" in t:
            continue
        if "СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ" in t:
            continue
        if "ПРИЛОЖЕНИЕ" in t:
            continue
        if "ИСТОЧНИКОВ" in t:
            if "СПИСОК" in t:
                continue
        result["has_custom_par"] = True

    for n in words:
        # if contents has number of page: make numbering invalid
        if n == "":
            continue
        elif ("СОДЕРЖАНИЕ" in n) or ("ОГЛАВЛЕНИЕ" in n):
            if (n[-1:].isdigit()):
                print("ALARM AT CONT", n)
                result["has_numbering"] = False
                break
        # if anything else doesn't have a page number it's also invalid
        elif not (n[-1:].isdigit()):
            print("ALARM AT SHIT", n, n[-1:], n[-1:].isdigit())
            result["has_numbering"] = False
            break
    logger.info("Result:")
    logger.info(result)
    return result


def backup_link(xmlfile):
    logger.info("entered backup_link")
    anchors = []
    words = []
    result = dict()
    result["has_custom_par"] = False
    result["has_numbering"] = True
    result["extra_contents"] = False
    document = xmltodict.parse(open(xmlfile, "r", errors="replace").read())
    content = document["w:document"]["w:body"]["w:p"]
    for p in content:
        cur_text = ""
        vk = dict_extract_kvs(p)
        for el, txt in vk:
            if el == 'w:t':
                cur_text += txt.upper()
        words.append(cur_text)

        for k, v in vk:
            if k == '@w:anchor':
                anchors.append(v)

    if not ("СОДЕРЖАНИЕ" in words[0]) or ("ОГЛАВЛЕНИЕ" in words[0]):
        result["extra_contents"] = True
    for t in words:
        if "СОДЕРЖАНИЕ" in t:
            continue
        if "ОГЛАВЛЕНИЕ" in t:
            continue
        if "СПИСОК ТЕРМИНОВ И ОПРЕДЕЛЕНИЙ" in t:
            continue
        if "ВВЕДЕНИЕ" in t:
            continue
        if "ЗАКЛЮЧЕНИЕ" in t:
            continue
        if "СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ" in t:
            continue
        if "ПРИЛОЖЕНИЕ" in t:
            continue
        if "ИСТОЧНИКОВ" in t:
            if "СПИСОК" in t:
                continue
        result["has_custom_par"] = True

    for n in words:
        # if contents has number of page: make numbering invalid
        if n == "":
            continue
        elif ("СОДЕРЖАНИЕ" in n) or ("ОГЛАВЛЕНИЕ" in n):
            if (n[-1:].isdigit()):
                print("ALARM AT CONT", n)
                result["has_numbering"] = False
                break
        # if anything else doesn't have a page number it's also invalid
        elif not (n[-1:].isdigit()):
            print("ALARM AT SHIT", n, n[-1:], n[-1:].isdigit())
            result["has_numbering"] = False
            break
    logger.info("Result:")
    logger.info(result)

    return result


def check_table(xmlfile):
    logger.info("entered check_table")
    f = open(xmlfile, "r", errors="replace")
    a = f.read().split("<")
    for line in a:
        if "w:tbl" in line:
            return True
    return False


def check_size(xmlfile):
    logger.info("entered check_size")
    f = open(xmlfile, "r", errors="replace")
    a = f.read().split("<")
    for line in a:
        if "w:sz w:val=\"26\"" in line:
            return True
        if "w:sz w:val=\"28\"" in line:
            return True
    return False


def check_interval(xmlfile):
    logger.info("entered check_interval")
    f = open(xmlfile, "r", errors="replace")
    a = f.read().split("<")
    abzac = False
    spacing_auto = False
    for line in a:
        if "lineRule" in line:
            if "auto" in line:
                spacing_auto = True
        if "w:firstLine=\"709\"" in line:
            abzac = True
    logger.info("result: abzac is ", abzac, " spacing_auto is ", spacing_auto)
    return abzac and spacing_auto


def check_drawing(xmlfile):
    logger.info("entered check_interval")
    f = open(xmlfile, "r", errors="replace")
    a = f.read().split("<")
    for line in a:
        if "w:drawing" in line:
            return True
    return False


def main(path):
    logger.info("entered main func of docxAnalyzer")
    result = dict()

    docx_dir = extract_file(path)
    styles_file = docx_dir + "/word/styles.xml"
    document_file = docx_dir + "/word/document.xml"
    footer_files = glob.glob(docx_dir + "/word/footer*.xml")

    key_words = has_key_words(document_file)
    try:
        links = check_links(document_file)
    except KeyError as e:
        logger.warn("failed to do a regular link check, doing backup_link check")
        links = backup_link(document_file)

    footer = any([has_centered_footer(x) for x in footer_files])

    result["Введение и заключение"] = (
        key_words["ВВЕДЕНИЕ"] and key_words["ЗАКЛЮЧЕНИЕ"])
    result["Заголовки"] = key_words["все"]
    result["Иллюстрации"] = check_drawing(document_file)
    result["Интервалы текста"] = check_interval(styles_file)
    result["Навигация"] = check_links(document_file)
    result["Нумерация"] = (footer and links["has_numbering"])
    result["Приложение"] = key_words["ПРИЛОЖЕНИЕ"]
    result["Разделы и подразделы"] = (
        links["has_custom_par"] and (not links["extra_contents"]))
    result["Размеры полей"] = check_interval(styles_file)
    result["Состав содержания"] = (
        all([x for x in key_words.values()]) and links["has_custom_par"])
    result["Список источников"] = key_words["СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]
    result["Список сокращений"] = key_words["СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ"]
    result["Список терминов"] = key_words["СПИСОК ТЕРМИНОВ И ОПРЕДЕЛЕНИЙ"]
    result["Таблицы"] = check_table(document_file)
    result["Формат листа"] = check_A4(document_file)
    result["Шрифт"] = check_fonts(document_file)
    logger.info("Check finished. Results:", result, links)
    return result
