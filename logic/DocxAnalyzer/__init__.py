""
from enum import Enum
from typing import NamedTuple
import zipfile
import os
import glob
import xmltodict
import shutil
# with zipfile.ZipFile("sukhareva.docx", 'r') as zip_ref:
#    zip_ref.extractall("extracted")


class style_enum(Enum):
    PAR = "paragraph"
    CHAR = "character"
    NUM = "numbering"
    TAB = "table"
    DXA = "dxa"


class style(NamedTuple):
    id: str
    based_on: str
    is_custom: bool
    style_type: style_enum
    name: str
    priority: int
    fonts: list
    color: str
    spacing: str
    line_rule: str


# creates a dir with files extracted from .docx file
def extract_file(filename: str) -> str:
    output_dir = filename[:len(filename)-5] + "_extracted"
    if os.path.isdir(output_dir):
        shutil.rmtree(output_dir)
    os.mkdir(output_dir)

    with zipfile.ZipFile(filename, 'r') as zip_ref:
        zip_ref.extractall(output_dir)
    return output_dir


def dict_extract_kvs(d):
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
    doc = xmltodict.parse(open(xmlfile, "r", errors="replace").read())
    kvs = dict_extract_kvs(doc)
    for kv in kvs:
        if kv[0] == "@w:ascii" and \
                not (kv[1] == "Times New Roman" or kv[1] == "Cambria Math"):
            return False
    return True


def check_A4(xmlfile) -> bool:
    f = open(xmlfile, "r", errors="replace")
    a = f.read().split("<")
    for line in a:
        if "w:pgSz w:w=\"11906\" w:h=\"16838\"" in line:
            return True
        if "w:pgSz w:w=\"11907\" w:h=\"16840\"" in line:
            return True
    return False


def has_centered_footer(xmlfile) -> bool:
    return 'w:jc w:val="center"' in open(xmlfile, "r", errors="replace").read()


def has_key_words(xmlfile) -> dict:
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
    return res


def check_links(xmlfile):
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

    return result


def backup_link(xmlfile):
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

    return result


def check_table(xmlfile):
    f = open(xmlfile, "r", errors="replace")
    a = f.read().split("<")
    for line in a:
        if "w:tbl" in line:
            return True
    return False


def check_size(xmlfile):
    f = open(xmlfile, "r", errors="replace")
    a = f.read().split("<")
    for line in a:
        if "w:sz w:val=\"26\"" in line:
            return True
        if "w:sz w:val=\"28\"" in line:
            return True
    return False


def check_interval(xmlfile):
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

    return abzac and spacing_auto


def check_drawing(xmlfile):
    f = open(xmlfile, "r", errors="replace")
    a = f.read().split("<")
    for line in a:
        if "w:drawing" in line:
            return True
    return False


def main(path):
    result = dict()

    docx_dir = extract_file(path)
    styles_file = docx_dir + "/word/styles.xml"
    document_file = docx_dir + "/word/document.xml"
    footer_files = glob.glob(docx_dir + "/word/footer*.xml")

    key_words = has_key_words(document_file)
    try:
        links = check_links(document_file)
    except KeyError as e:
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
    print(result, links)
    return result
