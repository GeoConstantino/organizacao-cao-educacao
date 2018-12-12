import csv
import difflib
import os
import re
import sys
import zipfile

from collections import namedtuple, defaultdict
from operator import itemgetter
from subprocess import Popen, PIPE

import docx
import numpy as np

from docx.opc.exceptions import PackageNotFoundError


RELEVANT_PATTERN = '(?<=MPRJ CAO Educação):?\s?\d{1,}'


def read_craais():
    with open('craais.csv', 'r', encoding='utf8') as fobj:
        reader = csv.reader(fobj, delimiter=',')
        craais = {}
        for craai in reader:
            for city in craai:
                if city:
                    regional = craai[0]
                    if city != 'CAMPOS':
                        craais[city] = regional

    return craais


def read_city_csv():
    return [
        row.decode('latin-1').strip() for row in
        open('lista_municipio.csv', 'rb')][1:]


def read_keywords():
    rows = []
    for row in open('cardapio.csv'):
        rows.append(row.strip().split(';'))

    return {r[0]: r[1].lower().strip() for r in rows}


def translate(city_name, cities):
    city_name_lower = city_name.lower()
    if city_name_lower in cities:
        return city_name

    ratios = []
    for city in cities:
        ratios.append(
            difflib.SequenceMatcher(
                None, city_name_lower, city.lower()).ratio()
        )

    ratios = np.array(ratios)
    return cities[ratios.argmax()]


def read_content(complete_path):
    content = ''
    if complete_path.endswith('.doc'):
        ext = '.doc'
        content = get_doc_content(complete_path)
    elif complete_path.endswith('.docx'):
        content = get_docx_content(complete_path)
        ext = '.docx'

    return content, ext


def get_first_page(content):
    ident_line = re.search(
        RELEVANT_PATTERN,
        content,
        re.IGNORECASE
    )
    return content[:ident_line.span()[1]]


def get_doc_content(complete_path):
    process_output = Popen(
        ['antiword', complete_path],
        stdout=PIPE
    )
    out_content = process_output.stdout.read().decode('utf-8')
    return re.sub('\s{1,}', ' ', out_content)


def get_docx_content(complete_path):
    try:
        docx_obj = docx.Document(complete_path)
    except PackageNotFoundError:
        return ''

    return re.sub(
        '\s{1,}',
        ' ',
        '\n'.join([p.text for p in docx_obj.paragraphs])
    )


def get_school_name(complete_path):
    return (complete_path.split('/')[-2]).strip()


def get_folder_class(content, year):
    folder_class = re.search(
        '(Análise|Parecer) técnico-pedagógic[ao].{1,5}\d{1,}.{1,3}/2018',
        content,
        re.IGNORECASE
     )
    if folder_class is None:
        return "class_nao_encontrada"
    return folder_class.group(0).strip()


def create_file_name(content, complete_path, year):
    school_name = get_school_name(complete_path)
    folder_class = get_folder_class(content, year)
    file_name = (school_name + " " + folder_class).lower()
    patterns = [
        ('(\(|\)|\s+|/|-)', '_'),
        ('[!?.]', ''),
        ('_{1,}', '_')
    ]
    for pat in patterns:
        file_name = re.sub(*pat, string=file_name)

    return file_name + '.zip'


def create_output_path(output, key_word, content, complete_path, year):
    return os.path.join(output, create_file_name(content, complete_path, year))


def is_document_relevant(complete_path, content):
    """
        Para decidir se um documento é relevante, utilizamos os seguintes
        critérios:

            - Contém a palavra Relatório
            - Contém as palavras Dados internos
            - Possui número MPRJ ou Possui informação de Pedaggoda
    """
    has_matricula = re.search(
        RELEVANT_PATTERN,
        content,
        re.IGNORECASE
    )
    n_levels = re.split(r'\d{4}', complete_path)[1].count('/') >= 4
    not_anexo = 'anexo' not in complete_path.lower()

    is_relevant = (has_matricula is not None and n_levels and not_anexo)

    return is_relevant


def last_modified_document(cur, paths):
    doc_info = namedtuple(
        'DocInfo',
        ['path', 'content', 'l_modified', 'name', 'pdf_path']
    )
    last_modified = []
    for path in paths:
        complete_path = os.path.join(os.path.join(cur, path))

        if ((path.endswith('.doc') or path.endswith('.docx'))
                and not path.startswith('~$')):
            content, ext = read_content(complete_path)

            if is_document_relevant(complete_path, content):
                try:
                    last_modified.append(
                        doc_info(
                            complete_path,
                            content,
                            os.path.getmtime(complete_path),
                            os.path.basename(complete_path),
                            complete_path.replace(ext, '.pdf')
                        )
                    )
                except OSError:
                    continue

    if last_modified:
        return sorted(last_modified, key=itemgetter(2), reverse=True)[0]


def find_city(content):
    city = re.search(
        '(?<=Munic[ií]pio)\s*:\s*\w+(\w+\s*)+',
        content,
        re.IGNORECASE
    )
    if city is None:
        city = re.search(
            '(?<=Munic[ií]pio)\s*(do|de|da)s?\s+\w+(\w+\s*)+',
            content,
            re.IGNORECASE
        )

    if city is None:
        city = re.search('\w+(\w\s*)+(?=/RJ)', content, re.IGNORECASE)

    if city is not None:
        city_name = city.group(0)
    else:
        city_name = ''

    return city_name


def clean_city_name(city_name):
    prepared_name = city_name.strip()
    prepared_name = re.sub(':\s?', '', city_name)
    prepared_name = prepared_name.split('\n')[0]
    de_match = re.match('\s*(de|do|da)\s*', prepared_name)
    if de_match is not None:
        prepared_name = prepared_name.replace(de_match.group(0), '', 1)

    return prepared_name


def find_keyord(content, keywords):
    content = content.lower()
    for key, value in keywords.items():
        if value in content:
            return key.strip()

    return 'OUTRAS TEMÁTICAS'


def zipdir(doc_info, ziph):
    path = doc_info.path
    path = os.path.abspath(os.path.join(path, os.pardir))

    parent_dir = os.path.sep.join(path.split(os.path.sep))
    pdf_basename = os.path.basename(doc_info.pdf_path)
    for root, dirs, files in os.walk(path):
        for file in files:
            path_diff = os.path.relpath(root, path)

            if root == parent_dir and (
                    file != doc_info.name and file != pdf_basename):
                continue
            ziph.write(
                os.path.join(root, file),
                arcname=os.path.join(path_diff, file)
            )


def zip_files(complete_output_path, doc_info):
    zipf = zipfile.ZipFile(complete_output_path, 'w', zipfile.ZIP_DEFLATED)
    zipdir(doc_info, zipf)


keyword_classes = defaultdict(int)

if __name__ == '__main__':
    root_path = sys.argv[1]
    year = sys.argv[2]
    # Pedagógica ou Contábil
    doc_type = sys.argv[3]

    cities = read_city_csv()
    keywords = read_keywords()

    output_path = 'out/{year}/{craai}/{city}/' + doc_type + '/{keyword}'

    errors = []
    for cur, dirs, files in os.walk(root_path):
        doc_info = last_modified_document(cur, files)

        if doc_info is not None:
            # Get city name
            city_name = translate(
                clean_city_name(find_city(doc_info.content)), cities
            )

            keyword = find_keyord(get_first_page(doc_info.content), keywords)
            keyword_classes[keyword] += 1
            try:
                os.makedirs(output_path.format(
                    year=year,
                    city=city_name,
                    keyword=keyword
                ))
            except FileExistsError:
                pass

            # Create output path
            complete_output_path = create_output_path(
                output_path,
                keyword,
                doc_info.content,
                doc_info.path,
                year
            ).format(year=year, city=city_name, keyword=keyword)

            zip_files(complete_output_path, doc_info)

            if 'class_nao_encontrada' in complete_output_path:
                errors.append(doc_info.path)
