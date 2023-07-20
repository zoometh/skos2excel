#!/usr/bin/python3
# SKOS to Excel
# By Ash Smith, MarEA project, University of Southampton
# Based on PO2Excel, by Thomas Huet and Ash Smith

import os, re, csv, sys, argparse, json
from deep_translator import GoogleTranslator
from deep_translator.exceptions import NotValidPayload
from progress.bar import IncrementalBar as progressbar
from openpyxl import Workbook
from rdflib import Graph, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, SKOS, DCTERMS as DCT
from hashlib import md5

argp = argparse.ArgumentParser()
argp.add_argument('SKOSFile', metavar='skos_file', type=str, help='The SKOS file to convert')
argp.add_argument('CSVFile', metavar='csv_file', type=str, help='The CSV file to write')
argp.add_argument('-lang', '--language', action='store', type=str, help='The 2-letter code for the language to use', default='')
argp.add_argument('-f', '--format', action='store', type=str, help='The file format to export, xlsx or csv (default is csv)', default='csv')
args = argp.parse_args()

target_language = args.language
target_format = args.format
path_fold = os.getcwd()
input_file = args.SKOSFile
output_file = args.CSVFile

g = Graph()
with open(input_file, 'r') as fp:
	data = fp.read()
g.parse(data=data, format='xml')

target_format = target_format.lower()
if target_format == 'xls':
	target_format = 'xlsx'
if target_format != 'xlsx':
	target_format = 'csv'

ret = []

ret.append(['id', 'language', 'text'])

for s, p, o in g.triples((None, RDF.type, SKOS.Concept)):

	uri = str(s)
	id = uri.split('/')[-1]
	trans = {}
	for ss, pp, oo in g.triples((s, SKOS.prefLabel, None)):

		data = json.loads(str(oo))
		lang = str(oo.language)
		if lang == 'en-us':
			lang = 'en'
		if not('value' in data):
			continue

		trans[lang] = data['value']

	if len(trans) == 0:
		continue

	if target_language != '':
		if (('en' in trans) & (not(target_language in trans))):
			try:
				trans[target_language] = GoogleTranslator(source='en', target=target_language).translate(trans['en'])
			except NotValidPayload:
				trans[target_language] = ''

	ret.append(['', '', ''])
	for k in trans.keys():
		lang = str(k)
		ret.append([id, lang, trans[lang]])

if target_format == 'csv':

        with open(output_file, 'w') as csv_file:
                csv_writer = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                for item in ret:
                        csv_writer.writerow(item)

if target_format == 'xlsx':

        wb = Workbook()
        ws = wb.active
        for item in ret:
                ws.append(item)

        wb.save(output_file)
