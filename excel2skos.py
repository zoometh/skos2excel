#!/usr/bin/python3
# Excel to SKOS
# By Ash Smith, MarEA project, University of Southampton

import os, re, csv, sys, argparse, json
from deep_translator import GoogleTranslator
from deep_translator.exceptions import NotValidPayload
from progress.bar import IncrementalBar as progressbar
from openpyxl.cell import Cell, MergedCell
from openpyxl import load_workbook
from rdflib import Graph, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, SKOS, DCTERMS as DCT
from uuid import uuid4

argp = argparse.ArgumentParser()
argp.add_argument('CSVFile', metavar='csv_file', type=str, help='The CSV or XLSX file to convert')
argp.add_argument('SKOSFile', metavar='skos_file', type=str, help='The SKOS file to write')
argp.add_argument('-b', '--base', action='store', type=str, help='The SKOS file to use as a base', default='')
args = argp.parse_args()

base_file = args.base
input_file = args.CSVFile
output_file = args.SKOSFile
mapping = {}

print(base_file)

g = Graph()
if os.path.exists(base_file):
	print(f"Base is: {base_file}")
	with open(base_file, 'r') as fp:
		data = fp.read()
	g.parse(data=data, format='xml')

sys.exit
data = []

if input_file.lower().endswith('.xlsx'):
	wb = load_workbook(input_file)
	ws = wb.active
	for row in ws.rows:
		item = []
		for cell in row:
			if cell.value is None:
				item.append('')
			else:
				item.append(str(cell.value))
		data.append(item)

if input_file.lower().endswith('.csv'):
	with open(input_file) as csvfile:
		csvreader = csv.reader(csvfile, delimiter=',', quotechar='"')
		for row in csvreader:
			data.append(row)

translations = {}
headers = []
for row in data:
	if len(headers) == 0:
		headers = row
		continue
	if len(row) > len(headers):
		continue
	item = {}
	itemlen = 0
	for i in range(0, len(headers)):
		k = str(headers[i])
		item[k] = row[i]
		itemlen = itemlen + len(row[i])
	if itemlen == 0:
		continue
	if not('id' in item):
		continue
	if not('language' in item):
		continue
	id = item['id']
	lang = item['language']
	text = item['text']
	if id == '':
		continue
	if lang == '':
		continue
	if text == '':
		continue
	if not(id in translations):
		translations[id] = {}
	translations[id][lang] = text

inserts = []
ct = 0 ; print_each = 100 ; g_len = len(g)
for s, p, o in g.triples((None, RDF.type, SKOS.Concept)):
	ct = ct + 1
	do = ct % print_each
	if do == 0:
		print(f"Read record: {ct}/{g_len}")
	uri = str(s)
	id = uri.split('/')[-1]
	trans = translations[id]
	for ss, pp, oo in g.triples((s, SKOS.prefLabel, None)):

		data = json.loads(str(oo))
		lang = str(oo.language)
		if lang == 'en-us':
			lang = 'en'
		if not('value' in data):
			continue
		if lang in trans:
			del(trans[lang])

	if len(trans) == 0:
		continue

	for k in trans.keys():
		lang = str(k)
		value = trans[lang]
		id = str(uuid4())
		data = json.dumps({'id': id, 'value': value})
		literal = Literal(data, lang=lang)

		inserts.append((s, SKOS.prefLabel, literal))

for item in inserts:
	g.add(item)

with open(output_file, 'w') as fp:
	fp.write(g.serialize(format='pretty-xml'))
print("Exported")
