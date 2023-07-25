SKOS <---> Excel
================

A tool to assist Arches translators.

Previous work: https://github.com/ads04r/po2excel

This script is currently in development. Its purpose is to take an
Arches thesaurus in SKOS(RDF/XML) format and spit out an Excel
file, optionally translating into an additional language, much
like po2excel did.

The reverse script, `excel2skos.py`, will take the checked
spreadsheet and recreate the thesaurus for importing back into
Arches. It requires the original SKOS file as a base, and only
adds triples, it never replaces or deletes existing them. The
base file is considered to be truth.

Install dependencies

```bash
pip3 install -r requirements.txt
```
Convert to French, as CSV

```bash
py skos2excel.py ./data/EAMENA.xml ./data/EAMENA_out.csv -lang fr -f csv 
```
Convert to French, as XLSX

```bash
py skos2excel.py ./data/EAMENA.xml ./data/EAMENA_out.xlsx -lang fr -f xlsx 
```
