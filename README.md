SKOS <---> Excel
================

A tool to assist Arches translators.

Previous work: 

This script is currently in development. Its purpose is to take an Arches thesaurus in SKOS(RDF/XML) format and spit out an Excel file, optionally translating into an additional language, much like [`po2excel`](https://github.com/ads04r/po2excel) did.

The reverse script (not written yet) will take the checked
spreadsheet and recreate the thesaurus for importing back into
Arches.

Install dependencies

```bash
pip3 install -r requirements.txt
```
## From XML to XLSX

Convert `EAMENA.xml` to French, as XLSX

```bash
py skos2excel.py ./data/EAMENA.xml ./data/EAMENA_fr.xlsx -lang fr -f xlsx 
```

## From XLSX to XML

```bash
py excel2skos.py ./data/EAMENA_fr.xlsx ./data/EAMENA_fr.xml -b ./data/EAMENA.xml
```
