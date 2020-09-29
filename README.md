# extractGenBankToExcel

## Description

This Python code extracts data and metadata from a GenBank record flat file and outputs an Excel spreadsheet. It uses BioPython, Collections, and Pandas internally.

## Requirements

* Python3
* [BioPython](https://biopython.org/)
* Collections
* [Pandas](https://pandas.pydata.org/)

## Download the code

### Linux:

```bash
wget https://raw.githubusercontent.com/santiagosnchez/extractGenBankToExcel/master/extractGenBankToExcel.py
```

## Install dependencies

```bash
pip install --user biopython collections pandas openpyxl
```

## Running the code

```bash
python extractGenBankToExcel.py /path/to/genbank_file.gb
```
