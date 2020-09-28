import collections as co
import pandas as pd
from Bio import SeqIO
import sys
import openpyxl

def extract_all_data(x, taxnames):
    data = []
    for record in x:
        datadict = co.defaultdict()
        for i in range(len(taxnames)):
            datadict[taxnames[i]] = record.annotations['taxonomy'][i+3]
        for i in list(record.features[0].qualifiers.keys()):
            datadict[i] = record.features[0].qualifiers[i][0]
        source = []
        for i in record.annotations['references']:
            source.append(i.journal)
        datadict['source'] = "; ".join(source)
        datadict['accession'] = record.id
        datadict['description'] = record.description
        data.append(datadict)
    return data

def extract_col_names(x):
    colnames = co.OrderedDict()
    for i in x:
        for j in list(i.keys()):
            colnames[j] = ''
    return list(colnames.keys())

# generate table
def write_final_dataframe(records_dict, column_names):
    final_datafile = co.OrderedDict()
    # initialize dictionary
    for col in column_names:
        final_datafile[col] = []
    # fill in dict
    for record in records_dict:
        for col in column_names:
            if record.get(col):
                final_datafile[col].append(record[col])
            else:
                final_datafile[col].append('NA')
    # make data frame in pandas
    pd.DataFrame.from_dict(final_datafile).to_excel("output_table.xlsx")

# prepare data file
input_file = sys.argv[1]
records = SeqIO.parse(input_file, "gb")
taxon_names = ["phylum","class","order","family","genus"]
records_dict = extract_all_data(records, taxon_names)
column_names = extract_col_names(records_dict)
write_final_dataframe(records_dict, column_names)
