#!/usr/bin/env python

#########
# Extracts all data/metadata from GenBank records and saves it as an Excel spreadsheet
#########

# flexible dictionary classes
import collections as co
# handles GenBank files
from Bio import SeqIO
import sys
# for saving in Excel format
import openpyxl
import pandas as pd

# https://stackoverflow.com/a/13654229/1706987
# Avoids going into error if no sequence is found
# for example in metagenomic or genomic seqs
def wrapper(gen):
  while True:
    try:
      yield next(gen)
    except StopIteration:
      break
    except Exception as e:
      print(e)

# extracts all data into a into a list
# each element of a list is a dictionary with all
# of the specified info
def extract_all_data(x, taxnames):
    counter = 0
    data = []
    # loop each record
    for record in x:
        # "growable dictionary"
        datadict = co.defaultdict()
        # splice taxonomy field into list
        for i in range(len(taxnames)):
            datadict[taxnames[i]] = record.annotations['taxonomy'][i+3]
        # extract all features available from the "source" block
        # save them to the dictionary
        for i in list(record.features[0].qualifiers.keys()):
            datadict[i] = record.features[0].qualifiers[i][0]
        # construct the source fieldf
        # no ref saves as "NA"
        source = []
        try:
            ref = record.annotations['references']
        except KeyError:
            datadict['source'] = "NA"
        else:
            for i in ref:
                source.append(i.journal)
            datadict['source'] = "; ".join(source)
        datadict['accession'] = record.id              # accession number
        datadict['description'] = record.description   # description field
        datadict['sequence'] = record.seq              # sequence data
        data.append(datadict)
        counter += 1
    return data,counter

# gets a list if unique column names for the final table
def extract_col_names(x):
    # use ordered dictionary
    colnames = co.OrderedDict()
    for i in x:
        for j in list(i.keys()):
            colnames[j] = ''
    return list(colnames.keys())

# generate final table
def write_final_dataframe(records_dict, column_names, outputfile):
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
    pd.DataFrame.from_dict(final_datafile).to_excel(outputfile)

# prepare data file
if __name__ == "__main__":
    outputfile = "GenBank_output_table.xlsx"
    # input file first argument
    input_file = sys.argv[1]
    # parse records as iterators
    records = wrapper(SeqIO.parse(input_file, "gb"))
    # names of taxonomic groups
    taxon_names = ["phylum","class","order","family","genus"]
    # get all records
    records_dict,number = extract_all_data(records, taxon_names)
    print("records extracted:", number)
    # get columns
    column_names = extract_col_names(records_dict)
    print("fields extracted:", len(column_names))
    # write file
    write_final_dataframe(records_dict, column_names, outputfile)
    print("output file saved to:",outputfile)
