# This will extract from a list of PDF files in a directory.

import PyPDF4 as pypdf
import pprint

from datetime import datetime
import time
import os, glob
import csv
import argparse


# ----------------------------------------
# Function definitions
# ----------------------------------------
# allow user supplied path, if not supplied
# default path is current working directory
def main(path, out_path, filename):

    # get all the files in current working directory
    files_in_directory = glob.glob(os.path.join(path, '*.pdf'))

    # start extracting from files. using list comprehension for readability
    output = [ extract_record(file) for file in files_in_directory ]
    print("Completed extracting {} files".format(len(output)))

    # save extracted data to file with today as timestamp
    save_as = "_report_{}.csv".format(datetime.today().strftime('%Y-%m-%d'))
    save_to_file(output, save_as)

# extract a record
def extract_record(file):
    # retrieve the timestamp and filename
    document_info = { "timestamp": time.ctime(os.path.getmtime(file))
                    , "filename": file.split("/")[-1]}

    try:
        pdfobject = open(file,'rb')
        pdf = pypdf.PdfFileReader(pdfobject)
        data = pdf.getFields()

        if isinstance(data, dict):
            return document_info | {k:v.get('/V', '') for (k,v) in data.items()}

    except:
        pass

    return document_info

# saves extracted data to file
def save_to_file(data, output_filename):
    with open(output_filename, 'w', newline='')  as output_file:
        dict_writer = csv.DictWriter(output_file, list(set().union(*data)))
        dict_writer.writeheader()
        dict_writer.writerows(data)

    print("Saved results to {}".format(output_filename))


# ----------------------------------------
# Program entry point
# ----------------------------------------
# Accept user supplied arguments
# -d  allows user to specify directory to perform extraction

parser = argparse.ArgumentParser(description='Text goes here.')
parser.add_argument('-d', '--directory', metavar='D', required=False, default=os.getcwd(),
                    help="directory path where files are located", action="store")
parser.add_argument('-o', '--output-directory', metavar='O', required=False,
                    action="store", help="output directory")
parser.add_argument('-f', '--filename', default="_report", required=False, action="store",
                    help='give a filename to output. otherwise it will be _report_YYYY-MM-DD.csv')

args = parser.parse_args()

# start from main method
main(path=args.directory,
    out_path=vars(args).get("output-directory", ""),
    filename=args.filename)
