import argparse
import json
import glob
import os
import progressbar
import requests
from openpyxl import load_workbook


refresh_date = "2021-03-15"


validator_url = "http://stage.iativalidator.iatistandard.org/api/v1/stats?date=2020-12-31"
all_validation = json.loads(requests.get(validator_url).content)
wb = load_workbook(filename = 'template.xlsx')
sheet = wb['Sheet1']


if __name__ == "__main__":
    arg_parser = argparse.ArgumentParser(description='Create publisher metadata')
    arg_parser.add_argument('publisher', type=str, help='Publisher\'s ID from the IATI Registry')
    args = arg_parser.parse_args()
    output_dir = os.path.join("output", args.publisher)
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)

    sheet['A1'] = "Organization: {}".format(args.publisher)
    sheet['A2'] = "AUTOMATIC DATA PULL - {}".format(refresh_date)

    pub_validation = [val for val in all_validation if val['publisher'] == args.publisher]
    critical_errors = pub_validation[0]["summaryStats"]["critical"]
    
    sheet['B8'] = critical_errors

    xml_path = os.path.join("/home/alex/git/IATI-Registry-Refresher/data", args.publisher, '*')
    xml_files = glob.glob(xml_path)
    # bar = progressbar.ProgressBar()
    # for xml_file in bar(xml_files):

    outfile = os.path.join(output_dir, "publisher.xlsx")
    wb.save(outfile)
        