import argparse
import glob
import json
from lxml import etree
from lxml.etree import XMLParser
from openpyxl import load_workbook
import os
import progressbar
import requests


refresh_date = "2021-03-15"
validator_url = "http://stage.iativalidator.iatistandard.org/api/v1/stats?date={}".format(refresh_date)
all_validation = json.loads(requests.get(validator_url).content)
large_parser = XMLParser(huge_tree=True)
parser = etree.XMLParser(remove_blank_text=True)
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

    number_of_activities = 0

    xml_path = os.path.join("/home/alex/git/IATI-Registry-Refresher/data", args.publisher, '*')
    xml_files = glob.glob(xml_path)
    bar = progressbar.ProgressBar()
    for xml_file in bar(xml_files):
        print(xml_file)
        tree = etree.parse(xml_file, parser=large_parser)
        root = tree.getroot()
        activities = root.xpath("iati-activity")
        number_of_activities += len(activities)
    
    sheet['B5'] = number_of_activities

    outfile = os.path.join(output_dir, "publisher.xlsx")
    wb.save(outfile)
        