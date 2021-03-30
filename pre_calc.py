import argparse
import glob
import json
from lxml import etree
from lxml.etree import XMLParser
from openpyxl import load_workbook
import os
import progressbar
import requests


refresh_date = "2021-03-29"
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

    indicators = [
        # (Name, excel location, xpath, function),
        ("Number of activities", "B5", "iati-activity", "len"),
        ("Number of activities with financials", "B6", "iati-activity[transaction|budget|planned-disbursement]", "len"),
        ("Number of activities with financials 2018", "C6", "iati-activity[transaction[transaction-date[starts-with(@iso-date, '2018')]]|budget[period-start[starts-with(@iso-date, '2018')]]|planned-disbursement[period-start[starts-with(@iso-date, '2018')]]]", "len")
    ]

    indicator_values = dict()

    xml_path = os.path.join("/home/alex/git/IATI-Registry-Refresher/data", args.publisher, '*')
    xml_files = glob.glob(xml_path)
    bar = progressbar.ProgressBar()
    for xml_file in bar(xml_files):
        print(xml_file)
        tree = etree.parse(xml_file, parser=large_parser)
        root = tree.getroot()
        
        for indicator_name, indicator_location, indicator_xpath, indicator_function in indicators:
            if indicator_function == "len":
                evaluated_value = len(root.xpath(indicator_xpath))
            else:
                evaluated_value = []
            if indicator_name not in indicator_values.keys():
                indicator_values[indicator_name] = evaluated_value
            else:
                indicator_values[indicator_name] += evaluated_value

    for indicator_name, indicator_location, indicator_xpath, indicator_function in indicators:
        accumulated_value = indicator_values[indicator_name]
        if type(accumulated_value) is list:
            accumulated_value = ", ".join(accumulated_value)
        sheet[indicator_location] = accumulated_value

    outfile = os.path.join(output_dir, "publisher.xlsx")
    wb.save(outfile)
        