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
        # (Name, excel location, xpath/formula, function),
        ("Number of activities", "B5", "iati-activity", "len"),
        ("Number of activities with financials", "B6", "iati-activity[transaction|budget|planned-disbursement]", "len"),
        ("Number of activities with financials 2018", "C6", "iati-activity[transaction[transaction-date[starts-with(@iso-date, '2018')]]|budget[period-start[starts-with(@iso-date, '2018')]]|planned-disbursement[period-start[starts-with(@iso-date, '2018')]]]", "len"),
        ("Number of activities with financials 2019", "D6", "iati-activity[transaction[transaction-date[starts-with(@iso-date, '2019')]]|budget[period-start[starts-with(@iso-date, '2019')]]|planned-disbursement[period-start[starts-with(@iso-date, '2019')]]]", "len"),
        ("Number of activities with financials 2020", "E6", "iati-activity[transaction[transaction-date[starts-with(@iso-date, '2020')]]|budget[period-start[starts-with(@iso-date, '2020')]]|planned-disbursement[period-start[starts-with(@iso-date, '2020')]]]", "len"),
        ("Number of activities with financials 2021", "F6", "iati-activity[transaction[transaction-date[starts-with(@iso-date, '2021')]]|budget[period-start[starts-with(@iso-date, '2021')]]|planned-disbursement[period-start[starts-with(@iso-date, '2021')]]]", "len"),
        ("Number of activities with financials 2022", "G6", "iati-activity[transaction[transaction-date[starts-with(@iso-date, '2022')]]|budget[period-start[starts-with(@iso-date, '2022')]]|planned-disbursement[period-start[starts-with(@iso-date, '2022')]]]", "len"),
        ("Number of activities with financials 2023", "H6", "iati-activity[transaction[transaction-date[starts-with(@iso-date, '2023')]]|budget[period-start[starts-with(@iso-date, '2023')]]|planned-disbursement[period-start[starts-with(@iso-date, '2023')]]]", "len"),
        ("Number of transactions", "B7", "iati-activity/transaction", "len"),
        ("Number of transactions in 2018", "C7", "iati-activity/transaction[transaction-date[starts-with(@iso-date, '2018')]]", "len"),
        ("Number of transactions in 2019", "D7", "iati-activity/transaction[transaction-date[starts-with(@iso-date, '2019')]]", "len"),
        ("Number of transactions in 2020", "E7", "iati-activity/transaction[transaction-date[starts-with(@iso-date, '2020')]]", "len"),
        ("Number of transactions in 2021", "F7", "iati-activity/transaction[transaction-date[starts-with(@iso-date, '2021')]]", "len"),
        ("Number of transactions in 2022", "G7", "iati-activity/transaction[transaction-date[starts-with(@iso-date, '2022')]]", "len"),
        ("Number of transactions in 2023", "H7", "iati-activity/transaction[transaction-date[starts-with(@iso-date, '2023')]]", "len"),

        ("Participating orgs with narrative and ref", "B14", "iati-activity/participating-org[@ref and narrative]", "len"),
        ("Participating orgs with ref", "B15", "iati-activity/participating-org[@ref]", "len"),
        ("Participating orgs with narrative", "B16", "iati-activity/participating-org[narrative]", "len"),
        ("Number of activities with funding org", "B18", "iati-activity[participating-org[@role='1']]", "len"),
        ("Percentage of activities with funding org", "B19", "round((indicator_values['Number of activities with funding org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities with accountable org", "B21", "iati-activity[participating-org[@role='2']]", "len"),
        ("Percentage of activities with accountable org", "B22", "round((indicator_values['Number of activities with accountable org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities with extending org", "B24", "iati-activity[participating-org[@role='3']]", "len"),
        ("Percentage of activities with extending org", "B25", "round((indicator_values['Number of activities with extending org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities with implementing org", "B27", "iati-activity[participating-org[@role='4']]", "len"),
        ("Percentage of activities with implementing org", "B28", "round((indicator_values['Number of activities with implementing org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions that include provider org", "B32", "iati-activity/transaction[provider-org]", "len"),
        ("Percentage of transactions that include provider org", "B33", "round((indicator_values['Number of transactions that include provider org'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions with provider org only ref", "B34", "iati-activity/transaction[provider-org[@ref and not(narrative)]]", "len"),
        ("Number of transactions with provider org only narrative", "B35", "iati-activity/transaction[provider-org[not(@ref) and narrative]]", "len"),
        ("Number of transactions with provider org ref and narrative", "B36", "iati-activity/transaction[provider-org[@ref and narrative]]", "len"),
        ("Number of transactions with provider org provider-activity-id", "B37", "iati-activity/transaction[provider-org[@provider-activity-id]]", "len"),
        ("Number of transactions that include receiver org", "B39", "iati-activity/transaction[receiver-org]", "len"),
        ("Percentage of transactions that include receiver org", "B40", "round((indicator_values['Number of transactions that include receiver org'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions with receiver org only ref", "B41", "iati-activity/transaction[receiver-org[@ref and not(narrative)]]", "len"),
        ("Number of transactions with receiver org only narrative", "B42", "iati-activity/transaction[receiver-org[not(@ref) and narrative]]", "len"),
        ("Number of transactions with receiver org ref and narrative", "B43", "iati-activity/transaction[receiver-org[@ref and narrative]]", "len"),
        ("Number of transactions with receiver org receiver-activity-id", "B44", "iati-activity/transaction[receiver-org[@receiver-activity-id]]", "len"),

        ("Number of transactions with type 1", "B48", "iati-activity/transaction[transaction-type[@code='1']]", "len"),
        ("Number of activities with transactions with type 1", "B49", "iati-activity[transaction[transaction-type[@code='1']]]", "len"),
        ("Percentage of activities with transactions with type 1", "B50", "round((indicator_values['Number of activities with transactions with type 1'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 2", "B52", "iati-activity/transaction[transaction-type[@code='2']]", "len"),
        ("Number of activities with transactions with type 2", "B53", "iati-activity[transaction[transaction-type[@code='2']]]", "len"),
        ("Percentage of activities with transactions with type 2", "B54", "round((indicator_values['Number of activities with transactions with type 2'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 3", "B56", "iati-activity/transaction[transaction-type[@code='3']]", "len"),
        ("Number of activities with transactions with type 3", "B57", "iati-activity[transaction[transaction-type[@code='3']]]", "len"),
        ("Percentage of activities with transactions with type 3", "B58", "round((indicator_values['Number of activities with transactions with type 3'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 4", "B60", "iati-activity/transaction[transaction-type[@code='4']]", "len"),
        ("Number of activities with transactions with type 4", "B61", "iati-activity[transaction[transaction-type[@code='4']]]", "len"),
        ("Percentage of activities with transactions with type 4", "B62", "round((indicator_values['Number of activities with transactions with type 4'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 5", "B64", "iati-activity/transaction[transaction-type[@code='5']]", "len"),
        ("Number of activities with transactions with type 5", "B65", "iati-activity[transaction[transaction-type[@code='5']]]", "len"),
        ("Percentage of activities with transactions with type 5", "B66", "round((indicator_values['Number of activities with transactions with type 5'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 6", "B68", "iati-activity/transaction[transaction-type[@code='6']]", "len"),
        ("Number of activities with transactions with type 6", "B69", "iati-activity[transaction[transaction-type[@code='6']]]", "len"),
        ("Percentage of activities with transactions with type 6", "B70", "round((indicator_values['Number of activities with transactions with type 6'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 7", "B72", "iati-activity/transaction[transaction-type[@code='7']]", "len"),
        ("Number of activities with transactions with type 7", "B73", "iati-activity[transaction[transaction-type[@code='7']]]", "len"),
        ("Percentage of activities with transactions with type 7", "B74", "round((indicator_values['Number of activities with transactions with type 7'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 8", "B76", "iati-activity/transaction[transaction-type[@code='8']]", "len"),
        ("Number of activities with transactions with type 8", "B77", "iati-activity[transaction[transaction-type[@code='8']]]", "len"),
        ("Percentage of activities with transactions with type 8", "B78", "round((indicator_values['Number of activities with transactions with type 8'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 9", "B80", "iati-activity/transaction[transaction-type[@code='9']]", "len"),
        ("Number of activities with transactions with type 9", "B81", "iati-activity[transaction[transaction-type[@code='9']]]", "len"),
        ("Percentage of activities with transactions with type 9", "B82", "round((indicator_values['Number of activities with transactions with type 9'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 10", "B84", "iati-activity/transaction[transaction-type[@code='10']]", "len"),
        ("Number of activities with transactions with type 10", "B85", "iati-activity[transaction[transaction-type[@code='10']]]", "len"),
        ("Percentage of activities with transactions with type 10", "B86", "round((indicator_values['Number of activities with transactions with type 10'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 11", "B88", "iati-activity/transaction[transaction-type[@code='11']]", "len"),
        ("Number of activities with transactions with type 11", "B89", "iati-activity[transaction[transaction-type[@code='11']]]", "len"),
        ("Percentage of activities with transactions with type 11", "B90", "round((indicator_values['Number of activities with transactions with type 11'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 12", "B92", "iati-activity/transaction[transaction-type[@code='12']]", "len"),
        ("Number of activities with transactions with type 12", "B93", "iati-activity[transaction[transaction-type[@code='12']]]", "len"),
        ("Percentage of activities with transactions with type 12", "B94", "round((indicator_values['Number of activities with transactions with type 12'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 13", "B96", "iati-activity/transaction[transaction-type[@code='13']]", "len"),
        ("Number of activities with transactions with type 13", "B97", "iati-activity[transaction[transaction-type[@code='13']]]", "len"),
        ("Percentage of activities with transactions with type 13", "B98", "round((indicator_values['Number of activities with transactions with type 13'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of budgets", "B101", "iati-activity/budget", "len"),
        ("Number of budgets in 2018", "C101", "iati-activity/budget[period-start[starts-with(@iso-date, '2018')]]", "len"),
        ("Number of budgets in 2019", "D101", "iati-activity/budget[period-start[starts-with(@iso-date, '2019')]]", "len"),
        ("Number of budgets in 2020", "E101", "iati-activity/budget[period-start[starts-with(@iso-date, '2020')]]", "len"),
        ("Number of budgets in 2021", "F101", "iati-activity/budget[period-start[starts-with(@iso-date, '2021')]]", "len"),
        ("Number of budgets in 2022", "G101", "iati-activity/budget[period-start[starts-with(@iso-date, '2022')]]", "len"),
        ("Number of budgets in 2023", "H101", "iati-activity/budget[period-start[starts-with(@iso-date, '2023')]]", "len"),
        ("Number of activities that contain budgets", "B102", "iati-activity[budget]", "len"),
        ("Percentage of activities that contain budgets", "B103", "round((indicator_values['Number of activities that contain budgets'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        
    ]

    indicator_values = dict()

    xml_path = os.path.join("/home/alex/git/IATI-Registry-Refresher/data", args.publisher, '*')
    xml_files = glob.glob(xml_path)
    bar = progressbar.ProgressBar()
    for xml_file in bar(xml_files):
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
        if indicator_function == "eval":
            indicator_values[indicator_name] = eval(indicator_xpath)
        accumulated_value = indicator_values[indicator_name]
        if type(accumulated_value) is list:
            accumulated_value = ", ".join(accumulated_value)
        sheet[indicator_location] = accumulated_value

    outfile = os.path.join(output_dir, "publisher.xlsx")
    wb.save(outfile)
        