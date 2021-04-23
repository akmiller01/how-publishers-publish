import argparse
import glob
import json
from lxml import etree
from lxml.etree import XMLParser
from openpyxl import load_workbook
import os
import progressbar
import requests


def destroy_tree(tree):
    root = tree.getroot()

    node_tracker = {root: [0, None]}

    for node in root.iterdescendants():
        parent = node.getparent()
        node_tracker[node] = [node_tracker[parent][0] + 1, parent]

    node_tracker = sorted([(depth, parent, child) for child, (depth, parent)
                           in node_tracker.items()], key=lambda x: x[0], reverse=True)

    for _, parent, child in node_tracker:
        if parent is None:
            break
        parent.remove(child)

    del tree


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

        ("Participating orgs with narrative and ref", "B14", "iati-activity/participating-org[(string-length(@ref) > 0) and (narrative[string-length(text()) > 0])]", "len"),
        ("Participating orgs with ref", "B15", "iati-activity/participating-org[string-length(@ref) > 0]", "len"),
        ("Participating orgs with narrative", "B16", "iati-activity/participating-org[narrative[string-length(text()) > 0]]", "len"),
        ("Number of activities with funding org", "B18", "iati-activity[participating-org[@role='1']]", "len"),
        ("Percentage of activities with funding org", "B19", "round((indicator_values['Number of activities with funding org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities with accountable org", "B21", "iati-activity[participating-org[@role='2']]", "len"),
        ("Percentage of activities with accountable org", "B22", "round((indicator_values['Number of activities with accountable org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities with extending org", "B24", "iati-activity[participating-org[@role='3']]", "len"),
        ("Percentage of activities with extending org", "B25", "round((indicator_values['Number of activities with extending org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities with implementing org", "B27", "iati-activity[participating-org[@role='4']]", "len"),
        ("Percentage of activities with implementing org", "B28", "round((indicator_values['Number of activities with implementing org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions that include provider org", "B32", "iati-activity/transaction[provider-org[(string-length(@ref) > 0) or (narrative[string-length(text()) > 0])]]", "len"),
        ("Percentage of transactions that include provider org", "B33", "round((indicator_values['Number of transactions that include provider org'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions with provider org only ref", "B34", "iati-activity/transaction[provider-org[(string-length(@ref) > 0) and not(narrative[string-length(text()) > 0])]]", "len"),
        ("Number of transactions with provider org only narrative", "B35", "iati-activity/transaction[provider-org[not(string-length(@ref) > 0) and (narrative[string-length(text()) > 0])]]", "len"),
        ("Number of transactions with provider org ref and narrative", "B36", "iati-activity/transaction[provider-org[(string-length(@ref) > 0) and (narrative[string-length(text()) > 0])]]", "len"),
        ("Number of transactions with provider org provider-activity-id", "B37", "iati-activity/transaction[provider-org[string-length(@provider-activity-id) > 0]]", "len"),
        ("Number of transactions that include receiver org", "B39", "iati-activity/transaction[receiver-org[(string-length(@ref) > 0) or (narrative[string-length(text()) > 0])]]", "len"),
        ("Percentage of transactions that include receiver org", "B40", "round((indicator_values['Number of transactions that include receiver org'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions with receiver org only ref", "B41", "iati-activity/transaction[receiver-org[(string-length(@ref) > 0) and not(narrative[string-length(text()) > 0])]]", "len"),
        ("Number of transactions with receiver org only narrative", "B42", "iati-activity/transaction[receiver-org[not(string-length(@ref) > 0) and (narrative[string-length(text()) > 0])]]", "len"),
        ("Number of transactions with receiver org ref and narrative", "B43", "iati-activity/transaction[receiver-org[(string-length(@ref) > 0) and (narrative[string-length(text()) > 0])]]", "len"),
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
        ("Budgets with start dates beyond 2024", "B104", "iati-activity/budget[period-start[number(translate(@iso-date, '-', '')) > 20231231]]", "len"),
        ("Number of activities with only original budgets", "B105", "iati-activity[budget[@type='1'] and not(budget[@type='2'])]", "len"),
        ("Number of activities with only revised budgets", "B106", "iati-activity[budget[@type='2'] and not(budget[@type='1'])]", "len"),
        ("Number of activities with original and revised budgets", "B107", "iati-activity[budget[@type='1'] and budget[@type='2']]", "len"),

        ("Number of planned-disbursements", "B110", "iati-activity/planned-disbursement", "len"),
        ("Number of planned-disbursements in 2018", "C110", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2018')]]", "len"),
        ("Number of planned-disbursements in 2019", "D110", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2019')]]", "len"),
        ("Number of planned-disbursements in 2020", "E110", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2020')]]", "len"),
        ("Number of planned-disbursements in 2021", "F110", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2021')]]", "len"),
        ("Number of planned-disbursements in 2022", "G110", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2022')]]", "len"),
        ("Number of planned-disbursements in 2023", "H110", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2023')]]", "len"),
        ("Number of activities that contain planned-disbursements", "B111", "iati-activity[planned-disbursement]", "len"),
        ("Percentage of activities that contain planned-disbursements", "B112", "round((indicator_values['Number of activities that contain planned-disbursements'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Planned-disbursements with start dates beyond 2024", "B113", "iati-activity/planned-disbursement[period-start[number(translate(@iso-date, '-', '')) > 20231231]]", "len"),

        ("Number of activities with recipient country or recipient region", "B120", "iati-activity[recipient-country|recipient-region]", "len"),
        ("Number of activities with recipient country", "B121", "iati-activity[recipient-country]", "len"),
        ("Number of activities with recipient region", "B122", "iati-activity[recipient-region]", "len"),
        ("Number of activities with recipient country and recipient region", "B123", "iati-activity[recipient-country and recipient-region]", "len"),
        ("Number of transactions in activities with recipient country or recipient region", "B124", "iati-activity[recipient-country|recipient-region]/transaction", "len"),
        ("Number of transactions with recipient country", "B125", "iati-activity/transaction[recipient-country]", "len"),
        ("Number of transactions with recipient region", "B126", "iati-activity/transaction[recipient-region]", "len"),
        ("Number of activities with transactions that have recipient country and recipient region", "B127", "iati-activity[transaction[recipient-country|recipient-region]]", "len"),

        ("Number of activities that include administrative or point", "B130", "iati-activity[location[administrative|point[pos]]]", "len"),
        ("Percentage of activities that include administrative or point", "B131", "round((indicator_values['Number of activities that include administrative or point'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities that include administrative", "B132", "iati-activity[location[administrative]]", "len"),
        ("Which administrative vocabularies are used", "B133", "iati-activity/location/administrative/@vocabulary", "unique"),
        ("Number of activities that include point", "B134", "iati-activity[location[point[pos]]]", "len"),

        ("Numbers of activities with sector", "B137", "iati-activity[sector]", "len"),
        ("Which vocabularies are used in activity sector", "B138", "iati-activity/sector/@vocabulary", "unique"),
        ("Number of activities including sector vocabulary 1", "B139", "iati-activity[sector[@vocabulary='1' or not(@vocabulary)]]", "len"),
        ("Percentage of activities including sector vocabulary 1", "B140", "round((indicator_values['Number of activities including sector vocabulary 1'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities including sector vocabulary 2", "B141", "iati-activity[sector[@vocabulary='2']]", "len"),
        ("Percentage of activities including sector vocabulary 2", "B142", "round((indicator_values['Number of activities including sector vocabulary 2'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities including sector vocabulary 1 and sector vocabulary 2", "B143", "iati-activity[sector[@vocabulary='1'] and sector[@vocabulary='2']]", "len"),
        ("Number of transactions with sector", "B144", "iati-activity/transaction[sector]", "len"),
        ("Which vocabularies are used in transaction sector", "B145", "iati-activity/transaction/sector/@vocabulary", "unique"),
        ("Number of transactions including sector vocabulary 1", "B146", "iati-activity/transaction[sector[@vocabulary='1' or not(@vocabulary)]]", "len"),
        ("Percentage of transactions including sector vocabulary 1", "B147", "round((indicator_values['Number of transactions including sector vocabulary 1'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions including sector vocabulary 2", "B148", "iati-activity/transaction[sector[@vocabulary='2']]", "len"),
        ("Percentage of transactions including sector vocabulary 2", "B149", "round((indicator_values['Number of transactions including sector vocabulary 2'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions including sector vocabulary 1 and sector vocabulary 2", "B150", "iati-activity/transaction[sector[@vocabulary='1'] and sector[@vocabulary='2']]", "len"),

        ("Number of activities with any SDG information", "B153", "iati-activity[sector[@vocabulary='7' or @vocabulary='8' or @vocabulary='9'] | transaction[sector[@vocabulary='7' or @vocabulary='8' or @vocabulary='9']] | tag[@vocabulary='2' or @vocabulary='3'] | result[indicator[reference[@vocabulary='9']]]]", "len"),
        ("Number of activities with SDG tag", "B154", "iati-activity[tag[@vocabulary='2' or @vocabulary='3']]", "len"),
        ("Number of activities with SDG sector", "B155", "iati-activity[sector[@vocabulary='7' or @vocabulary='8' or @vocabulary='9'] | transaction[sector[@vocabulary='7' or @vocabulary='8' or @vocabulary='9']]]", "len"),
        ("Number of activities with SDG result indicator", "B156", "iati-activity[result[indicator[reference[@vocabulary='9']]]]", "len"),
        ("Number of activities with policy marker", "B157", "iati-activity[policy-marker]", "len"),
        ("Which policy marker codes are being used", "B158", "iati-activity/policy-marker/@code", "unique"),

        ("Number of activities with humanitarian flag", "B162", "iati-activity[@humanitarian='1' or @humanitarian='true']", "len"),
        ("Percentage of activities with humanitarian flag", "B163", "round((indicator_values['Number of activities with humanitarian flag'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of transactions with humanitarian flag", "B164", "iati-activity/transaction[@humanitarian='1' or @humanitarian='true']", "len"),
        ("Number of activities with humanitarian scope", "B165", "iati-activity[humanitarian-scope]", "len"),
        ("Which humanitarian-scope types are being used", "B166", "iati-activity/humanitarian-scope/@type", "unique"),
        ("Which humanitarian-scope vocabularies are being used", "B167", "iati-activity/humanitarian-scope/@vocabulary", "unique"),

        ("Number of activities including default finance type", "B171", "iati-activity[default-finance-type]", "len"),
        ("Percentage of activities including default finance type", "B172", "round((indicator_values['Number of activities including default finance type'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of transactions including finance type", "B173", "iati-activity/transaction[finance-type]", "len"),
        ("Percentage of transactions including finance type", "B174", "round((indicator_values['Number of transactions including finance type'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Which transaction types have finance-type", "B175", "iati-activity/transaction[finance-type]/transaction-type/@code", "unique"),
        ("Which default finance type is being used", "B176", "iati-activity/default-finance-type/@code", "unique"),
        ("Number of activities including default flow type", "B178", "iati-activity[default-flow-type]", "len"),
        ("Percentage of activities including default flow type", "B179", "round((indicator_values['Number of activities including default flow type'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of transactions including flow type", "B180", "iati-activity/transaction[flow-type]", "len"),
        ("Percentage of transactions including flow type", "B181", "round((indicator_values['Number of transactions including flow type'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Which transaction types have flow-type", "B182", "iati-activity/transaction[flow-type]/transaction-type/@code", "unique"),
        ("Which default flow type is being used", "B183", "iati-activity/default-flow-type/@code", "unique"),

        ("Number of activities including default aid type", "B187", "iati-activity[default-aid-type]", "len"),
        ("Percentage of activities including default aid type", "B188", "round((indicator_values['Number of activities including default aid type'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of transactions including aid type", "B189", "iati-activity/transaction[aid-type]", "len"),
        ("Percentage of transactions including aid type", "B190", "round((indicator_values['Number of transactions including aid type'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Which transaction types have aid-type", "B191", "iati-activity/transaction[aid-type]/transaction-type/@code", "unique"),
        ("Which default aid type vocabularies are being used", "B192", "iati-activity/default-aid-type/@vocabulary", "unique"),
        ("Which aid type vocabularies are being used", "B193", "iati-activity/transaction/aid-type/@vocabulary", "unique"),
    ]

    indicator_values = dict()

    xml_path = os.path.join("/home/alex/git/IATI-Registry-Refresher/data", args.publisher, '*')
    xml_files = glob.glob(xml_path)
    bar = progressbar.ProgressBar()
    for xml_file in bar(xml_files):
        try:
            tree = etree.parse(xml_file, parser=large_parser)
        except etree.XMLSyntaxError:
            continue
        root = tree.getroot()
        
        for indicator_name, indicator_location, indicator_xpath, indicator_function in indicators:
            if indicator_function == "len":
                evaluated_value = len(root.xpath(indicator_xpath))
            elif indicator_function == "unique":
                evaluated_value = list(set(root.xpath(indicator_xpath)))
            else:
                evaluated_value = []
            if indicator_name not in indicator_values.keys():
                indicator_values[indicator_name] = evaluated_value
            else:
                indicator_values[indicator_name] += evaluated_value
        destroy_tree(tree)

    if len(indicator_values.keys()) > 0:
        for indicator_name, indicator_location, indicator_xpath, indicator_function in indicators:
            if indicator_function == "eval":
                try:
                    indicator_values[indicator_name] = eval(indicator_xpath)
                except ZeroDivisionError:
                    indicator_values[indicator_name] = 0
            elif indicator_function == "unique":
                indicator_values[indicator_name] = list(set(indicator_values[indicator_name]))
            accumulated_value = indicator_values[indicator_name]
            if type(accumulated_value) is list:
                accumulated_value = ", ".join(accumulated_value)
            sheet[indicator_location] = accumulated_value

    outfile = os.path.join(output_dir, "publisher.xlsx")
    wb.save(outfile)
        