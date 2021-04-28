import argparse
import glob
import json
from lxml import etree
from lxml.etree import XMLParser
from openpyxl import load_workbook
import os
import progressbar
import requests
import csv


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

publisher_list = dict()
with open("iati_publishers_list.csv", "r") as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['IATI Organisation Identifier'] and row['id']:
            publisher_list[row['id']] = row['IATI Organisation Identifier']

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
    danger_errors = pub_validation[0]["summaryStats"]["danger"]
    
    sheet['B8'] = critical_errors
    sheet['B9'] = danger_errors

    if args.publisher in publisher_list.keys():
        org_ref = publisher_list[args.publisher]
        datastore_fail_url = "https://iatidatastore.iatistandard.org/api/datasets/fails/?format=json&publisher_identifier={}".format(org_ref)
        datastore_fail_validation = json.loads(requests.get(datastore_fail_url).content)

        ds_critical_fails = 0
        for result in datastore_fail_validation["results"]:
            ds_critical_fails += result["validation_status"]["critical"]
        sheet['B10'] = ds_critical_fails

        datastore_pickup_url = "https://iatidatastore.iatistandard.org/api/datasets/failedpickups/?format=json&publisher_identifier={}".format(org_ref)
        datastore_pickup = json.loads(requests.get(datastore_pickup_url).content)
        ds_inaccessible = datastore_pickup["count"]
        sheet['B11'] = ds_inaccessible

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
        ("Activity status being used", "B12", "iati-activity/activity-status/@code", "unique"),
        ("Currencies being used", "B13", "(iati-activity/@default-currency)|(iati-activity/budget/value/@currency)|(iati-activity/planned-disbursement/value/@currency)|(iati-activity/transaction/value/@currency)", "unique"),

        ("Participating orgs with narrative and ref", "B15", "iati-activity/participating-org[(string-length(@ref) > 0) and (string-length(narrative/text()) > 0)]", "len"),
        ("Participating orgs with ref", "B16", "iati-activity/participating-org[string-length(@ref) > 0]", "len"),
        ("Participating orgs with narrative", "B17", "iati-activity/participating-org[string-length(narrative/text()) > 0]", "len"),
        ("Participating orgs with neither ref nor narrative", "B18", "iati-activity/participating-org[(not(@ref) or string-length(@ref) = 0) and (not(narrative) or string-length(narrative/text()) = 0)]", "len"),
        ("Number of activities with funding org", "B20", "iati-activity[participating-org[@role='1']]", "len"),
        ("Percentage of activities with funding org", "B21", "round((indicator_values['Number of activities with funding org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities with accountable org", "B23", "iati-activity[participating-org[@role='2']]", "len"),
        ("Percentage of activities with accountable org", "B24", "round((indicator_values['Number of activities with accountable org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities with extending org", "B26", "iati-activity[participating-org[@role='3']]", "len"),
        ("Percentage of activities with extending org", "B27", "round((indicator_values['Number of activities with extending org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities with implementing org", "B29", "iati-activity[participating-org[@role='4']]", "len"),
        ("Percentage of activities with implementing org", "B30", "round((indicator_values['Number of activities with implementing org'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions that include provider org", "B33", "iati-activity/transaction[provider-org[(string-length(@ref) > 0) or (string-length(narrative/text()) > 0)]]", "len"),
        ("Percentage of transactions that include provider org", "B34", "round((indicator_values['Number of transactions that include provider org'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions with provider org only ref", "B35", "iati-activity/transaction[provider-org[(string-length(@ref) > 0) and not(string-length(narrative/text()) > 0)]]", "len"),
        ("Number of transactions with provider org only narrative", "B36", "iati-activity/transaction[provider-org[not(string-length(@ref) > 0) and (string-length(narrative/text()) > 0)]]", "len"),
        ("Number of transactions with provider org ref and narrative", "B37", "iati-activity/transaction[provider-org[(string-length(@ref) > 0) and (string-length(narrative/text()) > 0)]]", "len"),
        ("Number of transactions with provider org provider-activity-id", "B38", "iati-activity/transaction[provider-org[string-length(@provider-activity-id) > 0]]", "len"),
        ("Number of blank provider orgs", "B39", "iati-activity/transaction/provider-org[(not(@ref) or string-length(@ref) = 0) and (not(narrative) or string-length(narrative/text()) = 0)]", "len"),
        ("List of transaction types provider orgs have been added for", "B40", "iati-activity/transaction[provider-org[(string-length(@ref) > 0) or (string-length(narrative/text()) > 0)]]/transaction-type/@code", "unique"),

        ("Number of transactions that include receiver org", "B42", "iati-activity/transaction[receiver-org[(string-length(@ref) > 0) or (string-length(narrative/text()) > 0)]]", "len"),
        ("Percentage of transactions that include receiver org", "B43", "round((indicator_values['Number of transactions that include receiver org'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions with receiver org only ref", "B44", "iati-activity/transaction[receiver-org[(string-length(@ref) > 0) and not(string-length(narrative/text()) > 0)]]", "len"),
        ("Number of transactions with receiver org only narrative", "B45", "iati-activity/transaction[receiver-org[not(string-length(@ref) > 0) and (string-length(narrative/text()) > 0)]]", "len"),
        ("Number of transactions with receiver org ref and narrative", "B46", "iati-activity/transaction[receiver-org[(string-length(@ref) > 0) and (string-length(narrative/text()) > 0)]]", "len"),
        ("Number of transactions with receiver org receiver-activity-id", "B47", "iati-activity/transaction[receiver-org[@receiver-activity-id]]", "len"),
        ("Number of blank receiver orgs", "B48", "iati-activity/transaction/receiver-org[(not(@ref) or string-length(@ref) = 0) and (not(narrative) or string-length(narrative/text()) = 0)]", "len"),
        ("List of transaction types receiver orgs have been added for", "B49", "iati-activity/transaction[receiver-org[(string-length(@ref) > 0) or (string-length(narrative/text()) > 0)]]/transaction-type/@code", "unique"),

        ("Number of transactions with type 1", "B52", "iati-activity/transaction[transaction-type[@code='1']]", "len"),
        ("Number of activities with transactions with type 1", "B53", "iati-activity[transaction[transaction-type[@code='1']]]", "len"),
        ("Percentage of activities with transactions with type 1", "B54", "round((indicator_values['Number of activities with transactions with type 1'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 2", "B56", "iati-activity/transaction[transaction-type[@code='2']]", "len"),
        ("Number of activities with transactions with type 2", "B57", "iati-activity[transaction[transaction-type[@code='2']]]", "len"),
        ("Percentage of activities with transactions with type 2", "B58", "round((indicator_values['Number of activities with transactions with type 2'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 3", "B60", "iati-activity/transaction[transaction-type[@code='3']]", "len"),
        ("Number of activities with transactions with type 3", "B61", "iati-activity[transaction[transaction-type[@code='3']]]", "len"),
        ("Percentage of activities with transactions with type 3", "B62", "round((indicator_values['Number of activities with transactions with type 3'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 4", "B64", "iati-activity/transaction[transaction-type[@code='4']]", "len"),
        ("Number of activities with transactions with type 4", "B65", "iati-activity[transaction[transaction-type[@code='4']]]", "len"),
        ("Percentage of activities with transactions with type 4", "B66", "round((indicator_values['Number of activities with transactions with type 4'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 5", "B68", "iati-activity/transaction[transaction-type[@code='5']]", "len"),
        ("Number of activities with transactions with type 5", "B69", "iati-activity[transaction[transaction-type[@code='5']]]", "len"),
        ("Percentage of activities with transactions with type 5", "B70", "round((indicator_values['Number of activities with transactions with type 5'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 6", "B72", "iati-activity/transaction[transaction-type[@code='6']]", "len"),
        ("Number of activities with transactions with type 6", "B73", "iati-activity[transaction[transaction-type[@code='6']]]", "len"),
        ("Percentage of activities with transactions with type 6", "B74", "round((indicator_values['Number of activities with transactions with type 6'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 7", "B76", "iati-activity/transaction[transaction-type[@code='7']]", "len"),
        ("Number of activities with transactions with type 7", "B77", "iati-activity[transaction[transaction-type[@code='7']]]", "len"),
        ("Percentage of activities with transactions with type 7", "B78", "round((indicator_values['Number of activities with transactions with type 7'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 8", "B80", "iati-activity/transaction[transaction-type[@code='8']]", "len"),
        ("Number of activities with transactions with type 8", "B81", "iati-activity[transaction[transaction-type[@code='8']]]", "len"),
        ("Percentage of activities with transactions with type 8", "B82", "round((indicator_values['Number of activities with transactions with type 8'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 9", "B84", "iati-activity/transaction[transaction-type[@code='9']]", "len"),
        ("Number of activities with transactions with type 9", "B85", "iati-activity[transaction[transaction-type[@code='9']]]", "len"),
        ("Percentage of activities with transactions with type 9", "B86", "round((indicator_values['Number of activities with transactions with type 9'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 10", "B88", "iati-activity/transaction[transaction-type[@code='10']]", "len"),
        ("Number of activities with transactions with type 10", "B89", "iati-activity[transaction[transaction-type[@code='10']]]", "len"),
        ("Percentage of activities with transactions with type 10", "B90", "round((indicator_values['Number of activities with transactions with type 10'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 11", "B92", "iati-activity/transaction[transaction-type[@code='11']]", "len"),
        ("Number of activities with transactions with type 11", "B93", "iati-activity[transaction[transaction-type[@code='11']]]", "len"),
        ("Percentage of activities with transactions with type 11", "B94", "round((indicator_values['Number of activities with transactions with type 11'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 12", "B96", "iati-activity/transaction[transaction-type[@code='12']]", "len"),
        ("Number of activities with transactions with type 12", "B97", "iati-activity[transaction[transaction-type[@code='12']]]", "len"),
        ("Percentage of activities with transactions with type 12", "B98", "round((indicator_values['Number of activities with transactions with type 12'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of transactions with type 13", "B100", "iati-activity/transaction[transaction-type[@code='13']]", "len"),
        ("Number of activities with transactions with type 13", "B101", "iati-activity[transaction[transaction-type[@code='13']]]", "len"),
        ("Percentage of activities with transactions with type 13", "B102", "round((indicator_values['Number of activities with transactions with type 13'] / indicator_values['Number of activities']) * 100, 2)", "eval"),

        ("Number of budgets", "B105", "iati-activity/budget", "len"),
        ("Number of budgets in 2018", "C105", "iati-activity/budget[period-start[starts-with(@iso-date, '2018')]]", "len"),
        ("Number of budgets in 2019", "D105", "iati-activity/budget[period-start[starts-with(@iso-date, '2019')]]", "len"),
        ("Number of budgets in 2020", "E105", "iati-activity/budget[period-start[starts-with(@iso-date, '2020')]]", "len"),
        ("Number of budgets in 2021", "F105", "iati-activity/budget[period-start[starts-with(@iso-date, '2021')]]", "len"),
        ("Number of budgets in 2022", "G105", "iati-activity/budget[period-start[starts-with(@iso-date, '2022')]]", "len"),
        ("Number of budgets in 2023", "H105", "iati-activity/budget[period-start[starts-with(@iso-date, '2023')]]", "len"),
        ("Number of activities that contain budgets", "B106", "iati-activity[budget]", "len"),
        ("Percentage of activities that contain budgets", "B107", "round((indicator_values['Number of activities that contain budgets'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Budgets with start dates beyond 2024", "B108", "iati-activity/budget[period-start[number(translate(@iso-date, '-', '')) > 20231231]]", "len"),
        ("Number of activities with only original budgets", "B109", "iati-activity[budget[@type='1'] and not(budget[@type='2'])]", "len"),
        ("Number of activities with only revised budgets", "B110", "iati-activity[budget[@type='2'] and not(budget[@type='1'])]", "len"),
        ("Number of activities with original and revised budgets", "B111", "iati-activity[budget[@type='1'] and budget[@type='2']]", "len"),

        ("Number of planned-disbursements", "B114", "iati-activity/planned-disbursement", "len"),
        ("Number of planned-disbursements in 2018", "C114", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2018')]]", "len"),
        ("Number of planned-disbursements in 2019", "D114", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2019')]]", "len"),
        ("Number of planned-disbursements in 2020", "E114", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2020')]]", "len"),
        ("Number of planned-disbursements in 2021", "F114", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2021')]]", "len"),
        ("Number of planned-disbursements in 2022", "G114", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2022')]]", "len"),
        ("Number of planned-disbursements in 2023", "H114", "iati-activity/planned-disbursement[period-start[starts-with(@iso-date, '2023')]]", "len"),
        ("Number of activities that contain planned-disbursements", "B115", "iati-activity[planned-disbursement]", "len"),
        ("Percentage of activities that contain planned-disbursements", "B116", "round((indicator_values['Number of activities that contain planned-disbursements'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Planned-disbursements with start dates beyond 2024", "B117", "iati-activity/planned-disbursement[period-start[number(translate(@iso-date, '-', '')) > 20231231]]", "len"),

        ("Number of activities with recipient country or recipient region", "B124", "iati-activity[recipient-country|recipient-region]", "len"),
        ("Number of activities with recipient country", "B125", "iati-activity[recipient-country]", "len"),
        ("Number of activities with recipient region", "B126", "iati-activity[recipient-region]", "len"),
        ("Number of activities with recipient country and recipient region", "B127", "iati-activity[recipient-country and recipient-region]", "len"),
        ("Number of transactions in activities with recipient country or recipient region", "B128", "iati-activity[recipient-country|recipient-region]/transaction", "len"),
        ("Number of transactions with recipient country", "B129", "iati-activity/transaction[recipient-country]", "len"),
        ("Number of transactions with recipient region", "B130", "iati-activity/transaction[recipient-region]", "len"),
        ("Number of activities with transactions that have recipient country and recipient region", "B131", "iati-activity[transaction[recipient-country|recipient-region]]", "len"),

        ("Number of activities that include administrative or point", "B134", "iati-activity[location[administrative|point[pos]]]", "len"),
        ("Percentage of activities that include administrative or point", "B135", "round((indicator_values['Number of activities that include administrative or point'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities that include administrative", "B136", "iati-activity[location[administrative]]", "len"),
        ("Which administrative vocabularies are used", "B137", "iati-activity/location/administrative/@vocabulary", "unique"),
        ("Number of activities that include point", "B138", "iati-activity[location[point[pos]]]", "len"),

        ("Numbers of activities with sector", "B141", "iati-activity[sector]", "len"),
        ("Which vocabularies are used in activity sector", "B142", "iati-activity/sector/@vocabulary", "unique"),
        ("Number of activities including sector vocabulary 1", "B143", "iati-activity[sector[@vocabulary='1' or not(@vocabulary)]]", "len"),
        ("Percentage of activities including sector vocabulary 1", "B144", "round((indicator_values['Number of activities including sector vocabulary 1'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities including sector vocabulary 2", "B145", "iati-activity[sector[@vocabulary='2']]", "len"),
        ("Percentage of activities including sector vocabulary 2", "B146", "round((indicator_values['Number of activities including sector vocabulary 2'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of activities including sector vocabulary 1 and sector vocabulary 2", "B147", "iati-activity[sector[@vocabulary='1'] and sector[@vocabulary='2']]", "len"),
        ("Number of transactions with sector", "B148", "iati-activity/transaction[sector]", "len"),
        ("Which vocabularies are used in transaction sector", "B149", "iati-activity/transaction/sector/@vocabulary", "unique"),
        ("Number of transactions including sector vocabulary 1", "B150", "iati-activity/transaction[sector[@vocabulary='1' or not(@vocabulary)]]", "len"),
        ("Percentage of transactions including sector vocabulary 1", "B151", "round((indicator_values['Number of transactions including sector vocabulary 1'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions including sector vocabulary 2", "B152", "iati-activity/transaction[sector[@vocabulary='2']]", "len"),
        ("Percentage of transactions including sector vocabulary 2", "B153", "round((indicator_values['Number of transactions including sector vocabulary 2'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Number of transactions including sector vocabulary 1 and sector vocabulary 2", "B154", "iati-activity/transaction[sector[@vocabulary='1'] and sector[@vocabulary='2']]", "len"),

        ("Number of activities with any SDG information", "B157", "iati-activity[sector[@vocabulary='7' or @vocabulary='8' or @vocabulary='9'] | transaction[sector[@vocabulary='7' or @vocabulary='8' or @vocabulary='9']] | tag[@vocabulary='2' or @vocabulary='3'] | result[indicator[reference[@vocabulary='9']]]]", "len"),
        ("Number of activities with SDG tag", "B158", "iati-activity[tag[@vocabulary='2' or @vocabulary='3']]", "len"),
        ("Number of activities with SDG sector", "B159", "iati-activity[sector[@vocabulary='7' or @vocabulary='8' or @vocabulary='9'] | transaction[sector[@vocabulary='7' or @vocabulary='8' or @vocabulary='9']]]", "len"),
        ("Number of activities with SDG result indicator", "B160", "iati-activity[result[indicator[reference[@vocabulary='9']]]]", "len"),
        ("Number of activities with policy marker", "B161", "iati-activity[policy-marker]", "len"),
        ("Which policy marker codes are being used", "B162", "iati-activity/policy-marker/@code", "unique"),

        ("Number of activities with humanitarian flag", "B166", "iati-activity[@humanitarian='1' or @humanitarian='true']", "len"),
        ("Percentage of activities with humanitarian flag", "B167", "round((indicator_values['Number of activities with humanitarian flag'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of transactions with humanitarian flag", "B168", "iati-activity/transaction[@humanitarian='1' or @humanitarian='true']", "len"),
        ("Number of activities with humanitarian scope", "B169", "iati-activity[humanitarian-scope]", "len"),
        ("Which humanitarian-scope types are being used", "B170", "iati-activity/humanitarian-scope/@type", "unique"),
        ("Which humanitarian-scope vocabularies are being used", "B171", "iati-activity/humanitarian-scope/@vocabulary", "unique"),

        ("Number of activities including default finance type", "B175", "iati-activity[default-finance-type]", "len"),
        ("Percentage of activities including default finance type", "B176", "round((indicator_values['Number of activities including default finance type'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of transactions including finance type", "B177", "iati-activity/transaction[finance-type]", "len"),
        ("Percentage of transactions including finance type", "B178", "round((indicator_values['Number of transactions including finance type'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Which transaction types have finance-type", "B179", "iati-activity/transaction[finance-type]/transaction-type/@code", "unique"),
        ("Which default finance type is being used", "B180", "iati-activity/default-finance-type/@code", "unique"),
        ("Number of activities including default flow type", "B182", "iati-activity[default-flow-type]", "len"),
        ("Percentage of activities including default flow type", "B183", "round((indicator_values['Number of activities including default flow type'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of transactions including flow type", "B184", "iati-activity/transaction[flow-type]", "len"),
        ("Percentage of transactions including flow type", "B185", "round((indicator_values['Number of transactions including flow type'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Which transaction types have flow-type", "B186", "iati-activity/transaction[flow-type]/transaction-type/@code", "unique"),
        ("Which default flow type is being used", "B187", "iati-activity/default-flow-type/@code", "unique"),

        ("Number of activities including default aid type", "B191", "iati-activity[default-aid-type]", "len"),
        ("Percentage of activities including default aid type", "B192", "round((indicator_values['Number of activities including default aid type'] / indicator_values['Number of activities']) * 100, 2)", "eval"),
        ("Number of transactions including aid type", "B193", "iati-activity/transaction[aid-type]", "len"),
        ("Percentage of transactions including aid type", "B194", "round((indicator_values['Number of transactions including aid type'] / indicator_values['Number of transactions']) * 100, 2)", "eval"),
        ("Which transaction types have aid-type", "B195", "iati-activity/transaction[aid-type]/transaction-type/@code", "unique"),
        ("Which default aid type codes are being used", "B196", "iati-activity/default-aid-type/@code", "unique"),
        ("Which aid type codes are being used", "B197", "iati-activity/transaction/aid-type/@code", "unique"),
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
        