import json
from lxml import etree
from lxml.etree import XMLParser
import os


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

large_parser = XMLParser(huge_tree=True)
parser = etree.XMLParser(remove_blank_text=True)

if __name__ == "__main__":

    output = dict()

    with open("iati_schema_xpaths.txt") as xpath_txt_file:
        xpaths = xpath_txt_file.read().splitlines()

    xml_file = "input.xml"
    try:
        tree = etree.parse(xml_file, parser=large_parser)
    except etree.XMLSyntaxError:
        pass
    root = tree.getroot()
    
    for xpath in xpaths:
        output[xpath] = root.xpath(xpath)

    # destroy_tree(tree)

    with open("schema_test_output.json", "w") as json_file:
        json_file.write(json.dumps(output, indent=4))

        