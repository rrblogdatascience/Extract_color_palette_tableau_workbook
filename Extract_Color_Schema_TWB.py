'''
Created on 15.06.2014

Extract customized color palettes from workbook files (TWB)
'''

from xml.etree import ElementTree
import argparse
import os

# parse command line arguments
parser = argparse.ArgumentParser()
parser.add_argument("--s", help="save as tps", action="store_true")
parser.add_argument("F", help="filename of twb (including dir")
parser.add_argument("VarNames", help="Names of the attributes (as they show up in Tableau)", nargs='+')
args = parser.parse_args()

# store variables
dimensionNames = args.VarNames
twbFileName = args.F
writeToFile = args.s

# open workbook
with open(twbFileName, 'rt') as f:
    tree = ElementTree.parse(f)
parent_map = {c:p for p in tree.iter() for c in p} # create map for parent-child relation

# start parsing the workbook
xmlColorPaletteCode = "<?xml version='1.0'?>\n\n<workbook>\n    <preferences>"
for node in tree.findall('./datasources/datasource/column'):
    name = node.attrib.get('name')
    caption = node.attrib.get('caption')
    if (caption is None and name.replace("[", "").replace("]", "") in dimensionNames) or caption in dimensionNames:
        parentElement = parent_map[node]
        print "++ found attribute with name = ", name, " and caption = ", caption, " in datasource ", parentElement.attrib.get('caption')
        for nodeCol in parentElement.findall('./column-instance'):
            if nodeCol.attrib.get('column') == name:
                styleName = nodeCol.attrib.get('name')
                parentElement = parent_map[nodeCol]
                for nodeStyle in parentElement.findall('./style/style-rule/encoding'):
                    if nodeStyle.attrib.get('field') == styleName and nodeStyle.attrib.get('type')  == "palette":
                        xmlColorPaletteCode = xmlColorPaletteCode + "\n        <color-palette name='"+ parentElement.attrib.get('caption') + "|" + name.replace("[", "").replace("]", "") + "' type='regular'>\n"
                        for nodeMap in nodeStyle.iter():
                            if nodeMap.tag == "map":
                                xmlColorPaletteCode = xmlColorPaletteCode + "            <color>" + nodeMap.attrib.get("to") + "</color>\n"
                        xmlColorPaletteCode = xmlColorPaletteCode + "        </color-palette>"

xmlColorPaletteCode = xmlColorPaletteCode + "\n    </preferences>\n</workbook>"

print "\n++ xml code for Preferences.tps:\n"
print xmlColorPaletteCode

if writeToFile:
    print "\n++ save "+os.getcwd()+"\\Preferences.tps"
    text_file = open("Preferences.tps", "w")
    text_file.write(xmlColorPaletteCode)
    text_file.close()
