'''Run this script with XML and excel files
in the same directory rename the xml files that
needs to be parsed as 1.xml,2.xml,3.xml,etc'''

#importing Lib/xml/etree/ElementTree.py for XML parsing

import xml.etree.ElementTree as ET #For parsing XML files
import shutil as st #For file copy
import xlrd # Reading an excel file using Python


#User engagement block
print()
print("  ******* Hope U have read the Readme.md before executing this Python Script *******  ")
print()
confirmation=input("  ******* Press ENTER to Proceed If you have renamed the XML files according to the Readme.md  *******  ")
print()
#end of engagement block

child_tags_to_find=() # Identity to be found

#guardian pattern to avoid any exceptions block

try :
    numberOfFiles=int(input("  ******* Enter the total number of XML files to be parsed for eg :1,2,etc *******  "))
    print()
    tags_to_change=input("  ******* Enter the tags to be identified with comma seperator for eg: title,label *******  ").split(',')
    print()
    child_tags_to_find=tuple(tags_to_change)

except ValueError:
    print("  ******* Oops...! Enter a valid number ********  ")
    quit()

#end of guardian pattern



''' This function is used to get the number of identities and xml file object as
input and then find out the identities value to create the new XML file'''

def translateFileContent(lc_tree,lc_child_tags_to_find,lc_new_file):

#empty list for maintaing tags
    tags={}

    root = lc_tree.getroot() #creating root ElementTree

    for childern in lc_child_tags_to_find: #definite loop for iteration over identity
        if childern not in tags:
            childern=str(childern)
            list_of_tags=[]
            for child in root.findall('.//'+childern):
                list_of_tags.append(child.text)
            tags[childern]=list_of_tags
    return tags

#end of main function


''' This function is used to get the open the Xls file and the Fetch the value'''
def openXlsFileAndFetchValue(lc_var,lc_new_tree):

    for key,val in lc_var.items():
        lc_root = lc_new_tree.getroot()
        for child in lc_root.findall('.//'+key):
            child_Val=child.text
            loc = ("JP lookup.xlsx")
            # To open Workbook
            wb = xlrd.open_workbook(loc)
            sheet = wb.sheet_by_index(0)
            for i in range(sheet.nrows):
                identity=sheet.cell_value(i, 0)
                if child_Val == identity:
                    japenese_value=sheet.cell_value(i, 1)
                    child.text=japenese_value
    return lc_new_tree
#end of opening the xml and getting japanese values function

#iteration over number of xml files to find the identites like label and title
try:
    while numberOfFiles>0:
        for count in range(1,numberOfFiles+1):
            file_name=str(count)+".xml"
            tree = ET.parse(file_name)
            new_file=str(count)+"_JP"+".xml"
            st.copy(file_name,new_file)
            new_tree=ET.parse(new_file)
            tags_list=translateFileContent(tree,child_tags_to_find,file_name)
            new_tree_obj=openXlsFileAndFetchValue(tags_list,new_tree)
            new_tree_obj.write(open(new_file, 'wb'),encoding='UTF-8')
            numberOfFiles-=1
    print("  ******* Successfully Translated ********  ")
    print()
except Exception:
    print("  ******* Some thing went wrong. Check with the Documentation *******  ")
    quit()
