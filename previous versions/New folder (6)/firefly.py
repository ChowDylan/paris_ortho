import xlrd
import json
import os
import re
import datetime
import pyautogui
import time
import PIL
from Tkinter import Tk

from PIL import Image
import cv
from Order import Order

#example of how to call fields
# o = Order()
# o.fields['pt-shoe-size']
o = Order()
o.extras = {}

# Importing Reference Images
#brandon_target = Image.open('brandon_target.bmp')

   # work computer pathing
# add_button = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/add_button.bmp')
# after_print = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/after_print.bmp')
# device_tab = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/device_tab.bmp')
# existing_orders_button = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/existing_orders_button.bmp')
# general_tab = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/general_tab.bmp')
# inside_pt_card = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/inside_pt_card.bmp')
# interrupt = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/interrupt.bmp')
# deadlock = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/live_deadlock.bmp')
# new_order = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/new_order.bmp')
# posted_sales_button = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/posted_button.bmp')
# pt_popup = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/pt_popup.bmp')
# pt_search_entry = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/pt_search_entry.bmp')
# remove_button = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/remove_button.bmp')
# similar_ptname = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/similar_ptname.bmp')
# status_history = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/status_history.bmp')
# sub_search_confirm = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/sub_search_confirm.bmp')
# not_responding = Image.open('C:/Users/dylanchow/PycharmProjects/work2/image_recog/not_responding.bmp')

""" Turn on os rename in file rename (
    replace file paths in file rename (raw_location)
    replace file paths in note pad maker (os path join)
    Need to check if the order of the sub order grabber is correct
    need to redo the image recognition for the suborder grabber
    use test file
    quantity sub 2 plus, keystroke path or image recognition 
    need to remove sameday suborder sub_order changer to new order from changed
    
    
    added screenshots to failure conditions
    added fill difference and left and right device only situation
    left only and right only is easy just marker and check
    left AND right you have to loop again
    + pt card number to flag report?
"""

# time.sleep(0.07)
# print "        	                                             _  _________  _"
# time.sleep(0.07)
# print "   ==   =  = ===== ====                                  \/,  ,,     , \\ |   "
# time.sleep(0.07)
# print "  =  =  =  =   =   =  =                                  | (|      /)   |/   "
# time.sleep(0.07)
# print "  ====  =  =   =   =  =                              -~-{   {""\"\"\"}     '  }-~-   "
# time.sleep(0.07)
# print "  =  =  ====   =   ====                                --{ ._;^:__,'    /--  "
# time.sleep(0.07)
# print "                                                         |              |     "
# time.sleep(0.07)
# print "  ==== ==== ==== ==== ==== =====                          |              \     "
# time.sleep(0.07)
# print "  =    =__  = =  = =  =__    =                           /            ,   \   "
# time.sleep(0.07)
# print "  ==== =    ==   ==   =      =        _   ~~~~~~ ~~~~-- //           ,,    :  "
# time.sleep(0.07)
# print "  =    ==== =  = =  = ====   =    /'                    /                   }  "
# time.sleep(0.07)
# print "                                /'                     {                    } "
# time.sleep(0.07)
# print "                              /''                      {                    }  "
# time.sleep(0.07)
# print "                          _ /          \\\                             '     ;   "
# time.sleep(0.07)
# print "                       /                ,               \             ,,   ;  "
# time.sleep(0.07)
# print "                      /    ___\         }                      {  ''' }    ;  "
# time.sleep(0.07)
# print "                     {    /    \       / __ ________ ~~~ \     |__,,,_|   |  "
# time.sleep(0.07)
# print "                      \   \_____|____/__/____/____        \    |      |   |  "
# time.sleep(0.07)
# print "                       \ _ __ _________________ _  \       |   |      |   |  "
# time.sleep(0.1)
# print "                                \__,,}   \__,,}   \;       {_,,}      {_,,}  "


# os.rename(C:/Users/dylanchow/PycharmProjects/work2/firefly_orders/raw_files)
# exit ()

# a = "0107552"
# a = a.lstrip("0")
# print a

#
# while pyautogui.locateOnScreen(brandon_target) is None:
#     print 'searching 1'
# else:
#     print 'found 1'

# while pyautogui.locateOnScreen(pt_popup, confidence=0.97) is None:
#     print 'searching 2'
# else:
#     print 'found 2'
# exit()
#
# now = datetime.datetime.now()
#
# print now.year
# print "Current date and time using instance attributes:"
# print "Current year: %d" % now.year
# print "Current month: %d" % now.month
# print "Current day: %d" % now.day
# print "Current hour: %d" % now.hour
# print "Current minute: %d" % now.minute
# print "Current second: %d" % now.second
# print "Current microsecond: %d" % now.microsecond
#
# print
# print "Current date and time using strftime:"

# process_start = datetime.datetime.now()
# print process_start
# #print process_start.strftime("%H:%M %m-%d-%Y")
# time.sleep(120)
# process_end = datetime.datetime.now()
# print process_end
# #print process_end.strftime("%H:%M %m-%d-%Y")
#
# elapsedtime = process_end - process_start
# print 'This took', elapsedtime, 'minutes'
# print 'Time Difference:', elapsedtime.strftime("%H:%M")
# print "Current date and time using isoformat:"
# print now.isoformat()

# exit()

# # Break Continue and Pass all demonstrated
# x = 1
# while x < 20:
#     print "loop:", x
#     x += 1
#     if x > 6:
#         continue
#     count = 1
#     while pyautogui.locateOnScreen(brandon_target, region=(0, 0, 400, 800)) is None:
#         print '*Looking for Brandon', count
#
#         time.sleep(2)
#         if count == 3:
#             pass  # more of a placeholder thing
#         print 'Brandon Not Found'
#         count += 1
#     else:
#         print '**Brandon Detected'
#
#     print 'loop complete'
#     time.sleep(1)
# exit()


# # can i do this such that, you can just place things in the brackets in the right order and they will take them
#def imageRecognitionDelay(image2search, left, top, width, height):
#     while pyautogui.locateOnScreen(image2search, region=(left, top, width, height)) is None:
#         time.sleep(1)
# smile = Image.open('smile.bmp')
# brandon_folder = Image.open('brandon.bmp')
# brandon_target = Image.open('brandon_target.bmp')


# pyautogui.doubleClick(pyautogui.locateCenterOnScreen(brandon_folder))
# print 'found'
# time.sleep(2)
# pyautogui.doubleClick(pyautogui.locateCenterOnScreen(brandon_target))
# print 'fin'
# time.sleep(2)
# pyautogui.typewrite(['f11'])
# exit()
# center_list = list(pyautogui.center(pyautogui.locateOnScreen(smile)))
# print center_list
# center_list[0] = center_list[0] - 80
# center = tuple(center_list)
# print center
#
# pyautogui.doubleClick(center)

# this is the device dictionary pertinent to the firefly order form
deviceDictionary = {"FIREFLY NHS FUNCTIONAL": "431750601", "FIREFLY NHS DRESS": "431750611",
                    "FIREFLY NHS SPORT": "431750621","FIREFLY SOCCER SPORT": "431750626",
                    "FIREFLY SOCCER SPORT (DM)": "431750627", "FIREFLY SPORT IMPACT": "431750622",
                    "FUNCTIONAL STANDARD": "43170ST01", "FUNCTIONAL DIRECT MILLED": "43170DM01",
                    "STANDARD SLIMLINE": "43171LA01", "LOW HEEL CUP SLIMLINE": "43171LA11",
                    "FLAT HEEL CUP": "43171LA21", "COBRA": "43171LA31", "MENS DRESS": "43171ME01",
                    "SPORT STANDARD - NEOPRENE TO TOES": "43172ST01",
                    "SPORT DIRECT MILLED - NEOPRENE TO TOES": "43172DM01",
                    "SPORT DIRECT MILLED - VINYL TO METS": "43172DM02", "SPORT LOW PROFILE": "43172LP01",
                    "SPORT SKI - ALPINE": "43172SI01", "SPORT SKI - NORDIC": "43172SI02",
                    "SPORT SKI - SNOWBOARD": "43172SI03", "SPORT SKATE - HOCKEY": "43172SA01",
                    "SPORT SKATE - FIGURE": "43172SA02", "MOLD STANDARD": "43173ST01", "MOLD LOW PROFILE": "43173LP11",
                    "FIREFLY DIABETIC TRIDENSITY": "431750671", "FIREFLY RA FLEXIBLE MOLD": "431750681",
                    "EVA": "43174EV01", "UCBL": "43174UC01", "UCBL CHILDREN": "43174UC02",
                    "ROBERTS WHITMAN": "43174RB01", "ROBERTS WHITMAN CHILDREN": "43174RB02",
                    "GAIT PLATE - INDUCE OUT-TOEING": "43174GP02", "GAIT PLATE - INDUCE IN-TOEING": "43174GP01",
                    "DIABETIC FLEXIBLE MOLD": "43174DB01",
                    "DIABETIC SEMI-FLEXIBLE MOLD": "43174DB04",
                    "RICHIE BRACE STANDARD RESTRICTED HINGE PERMANENTLY FIXED": "43329RB01",
                    "": ""}

#this is the device dictionary peritnent to the AM list of devices
# todo include all devices, including the wrong 'Mens Dre' from firefly
subDeviceDictionary = {"FNHS FUNCTIONAL": "431750601", "FNHS DRESS": "431750611", "FNHS SPORT": "431750621",
                       "FIREFLY SOCCER SPORT": "431750626", "FIREFLY SOCCER SPORT (DM)": "431750627",
                       "FIREFLY SPORT IMPACT": "431750622", "FUNCTIONAL STANDARD": "43170ST01",
                       "FUNCTIONAL DIRECT MILLED": "43170DM01", "LADIES DRESS STANDARD SLIMLINE": "43171LA01",
                       "LADIES DRESS LOW HEEL CUP SLIMLINE": "43171LA11", "LADIES DRESS FLAT HEEL CUP": "43171LA21",
                       "LADIES DRESS COBRA": "43171LA31", "MENS DRESS": "43171ME01",
                       "SPORT STANDARD NEOPRENE TO TOES": "43172ST01",
                       "SPORT DIRECT MILLED NEOPRENE TO TOES": "43172DM01",
                       "SPORT DIRECT MILLED VINYL TO METS": "43172DM02", "SPORT LOW PROFILE": "43172LP01",
                       "SPORT SKI ALPINE": "43172SI01", "SPORT SKI NORDIC": "43172SI02",
                       "SPORT SKI SNOWBOARD": "43172SI03", "SPORT SKATE HOCKEY": "43172SA01",
                       "SPORT SKATE FIGURE": "43172SA02", "MOLD STANDARD": "43173ST01", "MOLD LOW PROFILE": "43173LP11",
                       "FIREFLY DIABETIC TRIDENSITY": "431750671", "FIREFLY RA FLEXIBLE MOLD": "431750681",
                       "EVA": "43174EV01", "UCBL ADULT": "43174UC01", "UCBL CHILDREN": "43174UC02",
                       "ROBERTS WHITMAN ADULT": "43174RB01", "ROBERTS WHITMAN CHILDREN": "43174RB02",
                       "GAIT PLATE INDUCE OUT-TOEING C HILDREN": "43174GP02",
                       "GAIT PLATE INDUCE IN-TOEING CH ILDREN": "43174GP01", "RA SEMI-FLEXIBLE MOLD": "43174RA03",
                       "DIABETIC STANDARD TRIDENSITY": "43174DB02", "DIABETIC CHARCOT TRIDENSITY": "43174DB03",
                       "DIABETIC SEMI-RIGID MOLD": "43174DB06", "RA FLEXIBLE MOLD": "43174RA01",
                       "RA STANDARD TRIDENSITY": "43174RA02", "RA SEMI-FLEXIBLE MOLD LOW PROFILE": "43174RA04",
                       "RA SEMI-RIGID MOLD": "43174RA05", "RA SEMI-RIGID MOLD LOW PROFILE": "43174RA06",
                       "DIABETIC FLEXIBLE MOLD": "43174DB01",
                       "RICHIE BRACE STANDARD RESTRICTED HINGE PERMANENTLY FIXED": "43329RB01",
                       "DIABETIC SEMI-FLEXIBLE MOLD": "43174DB04",
                       "": ""}

# Todo: needs to be changed actual path at work
# output filehandle
cur_time = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M")

out_fh = open(os.path.join("firefly_orders", "order_report_" + str(cur_time) + ".txt"), "w")


# finding all the xls files placed in Orders Folder and making a list of them
targetdir = os.path.join("firefly_orders", "orders")
files = os.listdir(targetdir)
xlsfiles = []

def convertserialtodate(xlserial):
   basedate = datetime.date(1900,1,1)
   delta = datetime.timedelta(days=xlserial)
   newdate = basedate + delta
   newdate.strftime("%m%d%y")
   output = newdate.strftime("%m%d%y")
   return output


def getCleanFiles(rawfiles):
    cleanedfiles = []
    cleanfile_to_rawfile = {}
    # Loop through files and clean names
    for rawfile in rawfiles:
        # cleaning text
        name = str(rawfile).replace(" ", "")
        name = str(name).replace(".", "")
        name = name.upper()
        name = name.replace("RAW", "")
        cleanedfiles.append(name)
        if name not in cleanfile_to_rawfile:
            cleanfile_to_rawfile[name] = rawfile
    return cleanedfiles, cleanfile_to_rawfile


def findDuplicates(rawfiles):
    cleanedfiles, cleanfile_to_rawfile = getCleanFiles(rawfiles)
    # Check for duplicates
    name_counts = {}
    duplicates = []
    for cleanedfile in cleanedfiles:
        if cleanedfile not in name_counts:
            name_counts[cleanedfile] = 1
        else:
            # duplicates
            duplicates.append(cleanedfile)
            name_counts[cleanedfile] += 1

    # Return the duplicate cleannames to their original form
    rawname_duplicates = []
    for duplicate in duplicates:
        if duplicate in cleanfile_to_rawfile:
            # we can
            rawname_duplicates.append(cleanfile_to_rawfile[duplicate])
        else:
            print duplicate, "not in cleanfile_to_rawfile. Something wrong... :{"

    # Print out the dupicate files
    for i in rawname_duplicates:
        print i, 'Is a Duplicate Raw File'

    return rawname_duplicates


# Reading all the raw files placed into the raw_file folder and makes a list
rawdir = os.path.join("firefly_orders", "raw_files")
rawfiles = os.listdir(rawdir)
cleanedfiles, cleanfile_to_rawfile = getCleanFiles(rawfiles)
rawname_duplicates = findDuplicates(rawfiles)


#print convertserialtodate(41875)

#
# def detectHoldRequest(curr_order):
#     curr_order["issue_list"] = []
#     special_instructions = (re.findall("[a-zA-Z'-]+", curr_order["notes"]))
#     for word in special_instructions:
#         if word == 'HOLD':
#             curr_order['issue_list'].append('hold request ')
#         if word == 'ADDRESS':
#             curr_order['issue_list'].append('alternate address ')
#
#     return curr_order
#
# def testDetectHoldRequest():
#     print "Testing: testDetectHoldRequest()"
#     # Should find holds request add to current order
#     curr_order1 = {"notes": "MY HOLD REQUEST"}
#     actual1 = detectHoldRequest(curr_order1)
#     expected1 = {'notes': 'MY HOLD REQUEST', 'issue_list': ['hold request ']}
#     print actual1
#     print actual1 == expected1
#
#     curr_order2 = {"notes": "MY ADDRESS"}
#     actual2 = detectHoldRequest(curr_order2)
#     expected2 = {'notes': 'MY ADDRESS', 'issue_list': ['alternate address ']}
#     print actual2 == expected2
#
#     curr_order3 = {"notes": "MY HOLD ADDRESS"}
#     actual3 = detectHoldRequest(curr_order3)
#     expected3 = {'notes': 'MY HOLD ADDRESS', 'issue_list': ['hold request ', 'alternate address ']}
#     print actual3
#     print actual3 == expected3
#
# # Run tests
# testDetectHoldRequest()
# exit()

def previousOrderSource(raw_prev_po, curr_order):
    """
    Check the previous po number to see which database it came from
    (i.e., either AM or LFE)
    :param row: raw prev_po string from "QUANTITY / SUBSEQUENT PAIR" row
    :return: curr_order updated
    """
    DATABASE_PO_BORDER = 97820
    # Clean raw value
    prev_po = cleanAndReturnNumberString(raw_prev_po)

    prev_po = re.sub("\.0$", "", prev_po)  # removes tailing .0
    curr_order['prev_po'] = prev_po

    if prev_po == "":
        # Not relevant
        curr_order["data_base"] = ""
        return curr_order

    # Determine source
    # TODO Assumed that LFE was on the border
    if int(prev_po) < DATABASE_PO_BORDER:
        curr_order['data_base'] = 'LFE'
    if int(prev_po) > DATABASE_PO_BORDER:
        curr_order['data_base'] = 'AM'

    return curr_order

def testPreviousOrderSource():
    print "testPreviousOrderSource():"
    raw_prev_po = "1000000"
    curr_order = {}
    actual = previousOrderSource(raw_prev_po, curr_order)
    expected = {'data_base': 'AM', 'prev_po': '1000000'}
    print actual
    print expected

    raw_prev_po2 = "96500.0"
    curr_order = {}
    actual = previousOrderSource(raw_prev_po2, curr_order)
    expected = {'data_base': 'LFE', 'prev_po': '96500'}
    print actual
    print expected

    raw_prev_po2 = ""
    curr_order = {}
    actual = previousOrderSource(raw_prev_po2, curr_order)
    expected = {'data_base': '', 'prev_po': ''}
    print actual
    print expected

    raw_prev_po2 = 1000000
    curr_order = {}
    actual = previousOrderSource(raw_prev_po2, curr_order)
    print actual
    print expected

def cleanAndReturnNumberString(raw_num):
    raw_num = str(raw_num).strip() # removes whitespace
    raw_num = re.sub("\.0$", "", raw_num) # removes tailing .0
    return raw_num

def testCleanAndReturnNumberString():
    test1 = "10000.0"
    actual = cleanAndReturnNumberString(test1)
    expected = "10000"
    print actual == expected
    print actual
    print expected


# Run all the tests
# print "Run all tests:"
# testPreviousOrderSource()
# testCleanAndReturnNumberString()

#Time to Fisinish counters
NEWORDER_COUNT = 0
SUBORDER_COUNT = 0
LFE_COUNT = 0


###### PHASE 1: PROCESSING EXCEL ORDERS

for f in files:
    if f.endswith("xls"):
       xlsfiles.append(f)
    if f.endswith("xlsx"):
       xlsfiles.append(f)

def patient_name_ponum(row, curr_order):
    # Creation of the dictionary [why here, why not higher]
    curr_order = {} # Order()
    curr_order['wo_num'] = ''
    curr_order['pt_num'] = ''
    curr_order['po_num'] = ''
    curr_order['suborder_target'] = ''
    curr_order['issue_list'] = []
    # ASK, I would have thought that this would need a raw_num named variable
    curr_order['po_num'] = cleanAndReturnNumberString(row[-1].value)
    curr_order['name'] = row[1].value.strip()

    # finding all character strings, hyphen and apostraphy included
    nameList = (re.findall("[a-zA-Z'-]+", row[1].value.strip()))
    if len(nameList) > 0:
        namemarker = 0
        for name in nameList:
            # print nameList[namemarker]  # extranius word removal from name
            if nameList[namemarker] == 'DR':
                del nameList[namemarker]
                curr_order['issue_list'].append('DR removed')
            if nameList[namemarker] == 'IW':
                del nameList[namemarker]
                curr_order['issue_list'].append('IW removed')
            namemarker += 1

    # finding if there is a pt serial number
    nameNumber = (re.findall("[\d]+", row[1].value.strip()))
    # print nameList
    if curr_order['name'] == '':
        curr_order['issue_list'].append('name missing')
    curr_order['firstname'] = ''
    curr_order['nameLast'] = ''
    curr_order['nameNumber'] = ''
    namespace = ' '
    if len(nameList) == 1:
        namespace = ''

    if len(nameList) >= 1:
        curr_order['firstname'] = " ".join(nameList[0:-1])
        curr_order['nameLast'] = nameList[-1]
    if len(nameList) >= 3:
        # TODO: Logic?
        curr_order['issue_list'].append('Name Syntax Check')
    if len(nameNumber) == 1:
        curr_order['nameNumber'] = nameNumber[0]

    if len(nameNumber) > 1:
        curr_order['nameNumber'] = nameNumber[0]

    curr_order['FIRSTNAMECOMPLETE'] = curr_order['firstname'] + namespace + curr_order['nameNumber']

    curr_order['buddy_order'] = 'no'
    bud_marker = 1
    for i in orders:
        # might need to be refined to compare first name and last name separately
        # TODO: Logic?
        if curr_order['name'] == i['name']:
            curr_order['issue_list'].append('Buddy')
            curr_order['buddy_order'] = 'yes'
            curr_order['suborder_target'] = int(bud_marker)
            # print curr_order['suborder_target'], "poro"

            # print curr_order['suborder_target']
        bud_marker += 1

    return curr_order

def weight_footsize_priority(row, curr_order):
    curr_order['weight'] = ''
    curr_order['shoesize'] = ''
    weightLen = (re.findall('[\d]+', row[1].value))
    if len(weightLen) == 4:
        curr_order['weight'] = weightLen[1]
    if len(weightLen) == 2:
        curr_order['weight'] = weightLen[0]
    # print row[4]
    # if row[4].value == '':
    #     tesla = ''
    # else:
    #     tesla = re.findall('[.\d]+', row[4].value)
    #
    # print tesla
    # curr_order['shoesize'] = tesla
    curr_order['shoesize'] = re.findall('[.\d]+', row[4].value)
    curr_order['priority'] = row[6].value
    if curr_order['weight'] == '':
        curr_order['issue_list'].append('Weight missing')
    if len(curr_order['shoesize']) < 1:
        curr_order['issue_list'].append('Shoesize missing')
        # pyautogui.click(100, 100)
        # pyautogui.typewrite(curr_order['shoesize'])
        # pyautogui.typewrite(['return'])

    return curr_order

def quantity_sub_order_sameday(row, curr_order):
    global LFE_COUNT, SUBORDER_COUNT, NEWORDER_COUNT

    curr_order['quantity_num'] = row[1].value
    if curr_order['quantity_num'] == '':
        curr_order['issue_list'].append('Quantity: dummy used')
        curr_order['quantity_num'] = 1
        curr_order['quantity'] = 1

    curr_order['quantity_num'] = float(curr_order['quantity_num'])
    curr_order['quantity_num'] = int(curr_order['quantity_num'])
    curr_order['quantity'] = str(row[1].value)
    curr_order['quantityquantity'] = curr_order['quantity'].replace(".0", "")

    if curr_order['quantity_num'] > 1:
        curr_order['issue_list'].append('Quantity 2 Plus')
    raw_prev_po = row[4].value
    curr_order['data_base'] = ''
    curr_order = previousOrderSource(raw_prev_po, curr_order)
    if curr_order['data_base'] == 'LFE':
        LFE_COUNT += 1
        curr_order['issue_list'].append('LFE Order')

    sub_order = row[7].value
    # refine condition, they might leave blank by accident, is new if there is foot to scan
    if sub_order == '':
        if curr_order['po_num'] != '':
            NEWORDER_COUNT += 1
        sub_order = 'new order'
    if sub_order == 'CHANGED DEVICE (Select device and options)':
        SUBORDER_COUNT += 1
        sub_order = 'changed'
        if curr_order['prev_po'] == '':
            curr_order['issue_list'].append('missing prev_po')
            curr_order['prev_po'] = 'missing'
    elif sub_order == 'DUPLICATE DEVICE (No change)':
        SUBORDER_COUNT += 1
        sub_order = 'duplicate'
        if curr_order['prev_po'] == '':
            curr_order['issue_list'].append('missing prev_po')
            curr_order['prev_po'] = 'missing'

    curr_order['sub_order'] = sub_order

    # print curr_order['sub_order']
    b = 0
    curr_order['counter'] = b
    curr_order['sameday_suborder'] = 'no'
    # Suborder_target is the position of order (aka First order is position zero)

    for a in orders:
        curr_order['counter'] = b
        b += 1
        if curr_order['prev_po'] == a['po_num']:
            if curr_order['sub_order'] in ('changed', 'duplicate'):
                curr_order['sameday_suborder'] = 'yes'
                curr_order['issue_list'].append('Sameday Sub')
                curr_order['suborder_target'] = curr_order['counter']
                # TODO remove when sub order management is complete
                # curr_order['sub_order'] = 'new order'  # added as exception when sub orders weren't possible

    return curr_order

def outgrowth_stats_dob(row, curr_order, xl_workbook):
    global LFE_COUNT

    curr_order['replacement_og'] = ''
    curr_order['outgrowth'] = row[1].value.strip()
    match = re.search(r'Initial', curr_order['outgrowth'])
    match_2 = re.search(r'replacement', curr_order['outgrowth'])
    # ogConfirm = (re.findall("[a-zA-Z]"))

    if match:
        curr_order['outgrowth'] = 'yes'
        curr_order['issue_list'].append('Out Growth 1st')

    if match_2:
        curr_order['outgrowth'] = 'yes'
        curr_order['replacement_og'] = 'yes'
        curr_order['issue_list'].append('Out Growth Replacement')
        raw_prev_po = row[4].value
        curr_order['data_base'] = ''
        curr_order['sub_order'] = 'changed'
        curr_order = previousOrderSource(raw_prev_po, curr_order)
        if curr_order['data_base'] == 'LFE':
            LFE_COUNT += 1
            curr_order['issue_list'].append('LFE Order')

    curr_order['children_device'] = 'no'
    curr_order['dob'] = ''
    if curr_order['outgrowth'] == 'yes':
        curr_order['dob'] = str(row[7].value)
        print row[7], 'LOOK HERE'
        curr_order['children_device'] = 'yes'
    if row[7].ctype == 3:
        curr_order['dob'] = xlrd.xldate_as_datetime(int(row[7].value), xl_workbook.datemode).strftime("%m%d%y")
        # Needs more sophistication later to encompass all accounts
        curr_order['children_device'] = 'yes'
        # print curr_order['children_device']
        # print row
        # else:
        #      curr_order['dob'] = ""

    return curr_order

def device_code(row, curr_order):
    curr_order['device'] = str(row[1].value).strip().strip("\s")
    curr_order['device_l'] = str(row[4].value).strip().strip("\s")
    curr_order['device_r'] = str(row[7].value).strip().strip("\s")
    # Both
    poro = curr_order['device']
    device_code = deviceDictionary[poro]
    curr_order['device_code'] = device_code
    # Left
    poro_l = curr_order['device_l']
    device_code = deviceDictionary[poro_l]
    curr_order['device_code_l'] = device_code
    # Right
    poro_r = curr_order['device_r']
    device_code = deviceDictionary[poro_r]
    curr_order['device_code_r'] = device_code
    # print poro, curr_order['device_code']

    # checking right device
    if curr_order['device_code_r'] != '':
        curr_order['issue_list'].append('Right Device')
        curr_order['footedness'] = 'Right '
        curr_order['device_mod'] = 'yes'
    # checking left device
    if curr_order['device_code_l'] != '':
        curr_order['issue_list'].append('Left Device')
        curr_order['footedness'] = 'Left '
        curr_order['device_mod'] = 'yes'

    # Both
    if curr_order['sub_order'] == 'new order' or [
        'sub_order'] == 'changed':  # or curr_order['sameday_suborder'] == 'yes':
        if curr_order['device'] == '':
            curr_order['issue_list'].append('Device Missing')
            curr_order['device_code'] = '43170ST01'
    if curr_order['children_device'] == 'yes':
        # print 'child device needed', curr_order['device_code']
        if curr_order['device_code'] == '43174UC01':
            curr_order['device_code'] = '43174UC02'
            curr_order['issue_list'].append('Child Device Used')
        if curr_order['device_code'] == '43174RB01':
            curr_order['device_code'] = '43174RB02'
            curr_order['issue_list'].append('Child Device Used')
            # Left
        if curr_order['device_code_l'] == '43174UC01':
            curr_order['device_code_l'] = '43174UC02'
            curr_order['issue_list'].append('Child Device Used')
        if curr_order['device_code_l'] == '43174RB01':
            curr_order['device_code_l'] = '43174RB02'
            curr_order['issue_list'].append('Child Device Used')
            # Right
        if curr_order['device_code_r'] == '43174UC01':
            curr_order['device_code_r'] = '43174UC02'
            curr_order['issue_list'].append('Child Device Used')
        if curr_order['device_code_r'] == '43174RB01':
            curr_order['device_code_r'] = '43174RB02'
            curr_order['issue_list'].append('Child Device Used')

    return curr_order

def parse_shell(row, curr_order):
    print("row:", row)
    curr_order['shell'] = " ".join(str(row[1].value).strip().strip("\s").split(" ")[0:-1])
    curr_order['shell_thick'] = str(row[1].value).strip().strip("\s").split(" ")[-1].rstrip("mm")
    curr_order['shell_l'] = " ".join(str(row[4].value).strip().strip("\s").split(" ")[0:-1])
    curr_order['shell_l_thick'] = str(row[4].value).strip().strip("\s").split(" ")[-1].rstrip("mm")
    curr_order['shell_r'] = " ".join(str(row[7].value).strip().strip("\s").split(" ")[0:-1])
    curr_order['shell_r_thick'] = str(row[7].value).strip().strip("\s").split(" ")[-1].rstrip("mm")
    print(curr_order['shell'])
    print(curr_order['shell_thick'])
    print(curr_order['shell_l'])
    print(curr_order['shell_l_thick'])
    print(curr_order['shell_r'])
    print(curr_order['shell_r_thick'])
    return curr_order

def parse_fill_row(row, curr_order):
    curr_order['fill_l'] = str(row[4].value).strip().strip("\s")
    curr_order['fill_r'] = str(row[7].value).strip().strip("\s")
    if curr_order['fill_l'] != curr_order['fill_r']:
        curr_order['issue_list'].append('Fill Difference')

    return curr_order

def parse_extension_length(row, curr_order):
    # TODO logic to extract extension ...

    curr_order['ext-code_b'] = str(row[1].value).strip().strip("\s")
    curr_order['ext-length_b'] = str(row[3].value).strip().strip("\s")
    curr_order['ext-code_l'] = str(row[4].value).strip().strip("\s")
    curr_order['ext-length_l'] = str(row[6].value).strip().strip("\s")
    curr_order['ext-code_r'] = str(row[7].value).strip().strip("\s")
    curr_order['ext-length_r'] = str(row[9].value).strip().strip("\s")
    return curr_order

def parse_topcover_length_1(row, curr_order):
    print row
    curr_order['tc1-code_b'] = str(row[1].value).strip().strip("\s")
    curr_order['tc1-length_b'] = str(row[3].value).strip().strip("\s")
    curr_order['tc1-code_l'] = str(row[4].value).strip().strip("\s")
    curr_order['tc1-length_l'] = str(row[6].value).strip().strip("\s")
    curr_order['tc1-code_r'] = str(row[7].value).strip().strip("\s")
    curr_order['tc1-length_r'] = str(row[9].value).strip().strip("\s")
    print (curr_order['tc1-code_b'])
    print (curr_order['tc1-length_b'])
    print (curr_order['tc1-code_l'])
    print (curr_order['tc1-length_l'])
    print (curr_order['tc1-code_r'])
    print (curr_order['tc1-length_r'])
    return curr_order

def parse_topcover_length_2(row, curr_order):
    print row
    curr_order['tc2-code_b'] = str(row[1].value).strip().strip("\s")
    curr_order['tc2-length_b'] = str(row[3].value).strip().strip("\s")
    curr_order['tc2-code_l'] = str(row[4].value).strip().strip("\s")
    curr_order['tc2-length_l'] = str(row[6].value).strip().strip("\s")
    curr_order['tc2-code_r'] = str(row[7].value).strip().strip("\s")
    curr_order['tc2-length_r'] = str(row[9].value).strip().strip("\s")
    print (curr_order['tc2-code_b'])
    print (curr_order['tc2-length_b'])
    print (curr_order['tc2-code_l'])
    print (curr_order['tc2-length_l'])
    print (curr_order['tc2-code_r'])
    print (curr_order['tc2-length_r'])
    return curr_order

def parse_topcover_length_3(row, curr_order):
    print row
    curr_order['tc3-code_b'] = str(row[1].value).strip().strip("\s")
    curr_order['tc3-length_b'] = str(row[3].value).strip().strip("\s")
    curr_order['tc3-code_l'] = str(row[4].value).strip().strip("\s")
    curr_order['tc3-length_l'] = str(row[6].value).strip().strip("\s")
    curr_order['tc3-code_r'] = str(row[7].value).strip().strip("\s")
    curr_order['tc3-length_r'] = str(row[9].value).strip().strip("\s")
    print (curr_order['tc3-code_b'])
    print (curr_order['tc3-length_b'])
    print (curr_order['tc3-code_l'])
    print (curr_order['tc3-length_l'])
    print (curr_order['tc3-code_r'])
    print (curr_order['tc3-length_r'])

    return curr_order

def parse_bottom_cover_length(row, curr_order):
    print row
    curr_order['bc-code_b'] = str(row[1].value).strip().strip("\s")
    curr_order['bc-length_b'] = str(row[3].value).strip().strip("\s")
    curr_order['bc-code_l'] = str(row[4].value).strip().strip("\s")
    curr_order['bc-length_l'] = str(row[6].value).strip().strip("\s")
    curr_order['bc-code_r'] = str(row[7].value).strip().strip("\s")
    curr_order['bc-length_r'] = str(row[9].value).strip().strip("\s")
    print (curr_order['bc-code_b'])
    print (curr_order['bc-length_b'])
    print (curr_order['bc-code_l'])
    print (curr_order['bc-length_l'])
    print (curr_order['bc-code_r'])
    print (curr_order['bc-length_r'])

    return curr_order

def parse_extrinsic_rf_post(row, curr_order):
    print row
    curr_order['rf-post-code_b'] = row[1].value
    curr_order['rf-post-code_l'] = row[4].value
    curr_order['rf-post-code_r'] = row[7].value

    print (curr_order['rf-post-code_b'])
    print (curr_order['rf-post-code_l'])
    print (curr_order['rf-post-code_r'])


    return curr_order
# rf-post-type_b
# rf-post-type_l
# rf-post-type_r

def parse_extrinsic_ff_post(row, curr_order):
    print row,
    curr_order['ff-post-code_b'] = row[1].value
    curr_order['ff-post-code_l'] = row[4].value
    curr_order['ff-post-code_r'] = row[7].value

    print (curr_order['ff-post-code_b'])
    print (curr_order['ff-post-code_l'])
    print (curr_order['ff-post-code_r'])

    return curr_order

#todo wo accepts thickness separate from material, also from specific code
#todo need to combine and use item code
def parse_heel_lift(row, curr_order):
    print row
    curr_order['heel-lift-code_b'] = str(row[1].value).strip().strip("\s")
    curr_order['heel-lift-length_b'] = str(row[3].value).strip().strip("\s")
    curr_order['heel-lift-sp-code_b'] = str(row[3].value).strip().strip("\s")
    curr_order['heel-lift-code_l'] = str(row[4].value).strip().strip("\s")
    curr_order['heel-lift-length_l'] = str(row[6].value).strip().strip("\s")
    curr_order['heel-lift-sp-code_l'] = str(row[3].value).strip().strip("\s")
    curr_order['heel-lift-code_r'] = str(row[7].value).strip().strip("\s")
    curr_order['heel-lift-length_r'] = str(row[9].value).strip().strip("\s")
    curr_order['heel-lift-sp-code_r'] = str(row[9].value).strip().strip("\s")
    print (curr_order['heel-lift-code_b'])
    print (curr_order['heel-lift-length_b'])
    print (curr_order['heel-lift-sp-code_b'])
    print (curr_order['heel-lift-code_l'])
    print (curr_order['heel-lift-length_l'])
    print (curr_order['heel-lift-sp-code_l'])
    print (curr_order['heel-lift-code_r'])
    print (curr_order['heel-lift-length_r'])
    print (curr_order['heel-lift-sp-code_r'])


    return curr_order

def parse_addition_1(row, curr_order):
    print row
    curr_order['add1-code_b'] = str(row[1].value).strip().strip("\s")
    curr_order['add1-code_l'] = str(row[4].value).strip().strip("\s")
    curr_order['add1-code_r'] = str(row[7].value).strip().strip("\s")
    print (curr_order['add1-code_b'])
    print (curr_order['add1-code_l'])
    print (curr_order['add1-code_r'])
    return curr_order

def parse_addition_2(row, curr_order):
    print row
    curr_order['add2-code_b'] = str(row[1].value).strip().strip("\s")
    curr_order['add2-code_l'] = str(row[4].value).strip().strip("\s")
    curr_order['add2-code_r'] = str(row[7].value).strip().strip("\s")
    print (curr_order['add2-code_b'])
    print (curr_order['add2-code_l'])
    print (curr_order['add2-code_r'])
    return curr_order

def parse_addition_3(row, curr_order):
    print row
    curr_order['add3-code_b'] = str(row[1].value).strip().strip("\s")
    curr_order['add3-code_l'] = str(row[4].value).strip().strip("\s")
    curr_order['add3-code_r'] = str(row[7].value).strip().strip("\s")
    print (curr_order['add3-code_b'])
    print (curr_order['add3-code_l'])
    print (curr_order['add3-code_r'])

    return curr_order

def parse_heel_cup_depth(row, curr_order):
    print row
    curr_order['heel-cup-depth_b'] = row[1].value
    curr_order['heel-cup-depth_l'] = row[4].value
    curr_order['heel-cup-depth_r'] = row[7].value

    print (curr_order['heel-cup-depth_b'])
    print (curr_order['heel-cup-depth_l'])
    print (curr_order['heel-cup-depth_r'])

    return curr_order
#TODO these shell mods are weird, need to figure out how they are inputted into am
def parse_shell_modification_1(row, curr_order):
    print row
    curr_order['shell-mod1_b'] = row[1].value
    curr_order['shell-mod1_l'] = row[4].value
    curr_order['shell-mod1_r'] = row[7].value

    print (curr_order['shell-mod1_b'])
    print (curr_order['shell-mod1_l'])
    print (curr_order['shell-mod1_r'])

    return curr_order

def parse_shell_modification_2(row, curr_order):
    print row
    curr_order['shell-mod2_b'] = row[1].value
    curr_order['shell-mod2_l'] = row[4].value
    curr_order['shell-mod2_r'] = row[7].value

    print (curr_order['shell-mod2_b'])
    print (curr_order['shell-mod2_l'])
    print (curr_order['shell-mod2_r'])

    return curr_order

def parse_shell_modification_3(row, curr_order):
    print row
    curr_order['shell-mod3_b'] = row[1].value
    curr_order['shell-mod3_l'] = row[4].value
    curr_order['shell-mod3_r'] = row[7].value

    print (curr_order['shell-mod3_b'])
    print (curr_order['shell-mod3_l'])
    print (curr_order['shell-mod3_r'])

    return curr_order

def parse_cast_dressing(row, curr_order):
    print row
    curr_order['corr-type_b'] = row[1].value
    curr_order['corr-type_l'] = row[4].value
    curr_order['corr-type_r'] = row[7].value

    print (curr_order['corr-type_b'])
    print (curr_order['corr-type_l'])
    print (curr_order['corr-type_r'])

    return curr_order

def parse_foot_scanned(row, curr_order):
    curr_order['both'] = row[1].value
    curr_order['left'] = row[4].value
    curr_order['right'] = row[7].value

    if (not curr_order['both'] and not curr_order['right'] and not curr_order['left']):
        curr_order['foot2scan'] = ''
        if curr_order['sub_order'] == 'new order':
            curr_order['issue_list'].append('Foot2Scan missing')
    elif curr_order['both'] and (not curr_order['right'] and not curr_order['left']) \
            or (curr_order["right"] and curr_order["left"] and not curr_order["both"]):
        curr_order['foot2scan'] = 'BOTH'
    elif curr_order['right'] and not curr_order['left'] and not curr_order['both']:
        curr_order['foot2scan'] = 'RIGHT'
    elif curr_order['left'] and not curr_order['right'] and not curr_order['both']:
        curr_order['foot2scan'] = 'LEFT'
    else:
        curr_order['foot2scan'] = 'Human Help!'
        if curr_order['sub_order'] == 'new order':
            curr_order['issue_list'].append('Foot Scan Error')

    return curr_order

def parse_rf_post_angle(row, curr_order):
    print row
    curr_order['rf-angle_b'] = str(row[1].value).strip().strip("\s")
    curr_order['rf-angle-uom_b'] = str(row[3].value).strip().strip("\s")
    curr_order['rf-post-type_l'] = str(row[3].value).strip().strip("\s")

    curr_order['rf-angle_l'] = str(row[4].value).strip().strip("\s")
    curr_order['rf-angle-uom_l'] = str(row[6].value).strip().strip("\s")
    curr_order['rf-post-type_l'] = str(row[6].value).strip().strip("\s")

    curr_order['rf-angle_r'] = str(row[7].value).strip().strip("\s")
    curr_order['rf-angle-uom_r'] = str(row[9].value).strip().strip("\s")
    curr_order['rf-post-type_r'] = str(row[9].value).strip().strip("\s")
    print (curr_order['rf-angle_b'])
    print (curr_order['rf-angle_l'])
    print (curr_order['rf-angle_r'])
    return curr_order

# TODO names dont' make sense here, why in helen's thing it's ff ex angle when it could
# todo be extrinsic or intrinsic
# todo all 3 values have been merged, need separating
def parse_ff_post_angle(row, curr_order):
    print row
    curr_order['ff-angle_b'] = str(row[1].value).strip().strip("\s")
    curr_order['ff-angle-uom_b'] = str(row[3].value).strip().strip("\s")
    curr_order['ff-post-type_l'] = str(row[3].value).strip().strip("\s")

    curr_order['ff-angle_l'] = str(row[4].value).strip().strip("\s")
    curr_order['ff-angle-uom_l'] = str(row[6].value).strip().strip("\s")
    curr_order['ff-post-type_l'] = str(row[6].value).strip().strip("\s")

    curr_order['ff-angle_r'] = str(row[7].value).strip().strip("\s")
    curr_order['ff-angle-uom_r'] = str(row[9].value).strip().strip("\s")
    curr_order['ff-post-type_r'] = str(row[9].value).strip().strip("\s")
    print (curr_order['ff-angle_b'])
    print (curr_order['ff-angle_l'])
    print (curr_order['ff-angle_r'])
    nextrow = xl_sheet.row(order + 1)
    print nextrow

    return curr_order

def parse_skive(row, curr_order):
    print row
    curr_order['skive_b'] = row[1].value
    curr_order['skive_l'] = row[4].value
    curr_order['skive_r'] = row[7].value

    print (curr_order['skive_b'])
    print (curr_order['skive_l'])
    print (curr_order['skive_r'])
    return curr_order

def parse_inversion(row, curr_order):
    print row
    curr_order['inversion_b'] = row[1].value
    curr_order['inversion_l'] = row[4].value
    curr_order['inversion_r'] = row[7].value

    print (curr_order['inversion_b'])
    print (curr_order['inversion_l'])
    print (curr_order['inversion_r'])

    return curr_order

def parse_motion(row, curr_order):
    print row
    curr_order['motion_b'] = row[1].value
    curr_order['motion_l'] = row[4].value
    curr_order['motion_r'] = row[7].value

    print (curr_order['motion_b'])
    print (curr_order['motion_l'])
    print (curr_order['motion_r'])

    return curr_order


def parse_notes(xl_sheet, curr_order):

    nextrow = xl_sheet.row(order + 1)
    curr_order['notes'] = nextrow[0].value.upper()
    special_instructions = (re.findall("[a-zA-Z'-]+", curr_order["notes"]))
    addy = 0
    for word in special_instructions:
        if word == 'HOLD':
            curr_order['issue_list'].append('HOLD REQ')
        if word in ('ADDRESS', 'PAPERWORK', 'UPS'):
            addy += 1
    if addy > 0:
        curr_order['issue_list'].append('ALT-Address')
        # if word == 'ADDRESS':
        #     curr_order['issue_list'].append('ALT-Address')
        # if word == 'PAPERWORK':
        #     curr_order['issue_list'].append('ALT-Address')
        # if word == 'UPS':
        #     curr_order['issue_list'].append('ALT-Address')
    print nextrow
    print curr_order['notes']
    exit()
    return curr_order

def process_xlsfile(xfile):
    filepath = os.path.join(targetdir, xfile)
    # print filepath
    # dowork

    xl_workbook = xlrd.open_workbook(filepath)
    xl_sheet = xl_workbook.sheet_by_index(1)

    return xl_workbook, xl_sheet

orders = [] #########order processing############
for xfile in xlsfiles:
    xl_workbook, xl_sheet = process_xlsfile(xfile)
    # print ('Sheet name: %s' % xl_sheet.name)
    #print xlrd.xldate_as_datetime(42088,xl_workbook.datemode)

    # curr_order = []
    curr_order = None
    for order in range(xl_sheet.nrows):
        row = xl_sheet.row(order)

        # PO_NUMBER , PARSING NAME
        if row[0].value == 'PATIENT NAME / CODE NO.':
            curr_order = patient_name_ponum(row, curr_order)
        # WEIGHT, FOOT SIZE, PRIORITY
        if row[0].value == 'WEIGHT RANGE / SIZE OF FOOT / PRIORITY / TEMPLATE':
            curr_order = weight_footsize_priority(row, curr_order)
        # QUANTITY, SUB_ORDER STATUS, SAME DAY SUB ORDER, PREV_PO
        if row[0].value == 'QUANTITY / SUBSEQUENT PAIR':
            curr_order = quantity_sub_order_sameday(row, curr_order)
        # OUTGROWTH STATUS, DOB
        if row[0].value == 'OUTGROWTH PAIR / DOB':
            curr_order = outgrowth_stats_dob(row, curr_order, xl_workbook)
        # DEVICE CODE
        if row[0].value == 'DEVICE':
            curr_order = device_code(row, curr_order)
        # SHELL
        if row[0].value == 'SHELL':
            curr_order = parse_shell(row, curr_order)
        # FILL DIFFERENCE
        if row[0].value == "FILL":
            curr_order = parse_fill_row(row, curr_order)
        # EXTENSION / LENGTH
        if row[0].value == 'EXTENSION / LENGTH':
            #TODO Not implemented yet
            curr_order = parse_extension_length(row, curr_order)
        # TOP COVER #1 / LENGTH
        if row[0].value == 'TOP COVER #1 / LENGTH':
            curr_order = parse_topcover_length_1(row, curr_order)
        # TOP COVER #2 / LENGTH
        if row[0].value == 'TOP COVER #2 / LENGTH':
            curr_order = parse_topcover_length_2(row, curr_order)
        # TOP COVER #2 / LENGTH
        if row[0].value == 'TOP COVER #3 / LENGTH':
            curr_order = parse_topcover_length_3(row, curr_order)
        # BOTTOM COVER / LENGTH
        if row[0].value == 'BOTTOM COVER / LENGTH':
            curr_order = parse_bottom_cover_length(row, curr_order)
        # EXTRINSIC REARFOOT POST
        if row[0].value == 'EXTRINSIC REARFOOT POST':
            curr_order = parse_extrinsic_rf_post(row, curr_order)
        # EXTRINSIC FOREFOOT POST
        if row[0].value == 'EXTRINSIC FOREFOOT POST':
            curr_order = parse_extrinsic_ff_post(row, curr_order)
        # LIFT (mm)
        if row[0].value == 'HEEL LIFT':
            curr_order = parse_heel_lift(row, curr_order)
        # ADDITION #1
        if row[0].value == 'ADDITION #1':
            curr_order = parse_addition_1(row, curr_order)
        # ADDITION #2
        if row[0].value == 'ADDITION #2':
            curr_order = parse_addition_2(row, curr_order)
        # ADDITION #3
        if row[0].value == 'ADDITION #3':
            curr_order = parse_addition_3(row, curr_order)
        # HEEL CUP DEPTH (mm)
        if row[0].value == 'HEEL CUP DEPTH (mm)':
            curr_order = parse_heel_cup_depth(row, curr_order)
        # SHELL MODIFICATION #1
        if row[0].value == 'SHELL MODIFICATION #1':
            curr_order = parse_shell_modification_1(row, curr_order)
        # SHELL MODIFICATION #2
        if row[0].value == 'SHELL MODIFICATION #2':
            curr_order = parse_shell_modification_2(row, curr_order)
        # SHELL MODIFICATION #3
        if row[0].value == 'SHELL MODIFICATION #3':
            curr_order = parse_shell_modification_3(row, curr_order)
        # CAST DRESSING
        if row[0].value == 'CAST DRESSING':
            curr_order = parse_cast_dressing(row, curr_order)
        # FOOT TO SCAN
        if row[0].value == 'FOOT SCANNED':
            curr_order = parse_foot_scanned(row, curr_order)
        # REARFOOT POST ANGLE
        if row[0].value == 'REARFOOT POST ANGLE':
            curr_order = parse_rf_post_angle(row, curr_order)
        # FOREFOOT POST ANGLE
        if row[0].value == 'FOREFOOT POST ANGLE':
            curr_order = parse_ff_post_angle(row, curr_order)
        # SKIVE
        # TODO there is a useless space after here, need fixing
        if row[0].value == 'SKIVE ':
            curr_order = parse_skive(row, curr_order)
        # INVERSION
        if row[0].value == 'INVERSION':
            curr_order = parse_inversion(row, curr_order)
        # MOTION
        if row[0].value == 'MOTION':
            curr_order = parse_motion(row, curr_order)
        # SPECIAL INSTRUCTIONS
        if row[0].value == 'NOTES':
            curr_order = parse_notes(xl_sheet, curr_order)
            # End of order, append
            orders.append(curr_order)


##**************************************ALL ORDERS HAVE BEEN PROCESSED*******************************************

### PHASE 2: AUTO-DYLAN


name_position = 1
for person in orders:
    print name_position, '(' + person['po_num'] + ')', person['name']
    name_position += 1


#=================================<<<< CONTROL PANEL >>>>========================================


# pyautogui.click(100, 100)  # For Notepad
ORDER_START = 0   # for where you want to start in batch should be default '0'
oe_gate = 'ye1s'  # Gates OE part  (needs to be changed to something better for all orders)


ORDER_LIMIT = 1000  # How many orders from start you want (needs to be changed to pick which orders to do)
TIME_DELAY = 1.5  # Delay between most actions default: 1.5
SHORT_DELAY = 0.25  # Delay between quick actions default: 0.25
LAG = 36  # Delay for in and out of pt search default: 40, 23
sticker_gate = 'yes'  # Gates labels print out
file_rename = 'yes'  # Gates Raw file renaming
if oe_gate == 'yes':
    pyautogui.click(1046, 743)  # For OE

#=========================================================================================


def printOrderIntoCard(order, k, SUBORDER_COUNT, NEWORDER_COUNT):
    if order['po_num'] != '':
        if order['sub_order'] == 'new order' or 'duplicate' or 'changed':
            print '\n', '\n', '\n'
            print '==============================================================================='
            print '                          ', 'ORDER', k, ': Position', k-1, '\n'
            print order['sub_order']
            print order['priority']
            print '_________________________________'
            print 'FIRST NAME =', order['FIRSTNAMECOMPLETE']
            print 'LAST NAME  =', order['nameLast']
            print 'PO NUMBER =', order['po_num']
            print 'DOB =', order['dob'] + '\n' + 'O.G. =', order['outgrowth']
            print '_________________________________'
            print 'WEIGHT =', order['weight']
            print 'SHOE SIZE =', order['shoesize']
            print 'FOOT TO SCAN =', order['foot2scan']
            print '_________________________________'
            print 'DEVICE =', order['device'], '(', order['device_code'], ')'
            if order['children_device'] == 'yes':
                print "['children_device'] ==", order['children_device']
            print 'DEVICE L =', order['device_l'], "DEVICE R =", order['device_r']
            print 'FILL L =', order['fill_l'], 'FILL R=', order['fill_r']
            print 'QUANTITY =', order['quantity']
            print 'PREVIOUS PO# =', order['data_base'], order['prev_po']
            print '_________________________________'
            print 'NOTES =', order['notes']
            print ''
            print 'SAMEDAY SUBORDER =', order['sameday_suborder'] + '\n' + 'ORDER POSITION UP TO NOW =', order['counter']
            print 'SUBORDER TARGET =', order['suborder_target']


# AM INTERFACE FUNCTIONS

## Tab nav functions

def goto_general_order_info_tab(pyautogui):
    pyautogui.hotkey('alt', 'g')

def goto_dev_shell_mods_tab(pyautogui):
    pyautogui.hotkey('alt', 'm')

def goto_components_tab(pyautogui):
    pyautogui.hotkey('alt', 'c')

def goto_post_cast_dressing_tab(pyautogui):
    pyautogui.hotkey('alt', 'd')

def goto_discounts_charges_tab(pyautogui):
    pyautogui.hotkey('alt', 'h')

def fetchSubsequentOrder(order):  # Banana Grabber TM
    # Opens sales order search
    if order['prev_po'] == "missing":
        return '1'
    pyautogui.typewrite(['down', 'return'])
    while pyautogui.locateOnScreen(posted_sales_button, confidence=0.97, region=(960, 0, 960, 540)) is None:
        print "*Scanning for Sales Order Screen"
        time.sleep(1)
    else:
        print '**Sales Order Screen -Detected-'
        time.sleep(0.5)
        pyautogui.click(
            pyautogui.locateCenterOnScreen(posted_sales_button, confidence=0.97, region=(1264, 164, 150, 56)))
        print 'Opening Posted Sales...'
        count = 1
        while pyautogui.locateOnScreen(existing_orders_button, confidence=0.97, region=(960, 0, 960, 540)) is None:
            print '*Searching for Posted Sales', count
            if count == 10:
                pyautogui.click(
                    pyautogui.locateCenterOnScreen(posted_sales_button, confidence=0.97, region=(1264, 164, 150, 56)))
            count += 1
        else:
            print '-Posted Sales Detected-'

    # to clear any filters or previous auto searches
    pyautogui.hotkey('alt', 'c')
    # Click with in search result field
    pyautogui.click(850, 450)
    time.sleep(1)
    # Move to the most left, bringing you to po_num column
    pyautogui.typewrite(['home'])
    pyautogui.typewrite(order['prev_po'])
    pyautogui.typewrite(['return'])
    time.sleep(TIME_DELAY)
    x = 1
    while pyautogui.locateOnScreen(sub_search_confirm) is not None:
        time.sleep(1)
        x += 1
        print '*Confirming prev_po Provided is Valid#', x
        if x == 20:
            order['issue_list'].append('prev_po search failure')
            return '1'
    else:
        print '**Prev_po -Validated-'
        #pyautogui.typewrite(['return'])

    # TODO it checked if it was there which is good, but also needs to check that after enter
    # TODO that it's not there anymore, aka search was successful
    # If Statement skips code if prev ponum fails

    pyautogui.typewrite(['right'])
    pyautogui.typewrite(['down', 'up'], interval=0.3)
    pyautogui.hotkey('ctrl', 'c')
    order['prev_wo'] = (Tk().clipboard_get())
    print 'Previous WO:', order['prev_wo']
    time.sleep(0.5)

    pyautogui.typewrite(['right'])
    pyautogui.typewrite(['down', 'up'], interval=0.3)
    pyautogui.hotkey('ctrl', 'c')
    order['prev_ptnum'] = (Tk().clipboard_get())
    order['pt_num'] = order['prev_ptnum']
    print 'Previous Pt#', order['prev_ptnum']
    time.sleep(0.5)

    pyautogui.typewrite(['right'])
    pyautogui.typewrite(['down', 'up'], interval=0.3)
    pyautogui.hotkey('ctrl', 'c')
    order['prev_device'] = (Tk().clipboard_get())
    print order['prev_device']
    order['prev_device'] = order['prev_device'].upper()
    order['prev_device'] = order['prev_device'].rstrip(' ')
    print 'Previous Device', order['prev_device']
    order['prev_device'] = subDeviceDictionary[order['prev_device']]
    print 'Previous Device', order['prev_device']
    time.sleep(0.5)

    # Determining between true dupes and false dupes
    #todo true dupe, device = '' also compare names

    if order['device_code'] == order['prev_device']:
        order['alter_subdevice'] = 'no'
        print "Alter SubDevice:", order['alter_subdevice']
    else:
        order['alter_subdevice'] = 'yes'
        print "Alter SubDevice:", order['alter_subdevice']

def orderCreation(order):
    if order['sameday_suborder'] == 'yes':
        if orders[order['suborder_target']]['data_base'] == 'LFE':
            order['issue_list'].append('SAMEDAY_SUB LFE target')
            return '1'
    ## Open order
    # Selecting 'Paris Sales Order'
    #TODO image match general screen
    pyautogui.typewrite(['return'])
    m = 0
    while pyautogui.locateOnScreen(general_tab, confidence=0.97, region=(0, 170, 1920, 200)) is None:
        time.sleep(1)
        if m == 20:
            print "-Order Creation Reset Attempt-"
            pyautogui.typewrite(['esc', 'esc', 'up', 'up', 'up', 'return'])
        if m == 30:
            return '1'
        m += 1
        print '*Scanning for General Tab', m
    else:
        print '**General Tab Detected'
    time.sleep(0.25)
    # Create new order
    pyautogui.press(['f3'])
    ## TODO need to figure out why it's not seeing it

    # while pyautogui.locateOnScreen(new_order, region=(0, 0, 960, 1080)) is None:
    #     print "*Scanning for New Order"
    #     time.sleep(1)
    # else:
    #     print '**New Order Detected'
    # Move to work order number field

    ## in place of above code
    time.sleep(3)
    pyautogui.typewrite(['tab', 'up', 'up', 'up', 'up', 'up'], interval=SHORT_DELAY)
    time.sleep(0.5)
    # Copy work order number to clipboard
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.5)

    # Get created workorder number from clipboard
    # TODO: Maybe don't need () for grabbing from clipboard
    order['wo_num'] = (Tk().clipboard_get())
    print 'wo_num', order['wo_num']

    ## Select firefly as account
    # Select account number field
    goto_general_order_info_tab(pyautogui)
    time.sleep(TIME_DELAY)
    # Enter firefly account number
    pyautogui.press(['0', '7', '4', '5', 'return'])
    time.sleep(SHORT_DELAY)

    ## SELECT CLINICIAN (MAIN)
    # Opens clinician search screen
    pyautogui.typewrite('2172')
    time.sleep(SHORT_DELAY)

    # Confirms selection and moves to patient card search
    pyautogui.typewrite(['return'])


def patientCardHandler(order):

    if order['sameday_suborder'] == 'yes':
        pyautogui.typewrite(orders[order['suborder_target']]['pt_num'])

    else:
        if order['buddy_order'] == 'yes':
            pyautogui.typewrite(orders[order['suborder_target']]['pt_num'])
        elif order['sub_order'] in ('duplicate', 'changed'):
            pyautogui.typewrite(order['prev_ptnum'])
        elif order['sub_order'] == 'new order':
            pyautogui.typewrite(['f6'])
            time.sleep(LAG)  # ENTERING PT SEARCH, LONG DELAY
            # CREATING NEW CARD VIA SEARCHING
            # Checking to see if opening search screen has resolved

            count = 1

            while pyautogui.locateOnScreen(pt_search_entry, confidence=0.95, region=(0, 0, 1920, 540)) is None:
                print '*Opening Pt Search Screen Scan #', count
                count += 1
                time.sleep(2)
                while pyautogui.locateOnScreen(not_responding, confidence=0.95, region=(0, 0, 1920, 540)) is not None:
                    time.sleep(4.5)
                    print 'Program Unresponsive, waiting....'
                if count == 40:
                    return '1'
                if count == 8:
                    order['issue_list'].append('Opening Pt Search Slow')

            else:
                print '**Search Screen -Detected-'

            pyautogui.typewrite(order['FIRSTNAMECOMPLETE'])
            pyautogui.typewrite(['tab'])
            pyautogui.typewrite(order['nameLast'])
            # pyautogui.typewrite(['tab'])
            # pyautogui.typewrite(i["nameLast"])
            print 'Searching Pt Name...'
            pyautogui.hotkey('alt', 's')
            pyautogui.hotkey('alt', 'c')
            pyautogui.hotkey('alt', 's')
            pyautogui.hotkey('alt', 'c')
            # SEARCHING FOR PT CREATES LONG DELAY
            time.sleep(LAG)

            x = 1
            found = 'no'
            while found == 'no':
                print '*New Pt Popup Search #', x
                if x == 7:
                    pyautogui.hotkey('alt', 's')
                    pyautogui.hotkey('alt', 'c')
                    print 'Extra time to resolve...'
                    time.sleep(10)
                time.sleep(2)
                # Should all of this be abandoned and just force the pt card to happen regardless if there is a similar name
                # To see New Pt Pop Up Dialogue box
                while pyautogui.locateOnScreen(not_responding, confidence=0.95, region=(0, 0, 1920, 540)) is not None:
                    time.sleep(4.5)
                    print 'Program Unresponsive, waiting....'
                if pyautogui.locateOnScreen(pt_popup, confidence=0.90, region=(803, 457, 321, 186)) is not None:
                    print '**Popup Detected'
                    found = 'yes'
                    pyautogui.typewrite(['y'])
                    # TODO while loop for searching if it's inside the pt card yet
                    time.sleep(1)
                    # entering pt info
                    pyautogui.typewrite(['tab', 'tab', 'm', 'tab'])  # gender
                    time.sleep(SHORT_DELAY)
                    pyautogui.typewrite(order['dob'])
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite(order['weight'])
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite(order['shoesize'])
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite(order['shoesize'])
                    time.sleep(1)
                    if order['outgrowth'] == 'yes':
                        pyautogui.typewrite(['down', 'space'])
                        time.sleep(1)
                    else:
                        pyautogui.typewrite(['down'])

                    pyautogui.typewrite(['esc', 'return'], interval=SHORT_DELAY)
                    pyautogui.hotkey('ctrl', 'c')
                    order['pt_num'] = Tk().clipboard_get()
                    print 'pt_num =', order['pt_num']
                    time.sleep(TIME_DELAY)
                if x == 30:
                    return '1'
                x += 1

                if x > 9:
                    print '*Pt Search Completed? #', x - 9
                    time.sleep(SHORT_DELAY)
                    if pyautogui.locateOnScreen(similar_ptname, confidence=0.90, region=(0, 0, 960, 540)) is not None:
                        image7 = pyautogui.screenshot('ptname_{}.png'.format(order['po_num']))
                        print '**Pt Search has resolved, Similar Name -Assumed-'
                        pyautogui.hotkey('shift', 'f5')
                        time.sleep(SHORT_DELAY)
                        found = 'yes'
                        order['issue_list'].append('Similar Pt Name Issue')
                        time.sleep(TIME_DELAY)
                        pyautogui.typewrite(['f3', 'tab'])
                        pyautogui.typewrite(order['FIRSTNAMECOMPLETE'])
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(order['nameLast'])
                        pyautogui.typewrite(['tab', 'm', 'tab'])
                        time.sleep(TIME_DELAY)
                        pyautogui.typewrite(order["dob"])
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(order['weight'])
                        time.sleep(TIME_DELAY)
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(order['shoesize'])
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(order['shoesize'])
                        time.sleep(TIME_DELAY)
                        if order['outgrowth'] == 'yes':
                            pyautogui.typewrite(['down', 'space'])
                            time.sleep(TIME_DELAY)
                        else:
                            pyautogui.typewrite(['down'])
                        pyautogui.typewrite(['esc'])
                        time.sleep(SHORT_DELAY)
                        pyautogui.hotkey('ctrl', 'end')
                        time.sleep(SHORT_DELAY)
                        pyautogui.typewrite(['return'])
                        time.sleep(SHORT_DELAY)
                        pyautogui.hotkey('ctrl', 'c')
                        order['pt_num'] = Tk().clipboard_get()
                        print 'pt_num =', order['pt_num']
                        time.sleep(TIME_DELAY)


def footImpressionAndPONumberEntry(order, orders):
    # Moves from patient card number field to foot impression type field
    # was changed from just pyautogui.typeright(['tab'])
    goto_general_order_info_tab(pyautogui)
    pyautogui.typewrite(['right'])
    time.sleep(TIME_DELAY)

    if order['sameday_suborder'] == 'yes':
        # Use information from the original order
        if orders[order['suborder_target']]['foot2scan'] == 'Human Help!':
            # Original order had format issues
            pyautogui.typewrite(['space', 'tab'])
            time.sleep(SHORT_DELAY)
        else:
            pyautogui.typewrite(['delete', 'tab', 'tab'])
            # Lookup previous foot2scan
            pyautogui.typewrite(orders[order['suborder_target']]['foot2scan'])
            pyautogui.typewrite(['tab'])
    else:
        # For a new order
        if order['sub_order'] == 'new order' or order['replacement_og'] == 'yes':
            if order['foot2scan'] == 'Human Help!':
                pyautogui.typewrite(['space', 'tab'])
                time.sleep(SHORT_DELAY)
            else:
                # Process is normal
                pyautogui.typewrite(['a', 'tab', 'tab'], interval=SHORT_DELAY)
                time.sleep(SHORT_DELAY)
                # Put foot2scan
                pyautogui.typewrite(order['foot2scan'])
                pyautogui.typewrite(['tab'])

        elif order['sub_order'] in ('changed', 'duplicate'):
            # Move to foot2scan field for subsequent orders
            # TODO: Need to develop more (LFE and AM orders)
            pyautogui.typewrite(['delete', 'tab', 'tab', 'tab'])

    # PURCHASE ORDER NUMBER (MAIN)
    # Moves to PO number field
    pyautogui.typewrite(['tab'])
    # Write PO number
    pyautogui.typewrite(order['po_num'])


def press_add_button(pyautogui, add_button):
    pyautogui.click(pyautogui.locateCenterOnScreen(add_button,
                                                   confidence=0.95,
                                                   region=(1000, 345, 160, 140)))
    print "add button clicked"

def check_problem_image(target_problem_image, pyautogui, order, file_name='deadlock_{}.png'):
    if pyautogui.locateOnScreen(target_problem_image, confidence=0.95, region=(644, 387, 572, 320)):
        _ = pyautogui.screenshot(file_name.format(order['po_num']))
        return True
    return False

def check_image_device_select(target_image, add_button, order, pyautogui, file_name, case):
    """
    Helper function for deviceSelection. Finds target_image then reduces the device selection
    process for suborder == neworder
    """
    if check_problem_image(target_image, pyautogui, order, file_name):
        pyautogui.typewrite(['return']) # remove dialogue box
        goto_dev_shell_mods_tab(pyautogui) # return to
        if case == "new_order":
            pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
            pyautogui.typewrite(order['device_code'])
        elif case == "duplicate":
            pyautogui.typewrite(['down', 'down'])
            time.sleep(0.5)
            pyautogui.typewrite(order['prev_wo'])
            time.sleep(0.5)
            pyautogui.typewrite(['down', 'space'])
            time.sleep(0.5)
            # for now will act like all dupe orders are dupes
        elif case == "changed":
            pyautogui.typewrite(['down', 'down'])
            pyautogui.typewrite(order['prev_wo'])
            pyautogui.typewrite(['down', 'down', 'space'])
            if order['alter_subdevice'] == 'yes':
                if order['device_code'] == '':
                    pyautogui.typewrite(['return', 'return', 'return'])
                else:
                    pyautogui.typewrite(['space', 'tab', 'tab', 'tab'])
                    pyautogui.typewrite(order['device_code'])
            if order['alter_subdevice'] == 'no':
                pyautogui.typewrite(['return', 'return', 'return'])
            pyautogui.typewrite(['return', 'return'])
        elif case == "sameday_suborder":
            pyautogui.typewrite(orders[order['suborder_target']]['wo_num'])
            pyautogui.typewrite(['down'])
            pyautogui.typewrite('changed device')
            goto_dev_shell_mods_tab(pyautogui)
            pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
            pyautogui.typewrite(order['device_code'])
            if order['device_code'] == '':
                curr_order['issue_list'].append('Device Missing')
                pyautogui.typewrite(orders[order['suborder_target']]['device_code'])
        press_add_button(pyautogui, add_button)


def check_deadlock(case):
    check_image_device_select(deadlock, add_button, order, pyautogui, 'deadlock_{}.png', case)
    check_image_device_select(interupt, add_button, order, pyautogui, 'interupt_{}.png', case)

def deviceSelection(order):
    goto_dev_shell_mods_tab(pyautogui)
    count = 1
    while pyautogui.locateOnScreen(device_tab, confidence=0.95, region=(0, 0, 1920, 1080)) is None:
        if count == 30:
            return '1'
        print "*Scanning for Device Tab #", count
        count += 1
        time.sleep(2)
    print '*Device Tab Detected'

    if order['sameday_suborder'] == 'yes':
        count = 1
        pyautogui.typewrite(orders[order['suborder_target']]['wo_num'])
        pyautogui.typewrite(['down'])
        pyautogui.typewrite('changed device')
        goto_dev_shell_mods_tab(pyautogui)
        pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
        pyautogui.typewrite(order['device_code'])
        if order['device_code'] == '':
            curr_order['issue_list'].append('Device Missing')
            pyautogui.typewrite(orders[order['suborder_target']]['device_code'])
        press_add_button(pyautogui, add_button)

        while pyautogui.locateOnScreen(remove_button, confidence=0.95, region=(1070, 378, 109, 60)) is None:
            if count > 12:
                check_deadlock('sameday_suborder')
            if count == 30:
                return '1'
            print '*Resolving Device Selection #', count
            time.sleep(4.5)
            count += 1
        else:
            print '**Device Selection Resolved'
            if order['quantity_num'] > 1:
                order['issue_list'].append('Sub Quantity')
                goto_dev_shell_mods_tab(pyautogui)
                pyautogui.typewrite(['down', 'down', 'down', 'down', 'down', 'down'])
                pyautogui.typewrite(order['quantity'])

    else:
        if order['sub_order'] == 'new order' or order['replacement_og'] == 'yes':
            count = 1
            pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
            pyautogui.typewrite(order['device_code'])
            press_add_button(pyautogui, add_button)
            while pyautogui.locateOnScreen(remove_button, confidence=0.95, region=(1070, 378, 109, 60)) is None:
                if count > 12:
                    check_deadlock(case="new_order")
                if count == 30:
                    return '1'
                print '*Resolving Device Selection', count
                time.sleep(4.5)
                count += 1
            else:
                print '**Device Selection Resolved'

        if order['sub_order'] == 'changed' and order['replacement_og'] != 'yes':
            if order['data_base'] == 'AM':
                count = 1
                pyautogui.typewrite(['down', 'down'])
                pyautogui.typewrite(order['prev_wo'])
                pyautogui.typewrite(['down', 'down', 'space'])
                if order['alter_subdevice'] == 'yes':
                    if order['device_code'] == '':
                        pyautogui.typewrite(['return', 'return', 'return'])
                    else:
                        pyautogui.typewrite(['space', 'tab', 'tab', 'tab'], interval=0.2)
                        pyautogui.typewrite(order['device_code'])
                if order['alter_subdevice'] == 'no':
                    pyautogui.typewrite(['return', 'return', 'return'])
                    press_add_button(pyautogui, add_button)

                # Confirming process completion of device selection
                while pyautogui.locateOnScreen(remove_button, confidence=0.95, region=(1070, 378, 109, 60)) is None:
                    if count > 12:
                        check_deadlock(case='changed')
                    if count == 30:
                        return '1'
                    print '*Resolving Device Selection#', count
                    time.sleep(4.5)
                    count += 1
                else:
                    print '**Device Selection Resolved'
                    if order['quantity_num'] > 1:
                        order['issue_list'].append('Sub Quantity')
                        goto_dev_shell_mods_tab(pyautogui)
                        pyautogui.typewrite(['down', 'down', 'down', 'down', 'down', 'down'])
                        pyautogui.typewrite(order['quantity'])

        # TODO: Review code here for correctness, add quantity exception and normal dupe functionality
        # TODO: AM LFE differenciation

        if order['sub_order'] == 'duplicate':
            if order['data_base'] == 'AM':
                count = 1
                pyautogui.typewrite(['down', 'down'])
                time.sleep(0.5)
                pyautogui.typewrite(order['prev_wo'])
                time.sleep(0.5)
                pyautogui.typewrite(['down', 'space'])
                time.sleep(0.5)
                # for now will act like all dupe orders are dupes this selects device
                pyautogui.typewrite(['tab', 'tab', 'tab', 'tab', 'tab', 'return'], interval=0.20)
                print 'add button pressed'

                # TODO - need to make code to determin when a dupe order is actually a changed order
                # TODO - code will be left here as naive until more developement done

                while pyautogui.locateOnScreen(remove_button, confidence=0.95, region=(1070, 378, 109, 60)) is None:
                    print '*Resolving Device Selection#', count
                    if count > 10:
                        check_deadlock(case="duplicate")
                    if count == 30:
                        return '1'
                    time.sleep(4.5)
                    count += 1
                else:
                    print '**Device Selection Resolved'
                    if order['quantity_num'] > 1:
                        order['issue_list'].append('Sub Quantity')
                        goto_dev_shell_mods_tab(pyautogui)
                        pyautogui.typewrite(['down', 'down', 'down', 'down', 'down', 'down'])
                        pyautogui.typewrite(order['quantity'])

def setOrderPriority(order):
    goto_general_order_info_tab(pyautogui)
    while pyautogui.locateOnScreen(general_tab, confidence=0.95, region=(0, 0, 1920, 540)) is None:
        time.sleep(1)
        print '*Scanning for General Tab'
    else:
        print '**General Tab Detected'

    pyautogui.typewrite(['right'])
    # To delete auto filled foot impression type
    if order['sub_order'] in ('duplicate', 'changed') and order['replacement_og'] != 'yes':
        pyautogui.typewrite(['delete'])

    # Move from foot impression type to shipment priority
    pyautogui.typewrite(['tab', 'tab', 'tab'])
    if order['priority'] == 'RRU On Time':
        pyautogui.typewrite(['s', 'p', 'return', 'up', 'right', 'space'], interval=SHORT_DELAY)
        time.sleep(5)

    if order['priority'] == '3day Rush':
        pyautogui.typewrite(['3', 'return', 'return', 'up', 'right', 'space'], interval=SHORT_DELAY)
        time.sleep(5)

def fileRenamer(order):
    #raw_location = "C:/Users/Sleepy Face/PycharmProjects/paris_ortho/firefly_orders/raw_files/"
    # raw_location = "F:\ReceivedScans\ESCAN - FIREFLY\2018\06-2018\JUN06\New Folder\"
    raw_location = "C:/Users/dylanchow/PycharmProjects/work2/firefly_orders/raw_files/"

    # File checker
    def checkAndRenameFoot(foot):
        clean_expected_file = order['po_num'] + foot
        # if clean_expected_file in cleanfile_to_rawfile:
        #
        if clean_expected_file in cleanfile_to_rawfile:
            order['old_path' + foot] = raw_location + cleanfile_to_rawfile[clean_expected_file]
            print 'EXPECTED =', order['old_path' + foot]
            order['new_path' + foot] = raw_location + order['wo_num'].lstrip("0") + foot + '.RAW'
            # We know something exists
            if cleanfile_to_rawfile[clean_expected_file] not in rawname_duplicates:
                # And its not a duplicate
                os.rename(order['old_path' + foot], order['new_path' + foot])
                print 'RENAMED TO =', order['new_path' + foot]
            else:
                order['issue_list'].append('Duplicate Raw Files')
                print 'Duplicate Raw Files'
        else:
            order['issue_list'].append('Raw File Missing')
            print 'Raw File Missing'
        # Gross

    # Files get checked
    if order['sub_order'] == 'new order' or order['replacement_og'] == 'yes':
        print 'Raw file renaming paths'
        if order['foot2scan'] == 'BOTH':
            checkAndRenameFoot("L")
            checkAndRenameFoot("R")

        if order['foot2scan'] in ('LEFT', 'RIGHT'):

            if order['foot2scan'] == 'LEFT':
                checkAndRenameFoot("L")
            if order['foot2scan'] == 'RIGHT':
                checkAndRenameFoot("R")




#========================================================================================================
# Marking the Start time of whole OE process
process_start = datetime.datetime.now()
# The execution of pyautogui entering in orders                  <----------------------<<
k = 0
if oe_gate == "yes":
    pyautogui.typewrite(['up', 'up'], interval=0.3)
for i in range(len(orders)):
    order = orders[i]
    k = i+1
    error_code = ''
    # Prints out processed orders to a card
    printOrderIntoCard(order, k, SUBORDER_COUNT, NEWORDER_COUNT)

    # Start of Interfacing with the AM                                        <----------------<<
    if order['po_num'] != '':  # doesn't make orders for blank orders
        # TODO: Change gate from 'new order' to not database lfe?
        # if order['sub_order'] == oe_gate:
        if oe_gate == "yes":
            if order['data_base'] != "LFE":
                if ORDER_START <= k <= ORDER_LIMIT:
                    continue_code = '0'
                    # Order entering Reset
                    pyautogui.typewrite(['enter', 'enter'], interval=0.3)
                    time.sleep(0.5)
                    pyautogui.typewrite(['esc', 'esc', 'esc', 'esc', 'esc', 'up', 'up'], interval=0.3)
                    time.sleep(1)
                    # Marking the time of start of order
                    orderstart_time = datetime.datetime.now()
                    # Only AM2AM Sub Orders
                    if order['sameday_suborder'] == 'no':
                        if order['sub_order'] in ('duplicate', 'changed'):
                            continue_code = fetchSubsequentOrder(order)
                            if continue_code == '1':
                                print '<<Order Failed, Proceeding to Next Order (sub fetch)>>'
                                order['issue_list'].append('ORDER SKIPPED-sub_fetch')
                                image = pyautogui.screenshot('skip_subfetch_{}.bmp'.format(order['po_num']))
                                continue
                            pyautogui.typewrite(['esc', 'esc', 'esc', 'up', 'up'], interval=0.3)


                    # Start new order and collect wo#
                    continue_code = orderCreation(order)
                    if continue_code == '1':
                        print '<<Order Failed, Proceeding to Next Order(order creation)>>'
                        order['issue_list'].append('ORDER SKIPPED-order_creation')
                        image2 = pyautogui.screenshot('skip_ocreation_{}.bmp'.format(order['po_num']))
                        continue
                    # CREATE PATIENT CARD VIA SEARCH (MAIN)
                    continue_code = patientCardHandler(order)
                    if continue_code == '1':
                        print '<<Order Failed, Proceeding to Next Order(pt card)>>'
                        order['issue_list'].append('ORDER SKIPPED-pt_card')
                        image3 = pyautogui.screenshot('skip_ptcard{}.bmp'.format(order['po_num']))
                        continue
                    # IMPRESSION TYPE, FOOT TO SCAN (MAIN)
                    footImpressionAndPONumberEntry(order, orders)
                    # DEVICE SELECTION (MAIN)
                    continue_code = deviceSelection(order)
                    if continue_code == '1':
                        print '<<Order Failed, Proceeding to Next Order(device select)>>'
                        order['issue_list'].append('ORDER SKIPPED-device_selection')
                        image4 = pyautogui.screenshot('skip_device_{}.bmp'.format(order['po_num']))
                        continue
                    # FILE RENAMING TESTING
                    if file_rename == 'yes':
                        fileRenamer(order)
                    # PRIORITY SETTING RUSH/ONTIME
                    setOrderPriority(order)
                    # PRINT OUT STICKER LABELS
                    if sticker_gate == 'yes':
                        if order['data_base'] is not 'LFE':
                            pyautogui.hotkey('alt', 'e')
                            marker = 1
                            while pyautogui.locateOnScreen(after_print, confidence=0.95, region=(0, 920, 1920, 160)) is None:
                                print "*Scanning for Label Print Completion #", marker
                                time.sleep(0.8)
                                if marker > 10:
                                    if pyautogui.locateOnScreen(status_history, confidence=0.95, region=(644, 387, 572, 320)):
                                        image5 = pyautogui.screenshot('label_{}.bmp'.format(order['po_num']))
                                        pyautogui.typewrite(['return'])
                                        pyautogui.hotkey('alt', 'e')
                                    if pyautogui.locateOnScreen(deadlock, confidence=0.95,
                                                                region=(644, 387, 572, 320)):
                                        image5 = pyautogui.screenshot('deadlock_{}.bmp'.format(order['po_num']))
                                        pyautogui.typewrite(['return'])
                                        pyautogui.hotkey('alt', 'e')
                                marker += 1
                            else:
                                print '*Label Printed'
                                pyautogui.typewrite(['esc'])


                    ##  SECOND PASS ORDER ENTRY

                    # Order Quanity 2 or more
                    q_plus = order['quantity_num']
                    # Gate for only new orders not sub orders
                    if order['sub_order'] == 'new order':
                        if q_plus - 1 > 0:
                            q_plus = q_plus - 1
                            # TODO need to figure out to have an exception quantity == true for this func, what neils
                            # TODO tried to do, accessing only a section of the function only if one things is true
                            #orderCreation(order)
                            pyautogui.press(['f3'])
                            # while pyautogui.locateOnScreen(new_order, region=(0, 0, 1920, 1080)) is None:
                            #     print "*Scanning for New Order"
                            #     time.sleep(1)
                            # else:
                            #     print '**New Order Detected'
                            time.sleep(3)

                            goto_general_order_info_tab(pyautogui)
                            while pyautogui.locateOnScreen(general_tab, confidence=0.95, region=(0, 0, 1920, 540)) is None:
                                time.sleep(1)
                                print '*Scanning for General Tab'
                            else:
                                print '**General Tab Detected'

                            pyautogui.press(['0', '7', '4', '5', 'return'])
                            time.sleep(TIME_DELAY)
                            pyautogui.typewrite('2172')
                            time.sleep(TIME_DELAY)

                            # Confirms selection and moves to patient card search
                            pyautogui.typewrite(['return'], interval=TIME_DELAY)

                            #pick pt card by saved pt_num number
                            pyautogui.typewrite(order['pt_num'])
                            time.sleep(TIME_DELAY)

                            # TODO: Check if footImpression function works here
                            # See if special case with quanity=True matches the behavior of the commented out code
                            # footImpressionAndPONumberEntry(order, orders, quantity=True)
                            # # Impression type, foot to scan
                            pyautogui.typewrite(['tab', 'delete', 'tab', 'tab'], interval=SHORT_DELAY)
                            time.sleep(SHORT_DELAY)
                            pyautogui.typewrite(order['foot2scan'])
                            pyautogui.typewrite(['tab'])

                            # PO number
                            pyautogui.typewrite(['tab'])
                            pyautogui.typewrite(order['po_num'])

                            ## DEVICE SELECTION
                            # TODO: Could replace w/ deviceSelection() if duplicate case corrected to allow multiple quantity
                            # moves to device selection screen, also LFE WO number ref field
                            goto_dev_shell_mods_tab(pyautogui)
                            while pyautogui.locateOnScreen(device_tab, confidence=0.95, region=(0, 0, 1920, 1080)) is None:
                                print "*Scanning for Device Tab"
                                time.sleep(1)
                            else:
                                print '*Device Tab Detected'
                            # In LFE WO Ref field use target workorder number
                            pyautogui.typewrite(order['wo_num'])
                            pyautogui.typewrite(['down'])
                            pyautogui.typewrite('duplicate')
                            # reset location on page to WO REF LFE
                            goto_dev_shell_mods_tab(pyautogui)
                            # Moves to and preps Device Code Field
                            pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
                            # Use stored device code from processing
                            pyautogui.typewrite(order['device_code'])
                            # goto_dev_shell_mods_tab(pyautogui)
                            # pyautogui.typewrite(['down', 'down', 'down', 'down', 'down', 'down'])
                            # q_plus = str(q_plus)
                            # pyautogui.typewrite(q_plus)
                            # pyautogui.typewrite(['right'])
                            # pyautogui.typewrite(order['footedness'])

                            # Confirm selection of device code and wait for it to resolve
                            while pyautogui.locateOnScreen(add_button, confidence=0.95,
                                                           region=(1000, 345, 160, 140)) is None:
                                print 'Adding Device...'
                            else:
                                print "Device Added"
                                pyautogui.click(
                                    pyautogui.locateCenterOnScreen(add_button, confidence=0.95,
                                                                   region=(1000, 345, 160, 140)))
                            # pyautogui.typewrite(['return', 'return'], interval=SHORT_DELAY)
                            count = 1
                            while pyautogui.locateOnScreen(remove_button, confidence=0.95, region=(1070, 378, 109, 60)) is None:
                                print '*Resolving Device Selection', count

                                if count > 10:
                                    if pyautogui.locateOnScreen(deadlock, confidence=0.95,
                                                                region=(644, 387, 572, 320)):
                                        image6 = pyautogui.screenshot('deadlock_{}.png'.format(order['po_num']))
                                        pyautogui.typewrite(['return'])
                                        goto_dev_shell_mods_tab(pyautogui)
                                        pyautogui.typewrite(['down', 'down'])
                                        time.sleep(0.5)
                                        pyautogui.typewrite(order['prev_wo'])
                                        time.sleep(0.5)
                                        pyautogui.typewrite(['down', 'space'])
                                        time.sleep(0.5)
                                        # for now will act like all dupe orders are dupes
                                        pyautogui.typewrite(['tab', 'tab', 'tab', 'tab', 'tab', 'return'],
                                                            interval=0.10)

                                    if pyautogui.locateOnScreen(interrupt, confidence=0.95,
                                                                region=(750, 436, 420, 225)):
                                        image6 = pyautogui.screenshot('deadlock_{}.png'.format(order['po_num']))
                                        pyautogui.typewrite(['return'])
                                        goto_dev_shell_mods_tab(pyautogui)
                                        pyautogui.typewrite(['down', 'down'])
                                        time.sleep(0.5)
                                        pyautogui.typewrite(order['prev_wo'])
                                        time.sleep(0.5)
                                        pyautogui.typewrite(['down', 'space'])
                                        time.sleep(0.5)
                                        # for now will act like all dupe orders are dupes
                                        pyautogui.typewrite(['tab', 'tab', 'tab', 'tab', 'tab', 'return'],
                                                            interval=0.10)
                                time.sleep(4.5)
                                count += 1
                            else:
                                print '**Device Selection Resolved'
                                if order['quantity_num'] > 1:
                                    order['issue_list'].append('Sub Quantity')
                                    goto_dev_shell_mods_tab(pyautogui)
                                    pyautogui.typewrite(['down', 'down', 'down', 'down', 'down', 'down'])
                                    pyautogui.typewrite(order['quantity'])
                            setOrderPriority(order)

                            # # return to main screen, account number field
                            # goto_general_order_info_tab(pyautogui)
                            # time.sleep(8)
                            # pyautogui.typewrite(['right'])
                            #
                            # ## RUSH OR ONTIME SPECIFIED
                            # #TODO code to detect printer is out of stickers
                            # # Move to priority field
                            # pyautogui.typewrite(['tab', 'tab', 'tab'])
                            # if order['priority'] == 'RRU On Time':
                            #     pyautogui.typewrite(['s', 'p', 'return', 'up', 'right', 'space'])
                            #     time.sleep(5)
                            # if order['priority'] == '3day Rush':
                            #     pyautogui.typewrite(['3', 'return', 'return', 'up', 'right', 'space'])
                            #     time.sleep(5)

                            # Printing out the Sitcker Label
                            # LFE orders need to be done by hand
                            if sticker_gate == 'yes':
                                if order['data_base'] is not 'LFE':
                                    pyautogui.hotkey('alt', 'e')
                                    time.sleep(2)
                                    marker = 1
                                    while pyautogui.locateOnScreen(after_print, confidence=0.95, region=(0, 0, 1920, 1080)) is None:
                                        print "*Scanning for Label Print Completion #", marker
                                        time.sleep(0.5)
                                    else:
                                        print '**Label Printed'
                                        pyautogui.typewrite(['esc'])

                # Marking the Time of end of order
                    orderend_time = datetime.datetime.now()
                    orderstart_time.strftime("%H:%M")
                    order_duration = orderend_time - orderstart_time
                    print 'Elapsed time during order:', order_duration

                    #end of order clean reset

                    #pyautogui.press(['esc', 'esc', 'enter', 'esc', 'esc', 'enter'], interval=0.25)
# Marks the time of the end of the whole OE process
process_end = datetime.datetime.now()

def printOrdersOfInterest(orders):
    """
    Prints orders with issues out to standard out in a nice format.
    """

    header = """
==========================
FLAGGED ORDERS OF INTEREST
==========================
"""
    print header
    out_fh.write(header)

    n = 1
    for order in orders:
        if order['sub_order'] in ('changed', 'duplicate'):
            sub_marker = 'SUB'
        else:
            sub_marker = ''
        if len(order['issue_list']) >= 1:
            if order['nameLast'] != '':
                #print order['wo_num']
                line = " ".join(map(str,[
                    '[PO#', order['po_num'] + ']', '[WO#', order['wo_num'] + ']', order['FIRSTNAMECOMPLETE'], order['nameLast'],
                     order['issue_list'], "\n\n"
                ]))
                print '[PO#', order['po_num'] + ']','[WO#', order['wo_num'] + ']', order['FIRSTNAMECOMPLETE'], order['nameLast'], \
                     order['issue_list']

                out_fh.write(line)
                print ''
                out_fh.write("")
        n += 1
time_difference = process_end - process_start
timestamp = " ".join(map(str, ['', 'Start Time:', process_start.strftime("%H:%M %m-%d-%Y"), "\n", 'End Time:',
                              process_end.strftime("%H:%M %m-%d-%Y"), "\n", 'Time Elapsed:', time_difference, "\n",
                               'New Orders:', NEWORDER_COUNT, 'Sub Orders:', SUBORDER_COUNT, '/', LFE_COUNT]))
out_fh.write(timestamp)
# Printing out all order of interest
printOrdersOfInterest(orders)

def testfileRenamer():

    for key in cleanfile_to_rawfile:
        print key

    print "Testing testfileRenamer:"
    order1 = {}
    order1['po_num'] = '108092'
    order1['foot2scan'] = 'RIGHT'
    order1['wo_num'] = '812345'
    order1['sub_order'] = 'new order'
    order1['issue_list'] = []
    fileRenamer(order1)
    #print order1

    order2 = {}
    order2['po_num'] = '108093'
    order2['foot2scan'] = 'LEFT'
    order2['wo_num'] = '922345'
    order2['sub_order'] = 'new order'
    order2['issue_list'] = []
    fileRenamer(order2)
    #print order2

    order3 = {}
    order3['po_num'] = '108098'
    order3['foot2scan'] = 'BOTH'
    order3['wo_num'] = '32345'
    order3['sub_order'] = 'new order'
    order3['issue_list'] = []
    fileRenamer(order3)
    #print order3
    # missing
    order4 = {}
    order4['po_num'] = '188888'
    order4['foot2scan'] = 'RIGHT'
    order4['wo_num'] = '42345'
    order4['sub_order'] = 'new order'
    order4['issue_list'] = []
    fileRenamer(order4)
    #print order4

    # duplicate
    order5 = {}
    order5['po_num'] = '108113'
    order5['foot2scan'] = 'RIGHT'
    order5['wo_num'] = '52345'
    order5['sub_order'] = 'new order'
    order5['issue_list'] = []
    fileRenamer(order5)
    #print order5

#testfileRenamer()

#Look for original_file(s) in raw_files
# If cant be found log error and skip # Error 2: PO raw file not found
#Rename old to new files
#Create full paths old and new C:\\where\I\need\to\be\{po}L.raw, C:\\where\I\need\to\be\{wo}.raw
#os.rename(old, new)
#Rename a raw file >:B no carets




# # Creates expected list of raw file names from po num and foot to scan
# expected_raw = []
# for order in orders:
#     if order['sub_order'] == 'new order':
#         expected_raw.append(order['po_num'] + order['foot2scan'] + '.RAW')
#         renamed_raw = order['wo_num'], order['foot2scan']
#         print renamed_raw
#
# print expected_raw
#
#
#
# # Iterate through po numbers and foot to scan to check against the names found in the raw_files directory
#
# for rawfile in rawfiles:
#     print rawfile
#     for order in orders:
#         order['raw_found'] = 'no'
#         if rawfile == order['old_path']:
#             order['raw_found'] = 'yes'
#             os.rename(order['old_path'], order['new_path'])




# Printing out info on how long it took and how look it is estimated to take
print ''
print '________________________________'
print ' ', 'Start Time:', process_start.strftime("%H:%M %m-%d-%Y")
print ' ', 'End Time:', process_end.strftime("%H:%M %m-%d-%Y")
time_difference = process_end - process_start
print ' ', 'Time Elapsed:', time_difference
print '________________________________________'
print ' ', 'Number of New Orders =', NEWORDER_COUNT, '(', (NEWORDER_COUNT * 98)/60, 'mins', ')'
print ' ', 'Number of Sub orders =', SUBORDER_COUNT, '(', ((SUBORDER_COUNT - LFE_COUNT) * 45)/60, 'mins', ')'
print ' ', 'Number of LFE =', LFE_COUNT
print '________________________________________'
print ''
print '                            <[ [ [PROGRAM COMPLETE] ] ]>'
print ''

print "Processed", len(orders), "orders..."
out_fh.close()