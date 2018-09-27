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
#import opencv


# Importing Reference Images
brandon_target = Image.open('brandon_target.bmp')
add_button = Image.open('add_button.bmp')
inside_pt_card = Image.open('inside_pt_card.bmp')
posted_sales_button = Image.open('posted_button.bmp')
pt_popup = Image.open('pt_popup.bmp')
#pt_search_entry = Image.open('pt_search_entry.bmp')
remove_button = Image.open('remove_button.bmp')
similar_ptname = Image.open('similar_ptname.bmp')
sub_search_confirm = Image.open('sub_search_confirm.bmp')

count = 1
while pyautogui.locateOnScreen(brandon_target, region=(0, 0, 400, 400)) is None:
    print '*Looking for Brandon', count
    count += 1
    time.sleep(2)
    if count == 20:
        break
else:
    print '**Brandon Detected'




exit()
# can i do this such that, you can just place things in the brackets in the right order and they will take them
def imageRecognitionDelay(image2search, left, top, width, height):
    while pyautogui.locateOnScreen(image2search, region=(left, top, width, height)) is None:
        time.sleep(1)
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

# print pyautogui.locateCenterOnScreen(brandon_target)
# exit()
# x = 1
# found = 'no'
# while found == 'no':
#     time.sleep(1)
#     print "searching for Brandon", x
#     if pyautogui.locateCenterOnScreen(brandon_target, region=(0, 0, 400, 400)) != None:
#         found = 'yes'
#         print 'Hi Brandon'
#     if pyautogui.locateCenterOnScreen(brandon_target, region=(0, 0, 400, 400)) == None:
#         print 'not found'
#     x = x + 1
#     if x == 10:
#         found = 'yes'
#
# print 'finished'
# exit()



# myImage = Image.open('hello.PNG')
# print '======= START ======='
# myImage.filename
# pyautogui.locateOnScreen(myImage)
# print '=======  END  ======='

deviceDictionary = {
   #"FUNCTIONAL STANDARD":"STANDARD FUNCTIONAL", "EVA":"EVA", "SPORT STANDARD - NEOPRENE TO TOES":"43172ST01"

"FIREFLY NHS FUNCTIONAL":"431750601", "FIREFLY NHS DRESS":"431750611", "FIREFLY NHS SPORT":"431750621",
"FIREFLY SOCCER SPORT":"431750626", "FIREFLY SOCCER SPORT (DM)":"431750627", "FIREFLY SPORT IMPACT":"431750622",
"FUNCTIONAL STANDARD":"43170ST01", "FUNCTIONAL DIRECT MILLED":"43170DM01", "STANDARD SLIMLINE":"43171LA01",
"LOW HEEL CUP SLIMLINE":"43171LA11", "FLAT HEEL CUP":"43171LA21", "COBRA":"43171LA31", "MENS DRESS":"43171ME01",
"SPORT STANDARD - NEOPRENE TO TOES":"43172ST01", "SPORT DIRECT MILLED - NEOPRENE TO TOES":"43172DM01",
"SPORT DIRECT MILLED - VINYL TO METS":"43172DM02", "SPORT LOW PROFILE":"43172LP01", "SPORT SKI - ALPINE":"43172SI01",
"SPORT SKI - NORDIC":"43172SI02", "SPORT SKI - SNOWBOARD":"43172SI03", "SPORT SKATE - HOCKEY":"43172SA01",
"SPORT SKATE - FIGURE":"43172SA02", "MOLD STANDARD":"43173ST01", "MOLD LOW PROFILE":"43173LP11",
"FIREFLY DIABETIC TRIDENSITY":"431750671", "FIREFLY RA FLEXIBLE MOLD":"431750681", "EVA":"43174EV01",
"UCBL":"43174UC01", "UCBL CHILDREN":"43174UC02", "ROBERTS WHITMAN":"43174RB01", "ROBERTS WHITMAN CHILDREN":"43174RB02", "GAIT PLATE - INDUCE OUT-TOEING":"43174GP02",
"GAIT PLATE - INDUCE IN-TOEING":"43174GP01", "":"43170ST01"
}
subDeviceDictionary = {

"FNHS FUNCTIONAL":"431750601", "FNHS DRESS":"431750611", "FNHS SPORT":"431750621",
"FIREFLY SOCCER SPORT":"431750626", "FIREFLY SOCCER SPORT (DM)":"431750627", "FIREFLY SPORT IMPACT":"431750622",
"FUNCTIONAL STANDARD":"43170ST01", "FUNCTIONAL DIRECT MILLED":"43170DM01", "LADIES DRESS STANDARD SLIMLINE":"43171LA01",
"LADIES DRESS LOW HEEL CUP SLIMLINE":"43171LA11", "LADIES DRESS FLAT HEEL CUP":"43171LA21", "LADIES DRESS COBRA":"43171LA31", "MENS DRESS":"43171ME01",
"SPORT STANDARD NEOPRENE TO TOES":"43172ST01", "SPORT DIRECT MILLED NEOPRENE TO TOES":"43172DM01",
"SPORT DIRECT MILLED VINYL TO METS":"43172DM02", "SPORT LOW PROFILE":"43172LP01", "SPORT SKI ALPINE":"43172SI01",
"SPORT SKI NORDIC":"43172SI02", "SPORT SKI SNOWBOARD":"43172SI03", "SPORT SKATE HOCKEY":"43172SA01",
"SPORT SKATE FIGURE":"43172SA02", "MOLD STANDARD":"43173ST01", "MOLD LOW PROFILE":"43173LP11",
"FIREFLY DIABETIC TRIDENSITY":"431750671", "FIREFLY RA FLEXIBLE MOLD":"431750681", "EVA":"43174EV01",
"UCBL ADULT":"43174UC01", "UCBL CHILDREN":"43174UC02", "ROBERTS WHITMAN ADULT":"43174RB01", "ROBERTS WHITMAN CHILDREN":"43174RB02", "GAIT PLATE INDUCE OUT-TOEING C HILDREN":"43174GP02",
"GAIT PLATE INDUCE IN-TOEING CH ILDREN":"43174GP01", "":"43170ST01"
}
#print deviceDictionary["SPORT STANDARD - NEOPRENE TO TOES"]
#print deviceDictionary


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

    DATABASE_PO_BORDER = 97629
    # Clean raw value

    prev_po = cleanAndReturnNumberString(raw_prev_po)
    prev_po = re.sub("\.0$", "", prev_po) # removes tailing .0
    curr_order['prev_po'] = prev_po

    if prev_po == "":
        # Not relevant
        curr_order["data_base"] = ""
        return curr_order

    # Determine source
    # TODO Assumed that LFE was on the border
    if int(prev_po) <= DATABASE_PO_BORDER:
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
    print actual == expected

    raw_prev_po2 = "96500.0"
    curr_order = {}
    actual = previousOrderSource(raw_prev_po2, curr_order)
    expected = {'data_base': 'LFE', 'prev_po': '96500'}
    print actual == expected

    raw_prev_po2 = ""
    curr_order = {}
    actual = previousOrderSource(raw_prev_po2, curr_order)
    expected = {'data_base': '', 'prev_po': ''}
    print actual == expected

    raw_prev_po2 = 1000000
    curr_order = {}
    actual = previousOrderSource(raw_prev_po2, curr_order)
    print actual

def cleanAndReturnNumberString(raw_num):
    raw_num = str(raw_num).strip() # removes whitespace
    raw_num = re.sub("\.0$", "", raw_num) # removes tailing .0
    return raw_num

def testCleanAndReturnNumberString():
    test1 = "10000.0"
    actual = cleanAndReturnNumberString(test1)
    expected = "10000"
    print actual == expected

# Run all the tests
print "Run all tests:"
testPreviousOrderSource()
testCleanAndReturnNumberString()


###### PHASE 1: PROCESSING EXCEL ORDERS

for f in files:
   if f.endswith("xls"):
       xlsfiles.append(f)

orders = [] #########order processing############
for xfile in xlsfiles:
   filepath = os.path.join(targetdir, xfile)
   #dowork
   xl_workbook = xlrd.open_workbook(filepath)
   xl_sheet = xl_workbook.sheet_by_index(1)
   print ('Sheet name: %s' % xl_sheet.name)
   #print xlrd.xldate_as_datetime(42088,xl_workbook.datemode)

   curr_order = []
   for order in range(xl_sheet.nrows):
       row = xl_sheet.row(order)

######## PO_NUMBER , PARSING NAME
       if row[0].value=='PATIENT NAME / CODE NO.':
          #print '===================================='
           curr_order = {}  ################ Creation of the dictionary [why here, why not higher]
           curr_order['wo_num'] = ''
           curr_order['pt_num'] = ''
           curr_order['po_num'] = ''
           curr_order['issue_list'] = []
           curr_order['po_num'] = cleanAndReturnNumberString(row[-1].value)
           curr_order['name'] = row[1].value.strip()
           #finding all character strings, hyphen and apostraphy included
           nameList = (re.findall("[a-zA-Z'-]+", row[1].value.strip()))
           #finding if there is a pt serial number
           nameNumber = (re.findall("[\d]+",row[1].value.strip()))
           #print nameList
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

           if len(nameNumber) == 1:
               curr_order['nameNumber'] = nameNumber[0]

           curr_order['FIRSTNAMECOMPLETE'] = curr_order['firstname'] + namespace + curr_order['nameNumber']



######## WEIGHT, FOOT SIZE, PRIORITY
       if row[0].value=='WEIGHT RANGE / SIZE OF FOOT / PRIORITY / TEMPLATE':
           curr_order['weight'] = ''
           curr_order['shoesize'] = ''
           weightLen = (re.findall('[\d]+',row[1].value))
           if len(weightLen)==4:
               curr_order['weight'] = weightLen[1]
           if len(weightLen)==2:
               curr_order['weight'] = weightLen[0]
           # print row[4]
           # if row[4].value == '':
           #     tesla = ''
           # else:
           #     tesla = re.findall('[.\d]+', row[4].value)
           #
           # print tesla
           #curr_order['shoesize'] = tesla
           curr_order['shoesize'] = re.findall('[.\d]+', row[4].value)
           curr_order['priority'] = row[6].value
           if curr_order['weight'] == '':
               curr_order['issue_list'].append('weight missing')
           if len(curr_order['shoesize']) < 1:
               curr_order['issue_list'].append('shoesize missing')
           # pyautogui.click(100, 100)
           # pyautogui.typewrite(curr_order['shoesize'])
           # pyautogui.typewrite(['return'])

######## QUANTITY, SUB_ORDER STATUS, SAME DAY SUB ORDER, PREV_PO
       if row[0].value=='QUANTITY / SUBSEQUENT PAIR':
           curr_order['quantity'] = row[1].value
           raw_prev_po = row[4].value
           curr_order = previousOrderSource(raw_prev_po, curr_order)

           sub_order = row[7].value
           if sub_order == '':
               sub_order = 'new order'
           elif sub_order == 'CHANGED DEVICE (Select device and options)':
               sub_order = 'changed'
               if curr_order['prev_po'] == '':
                   curr_order['issue_list'].append('missing prev_po')
           elif sub_order == 'DUPLICATE DEVICE (No change)':
               sub_order = 'duplicate'
               if curr_order['prev_po'] == '':
                   curr_order['issue_list'].append('missing prev_po')

           curr_order['sub_order'] = sub_order
           #print curr_order['sub_order']
           b = 0
           curr_order['sameday_suborder'] = 'no'
           curr_order['counter'] = b
           curr_order['suborder_target'] = None
           for a in orders:

               curr_order['counter'] = b
               b = b + 1
               if curr_order['prev_po'] == a['po_num']:
                   curr_order['sameday_suborder'] = 'yes'
                   curr_order['issue_list'].append('sameday suborder')
                   curr_order['suborder_target'] = curr_order['counter']
                   curr_order['sub_order'] = 'new order' # added as exception when sub orders weren't possible

######## OUTGROWTH STATUS, BIRTHDAY
       if row[0].value=='OUTGROWTH PAIR / DOB':
           curr_order['outgrowth'] = row[1].value.strip()
           match = re.search(r'pair', curr_order['outgrowth'])
           #ogConfirm = (re.findall("[a-zA-Z]"))
           if match:
               curr_order['outgrowth'] = 'yes'
           else :
               curr_order['outgrowth'] = 'no'
           curr_order['dob'] = ''
           if row[7].ctype == 3:
               #curr_order['dob'] = xlrd.xldate_as_datetime(int(row[7].value),xl_workbook.datemode).strftime("%m%d%y")
               print row
           else:
               curr_order['dob']= ""
######## DEVICE CODE
       if row[0].value=='DEVICE':
           curr_order['device'] = str(row[1].value).strip().strip("\s")
           poro = curr_order['device']
           device_code = deviceDictionary[poro]
           #print poro
           curr_order['device_code'] = device_code
           if curr_order['sub_order'] == 'new order':
               if curr_order['device'] == '':
                   curr_order['issue_list'].append('device missing')


####### FOOT TO SCAN
       if row[0].value=='FOOT SCANNED':
           curr_order['both'] = row[1].value
           curr_order['left'] = row[4].value
           curr_order['right'] = row[7].value

           if (not curr_order['both'] and not curr_order['right'] and not curr_order['left']):
               curr_order['foot2scan'] = 'SUBSEQUENT PAIR ORDER'
           elif curr_order['both'] and (not curr_order['right'] and not curr_order['left']) \
                   or (curr_order["right"] and curr_order["left"] and not curr_order["both"] ) :
               curr_order['foot2scan'] = 'BOTH'
           elif curr_order['right'] and not curr_order['left'] and not curr_order['both']:
               curr_order['foot2scan'] = 'RIGHT'
           elif curr_order['left'] and not curr_order['right'] and not curr_order['both']:
               curr_order['foot2scan'] = 'LEFT'
           else:
               curr_order['foot2scan'] = 'Human Help!'
               if curr_order['sub_order'] == 'new order':
                   curr_order['issue_list'].append('foot to scan')

######## SPECIAL INSTRUCTIONS
       if row[0].value=='NOTES':
           nextrow = xl_sheet.row(order + 1)
           curr_order['notes'] = nextrow[0].value.upper()
           special_instructions = (re.findall("[a-zA-Z'-]+", curr_order["notes"]))
           for word in special_instructions:
               if word == 'HOLD':
                   curr_order['issue_list'].append('hold request')
               if word == 'ADDRESS':
                   curr_order['issue_list'].append('alternate address')
               if word == 'SHIP':
                   curr_order['issue_list'].append('alternate address')
               if word == 'PAPERWORK':
                   curr_order['issue_list'].append('alternate address')
               if word == 'UPS':
                   curr_order['issue_list'].append('alternate address')
           #curr_order = detectHoldRequest(curr_order)

           orders.append(curr_order)


##**************************************ALL ORDERS HAVE BEEN PROCESSED*******************************************

### PHASE 2: AUTO-DYLAN

#================================= CONTROL PANEL ========================================

pyautogui.click(1046,743)
#pyautogui.click(100,100)
ORDER_LIMIT =100 #how many orders from start you want
TIME_DELAY = 1.5 #delay between most actions default: 1.5
SHORT_DELAY = 0.25 #delay between quick actions default: 0.25
LAG = 23 #delay for in and out of pt search default: 40


#=========================================================================================

def printOrderIntoCard(order, k):
    if order['po_num'] != '':
         if order['sub_order'] == 'new order' or 'duplicate' or 'changed':
            print '\n', '\n', '\n'
            print '==============================================================================='
            print '                          ','ORDER', k, ': Position', k-1, '\n'
            print ((k+1) * 210)/60
            print 'This is a', order['sub_order']+ '\n' + 'NAME ON ORDER =', order['name'] + '\n' + 'FIRST NAME FIELD =', order['FIRSTNAMECOMPLETE']
            print 'LAST NAME FIELD =', order['nameLast'] + '\n' + 'DOB =', order['dob']+ '\n' + 'OUTGROWTH =', order['outgrowth']
            print 'WEIGHT =', order['weight']
            print 'SHOE SIZE =', order['shoesize']
            print 'PO NUMBER =', order['po_num']
            print 'PREVIOUS PO# =', order['data_base'], order['prev_po']
            print 'FOOT TO SCAN =', order['foot2scan'] + '\n' + 'PRIORITY =', order['priority'] + '\n' + 'QUANTITY =', order['quantity']
            print 'NOTES =', order['notes'] + '\n' + 'DEVICE =', order['device'] + '\n' + 'D.CODE =', order['device_code']
            print 'SAMEDAY SUBORDER =', order['sameday_suborder'] + '\n' + 'ORDER POSITION UP TO NOW =', order['counter']
            print 'SUBORDER TARGET =', order['suborder_target']
            print 'NOTES =', order['notes']


def fetchSubsequentOrder(order): #bannana grabber tm
    #TODO Add device grabber
    pyautogui.typewrite(['down', 'return'])
    # todo image find posted sales search pyauto click
    pyautogui.click(pyautogui.locateCenterOnScreen(posted_sales_button, region=(1264, 164, 150, 56)))
    # Select with in search result field
    pyautogui.click(850, 450)
    # Move to the most left, bringing you to po_num column
    pyautogui.typewrite(['home'])
    pyautogui.typewrite(order['prev_po'])
    pyautogui.typewrite(['return'])
    time.sleep(1)
    while pyautogui.locateCenterOnScreen(sub_search_confirm) == None:
        time.sleep(2)

    order['issue_list'].append('prev_ponum search fail')
    pyautogui.typewrite(['right'])
    pyautogui.hotkey(['ctrl', 'c'])
    order['prev_wo'] = (Tk().clipboard_get())
    pyautogui.typewrite(['right'])
    pyautogui.hotkey(['ctrl', 'c'])
    order['prev_ptnum'] = (Tk().clipboard_get())
    pyautogui.typewrite(['right'])
    pyautogui.hotkey(['ctrl', 'c'])
    order['prev_device'] = (Tk().clipboard_get())
    order['prev_device'] = subDeviceDictionary[order['prev_device']]
    if order['device_code'] == order['prev_device']:
        order['alter_subdevice'] = 'no'
    else:
        order['alter_subdevice'] = 'yes'

def orderCreation(order):
    ## Open order
    # Selecting 'Paris Sales Order'
    #TODO image match general screen
    pyautogui.typewrite(['return'])
    time.sleep(4)
    # Create new order
    pyautogui.press(['f3'])
    time.sleep(5)
    # Move to work order number field
    pyautogui.typewrite(['tab', 'up', 'up', 'up', 'up', 'up'], interval=SHORT_DELAY)
    time.sleep(1)
    # Copy work order number to clipboard
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)

    # Get created workorder number from clipboard
    # TODO: Maybe don't need ()
    order['wo_num'] = (Tk().clipboard_get())
    print 'wo_num', order['wo_num']

    ## Select firefly as account
    # Select account number field
    pyautogui.hotkey('alt', 'g')
    time.sleep(TIME_DELAY)
    # Enter firefly account number
    pyautogui.press(['0', '7', '4', '5', 'return'])
    time.sleep(TIME_DELAY)

    ## SELECT CLINICIAN (MAIN)
    # Opens clinician search screen
    pyautogui.typewrite('2172')
    time.sleep(TIME_DELAY)

    # Confirms selection and moves to patient card search
    pyautogui.typewrite(['return'], interval=TIME_DELAY)

def patientCardHandler(order):
    if order['sameday_suborder'] == 'yes':
        pyautogui.typewrite(orders[order['suborder_target']]['pt_num'])
    else:
        if order['sub_order'] in ('duplicate', 'changed'):
            pyautogui.typewrite(order['prev_ptnum'])
            pyautogui.typewrite(['enter'])
        if order['sub_order'] == 'new order':
            pyautogui.typewrite(['f6'])
            time.sleep(LAG)  # ENTERING PT SEARCH, LONG DELAY
            # CREATING NEW CARD VIA SEARCHING
            # Checking to see if opening search screen has resolved

            count = 1
            # TODO needs region to search
            while pyautogui.locateOnScreen(pt_search_entry) is None:
                print '*Pt Search Screen Scan #', count, '-Complete-'
                count += 1
                time.sleep(2)
                if count == 8:
                    order['issue_list'].append('Opening Pt Search')

            else:
                print '**Search Screen -Detected-'


            pyautogui.typewrite([order['FIRSTNAMECOMPLETE'], 'tab', order['nameLast']])
            # pyautogui.typewrite(['tab'])
            # pyautogui.typewrite(i["nameLast"])
            pyautogui.hotkey('alt', 's')
            pyautogui.hotkey('alt', 'c')
            pyautogui.hotkey('alt', 's')
            # SEARCHING FOR PT CREATES LONG DELAY
            time.sleep(LAG)

            x = 1
            found = 'no'
            while found == 'no':
                print '*New Pt Popup Search #', x
                x = x + 1
                time.sleep(2)
                # To see New Pt Pop Up Dialogue box
                if pyautogui.locateOnScreen(pt_popup, region=(803, 457, 321, 186)) is not None:
                    print '**Popup Detected'
                    found = 'yes'
                    pyautogui.typewrite(['y'])
                    time.sleep(1)
                    # entering pt info
                    pyautogui.typewrite(['tab', 'tab', 'm', 'tab'])  # gender
                    time.sleep(SHORT_DELAY)
                    pyautogui.typewrite(order["dob"])
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite(order['weight'])
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

                    pyautogui.typewrite(['esc', 'return'], interval=SHORT_DELAY)
                    pyautogui.hotkey('ctrl', 'c')
                    order['pt_num'] = Tk().clipboard_get()
                    print 'pt_num =', order['pt_num']
                    time.sleep(TIME_DELAY)
                if x == 5:
                    print '*Similar Name Search'
                    #TODO needs region to search
                    if pyautogui.locateOnScreen(similar_ptname) is not None:
                        found = 'yes'
                        print '**Similar Name Assumed'
                        order['issue_list'].append('Similar Pt Name Issue')
                        print order['issue_list']
                        pyautogui.hotkey('shift', 'f5')
                        pyautogui.typewrite(['f3', 'tab'])
                        pyautogui.typewrite(order['FIRSTNAMECOMPLETE'])
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(order['nameLast'])
                        pyautogui.typewrite(['tab', 'm', 'tab'])
                        pyautogui.typewrite(order["dob"])
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(order['weight'])
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
    pyautogui.typewrite(['tab'])
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
        if order['sub_order'] == 'new order':
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

def deviceSelection(order):
    pyautogui.hotkey('alt', 'm')
    time.sleep(1.5)
    if order['sameday_suborder'] == 'yes':
        pyautogui.typewrite(orders[order['suborder_target']]['wo_num'])
        pyautogui.typewrite(['down'])
        pyautogui.typewrite('changed device')
        pyautogui.hotkey('alt', 'm')
        pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
        pyautogui.typewrite(order['device_code'])
        pyautogui.typewrite(['return', 'return'], interval=SHORT_DELAY)
        while pyautogui.locateOnScreen(remove_button, region=(0, 0, 0, 0)) is None:
            time.sleep(1)
    else:
        if order['sub_order'] == 'new order':
            pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
            pyautogui.typewrite(order['device_code'])
            pyautogui.typewrite(['return', 'return'], interval=SHORT_DELAY)
            while pyautogui.locateOnScreen(remove_button, region=(0, 0, 0, 0)) is None:
                time.sleep(1)

        if order['sub_order'] == 'changed':
            if order['data_base'] == 'LFE':
                pyautogui.typewrite(['down'])
                pyautogui.typewrite('changed device')
                pyautogui.hotkey('alt', 'm')
                pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
                if order['device_code'] == '':
                    order['device_code'] = '43170ST01'
                pyautogui.typewrite(order['device_code'])
                pyautogui.typewrite(['return', 'return'], interval=SHORT_DELAY)
                while pyautogui.locateOnScreen(remove_button, region=(0, 0, 0, 0)) is None:
                    time.sleep(1)
            if order['data_base'] == 'AM':
                pyautogui.typewrite(['down', 'down'])
                pyautogui.typewrite(['prev_wo'])
                pyautogui.typewrite(['down', 'down', 'space'])
                if order['alter_subdevice'] == 'yes':
                    pyautogui.typewrite(['space', 'tab', 'tab', 'tab'])
                    if order['device_code'] == '':
                        pyautogui.typewrite(order['prev_device'])
                    else:
                        pyautogui.typewrite(order['device_code'])
                if order['alter_subdevice'] == 'no':
                    pyautogui.typewrite(['return', 'return', 'return'])

                pyautogui.typewrite(['return', 'return'])

            while pyautogui.locateOnScreen(remove_button, region=(0, 0, 0, 0)) is None:
                time.sleep(1)


        # TODO: Review code here for correctness, add quanity exception and normal dupe functionality
        # TODO: AM LFE differenciation
        if order['sub_order'] == 'duplicate':
            if order['data_base'] == 'LFE':
                pyautogui.typewrite(['down'])
                pyautogui.typewrite('duplicate')
                pyautogui.hotkey('alt', 'm')
                pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
                if order['device_code'] == '':
                    order['device_code'] = '43170ST01'
                pyautogui.typewrite(order['device_code'])
                pyautogui.typewrite(['return', 'return'], interval=SHORT_DELAY)
                time.sleep(DEVICE_SELECTION_DELAY)
            if order['data_base'] == 'AM':
                pyautogui.typewrite(['down', 'down'])
                pyautogui.typewrite(['prev_wo'])
                pyautogui.typewrite(['down', 'space'])
                # for now will act like all dupe orders are dupes
                pyautogui.typewrite(['tab', 'tab', 'tab', 'tab', 'tab', 'return'])
                #TODO - need to make code to determin when a dupe order is actually a changed order
                #TODO - code will be left here as naive until more developement done
                # if order['device_code'] == '':
                #     pyautogui.typewrite(['tab', 'tab', 'tab', 'tab','tab','return'])
                # if order['alter_subdevice'] == 'yes':
                #     pyautogui.typewrite(['space', 'tab', 'tab', 'tab', 'tab'])
                #     if order['device_code'] == '':
                #         pyautogui.typewrite(order['prev_device'])
                #     else:
                #         pyautogui.typewrite(order['device_code'])
                # if order['alter_subdevice'] == 'no':
                #     pyautogui.typewrite(['return', 'return', 'return'])
                #
                # pyautogui.typewrite(['return', 'return'])
            time.sleep(DEVICE_SELECTION_DELAY)


def setOrderPriority(order):
    if order['priority'] == 'RRU On Time':
        pyautogui.typewrite(['s', 'p', 'return', 'up', 'right', 'space'])
        time.sleep(5)
    if order['priority'] == '3day Rush':
        pyautogui.typewrite(['3', 'return', 'return', 'up', 'right', 'space'])
        time.sleep(5)

# The excecution of pyautogui entering in orders
k=0
for i in range(len(orders)):
    order = orders[i]
    k = i+1
    printOrderIntoCard(order, k) # Prints order out in formatted way
    ##  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<   CONSOLE OUTPUT ONLY    >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>##
    if order['po_num'] != '': # doesn't make orders for blank orders
        if order['sub_order'] == 'new order':
            if k <= ORDER_LIMIT:
                # start new order and collect wo#
                if order['sub_order'] in ('duplicate', 'changed'): # only for am to am
                    fetchSubsequentOrder(order)
                orderCreation(order)
                # CREATE PATIENT CARD VIA SEARCH (MAIN)
                patientCardHandler(order)
                # IMPRESSION TYPE, FOOT TO SCAN (MAIN)
                footImpressionAndPONumberEntry(order, orders)
                # DEVICE SELECTION (MAIN)
                deviceSelection(order)
                # if error:
                #     continue
                # Return to main screen
                pyautogui.hotkey('alt', 'g')
                time.sleep(2)
                pyautogui.typewrite(['right'])

                # Deleting impression if it was autofilled
                if order['sub_order'] in ('duplicate', 'changed'):
                    pyautogui.typewrite(['delete'])

                # Move from foot impression type to shipment priority
                pyautogui.typewrite(['tab', 'tab', 'tab'])
                setOrderPriority(order)

                # Order Quanity 2 or more
                q_plus = order['quantity']
                while q_plus - 1 > 0:
                    q_plus = q_plus - 1
                    orderCreation(order)

                    #pick pt card by saved pt_num number
                    pyautogui.typewrite(order['pt_num'])
                    time.sleep(1.5)

                    # TODO: Check if footImpression function works here
                    # See if special case with quanity=True matches the behavior of the commented out code
                    #footImpressionAndPONumberEntry(order, orders, quantity=True)
                    # # Impression type, foot to scan
                    pyautogui.typewrite(['tab', 'delete', 'tab', 'tab'], interval=SHORT_DELAY)
                    time.sleep(SHORT_DELAY)
                    pyautogui.typewrite(order['foot2scan'])
                    pyautogui.typewrite(['tab'])

                    # PO number
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite(order['po_num'])

                    ## DEVICE SELECTION
                    # TODO: Could be replaced with deviceSelection() if duplicate case corrected to allow multiple quanitty
                    #moves to device slection screen, also workorder number ref feild
                    pyautogui.hotkey('alt', 'm')
                    time.sleep(1)
                    # In Workorder Ref field use target workorder number
                    pyautogui.typewrite(order['wo_num'])
                    pyautogui.typewrite(['down'])
                    pyautogui.typewrite('duplicate')
                    # reset location on page to WO REF LFE
                    pyautogui.hotkey('alt', 'm')
                    # Moves to and preps Device Code Field
                    pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=SHORT_DELAY)
                    # If no device selected default to std fnc
                    if order['device_code'] == '':
                        order['device_code'] = '43170ST01'
                    # Use stored device code from processing
                    pyautogui.typewrite(order['device_code'])
                    # Confirm selection of device code and wait for it to resolve
                    pyautogui.typewrite(['return', 'return'], interval=SHORT_DELAY)
                    time.sleep(12)

                    # return to main screen, account number field
                    pyautogui.hotkey('alt', 'g')
                    time.sleep(8)
                    pyautogui.typewrite(['right'])

                    ## RUSH OR ONTIME SPECIFIED

                    # Move to priority field
                    pyautogui.typewrite(['tab', 'tab', 'tab'])
                    if order['priority'] == 'RRU On Time':
                        pyautogui.typewrite(['s', 'p', 'return', 'up', 'right', 'space'])
                        time.sleep(5)
                    if order['priority'] == '3day Rush':
                        pyautogui.typewrite(['3', 'return', 'return', 'up', 'right', 'space'])
                        time.sleep(5)

                # Order entering Reset
                pyautogui.typewrite(['esc', 'esc', 'esc', 'esc', 'esc', 'up', 'up'], interval=0.3)

                #end of order clean reset

                #pyautogui.press(['esc', 'esc', 'enter', 'esc', 'esc', 'enter'], interval=0.25)


def printOrdersOfInterest(orders):
    """
    Prints orders with issues out to standard out in a nice format.
    """
    print ''
    print ''
    print '=========================='
    print "FLAGGED ORDERS OF INTEREST"
    print '=========================='

    for order in orders:
        if len(order['issue_list']) >= 1:

            """
            
            """
            if order['nameLast'] != '':
                #print i['wo_num']
                print order['FIRSTNAMECOMPLETE'], order['nameLast'], ' :', order['po_num'], order['issue_list']


printOrdersOfInterest(orders)


print ''
print ''
print '                            <[ [ [PROGRAM COMPLETE] ] ]>'
print ''

print "Processed", len(orders), "orders..."