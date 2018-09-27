import xlrd
import json
import os
import re
import datetime
import pyautogui
import time
from Tkinter import Tk
import PIL
from PIL import Image

print 'hello'

catIm = Image.open('test.bmp')
print catIm.size
print catIm.format
location = pyautogui.locateOnScreen(catIm)
print location
here = pyautogui.center(location)

pyautogui.click(here)
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
"UCBL":"43174UC01", "ROBERTS WHITMAN":"43174RB01", "GAIT PLATE - INDUCE OUT-TOEING":"43174GP02",
"GAIT PLATE - INDUCE IN-TOEING":"43174GP01", "":"43170ST01"
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

for f in files:
   if f.endswith("xls"):
       xlsfiles.append(f)

orders = [] #########order processing############
for xfile in xlsfiles:
   filepath = os.path.join(targetdir, xfile)
   #dowork
   xl_workbook = xlrd.open_workbook(filepath)
   xl_sheet = xl_workbook.sheet_by_index(1)
   # print ('Sheet name: %s' % xl_sheet.name)
   # print xlrd.xldate_as_datetime(42088,xl_workbook.datemode)

   curr_order = None
   for i in range(xl_sheet.nrows):
       row = xl_sheet.row(i)

######## PO_NUMBER , PARSING NAME
       if row[0].value=='PATIENT NAME / CODE NO.':
          #print '===================================='
           curr_order = {}  ################ Creation of the dictionary [why here, why not higher]
           curr_order['wo_num'] = ''
           curr_order['pt_num'] = ''
           curr_order['po_num'] = ''
           #print 'hello'
           #print re.findall('[.\d]+', row[9].value)
           #poNumber = int(re.findall("[\d]+", row[9].value.strip()))
           #print poNumber
           #curr_order['po_num'] = str(row[-1].value).rstrip(".0")
           curr_order['po_num'] = str(row[-1].value).strip().strip("\s").rstrip('.0')
           curr_order['name'] = row[1].value.strip()
           #finding all character strings, hyphen and apostraphy included
           nameList = (re.findall("[a-zA-Z'-]+", row[1].value.strip()))
           #finding if there is a pt serial number
           nameNumber = (re.findall("[\d]+",row[1].value.strip()))
           #print nameList
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
           # pyautogui.click(100, 100)
           # pyautogui.typewrite(curr_order['shoesize'])
           # pyautogui.typewrite(['return'])

######## QUANTITY, SUB_ORDER STATUS, SAME DAY SUB ORDER
       if row[0].value=='QUANTITY / SUBSEQUENT PAIR':
           curr_order['quantity'] = row[1].value
           #removing white space, extra numbers, carriage return
           prev_po = str(row[4].value).strip().strip("\s").rstrip(".0")
           curr_order['prev_po'] = prev_po
           sub_order = row[7].value
           if sub_order == '':
               sub_order = 'new order'
           elif sub_order == 'CHANGED DEVICE (Select device and options)':
               sub_order = 'changed'
           elif sub_order == 'DUPLICATE DEVICE (No change)':
               sub_order = 'duplicate'
           else:
               sub_order = 'error'

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
                   curr_order['suborder_target'] = curr_order['counter']
                   curr_order['sub_order'] = 'new order'

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
               curr_order['dob'] = xlrd.xldate_as_datetime(int(row[7].value),xl_workbook.datemode).strftime("%m%d%y")
               #print row
           else:
               curr_order['dob']= ""
######## DEVICE CODE
       if row[0].value=='DEVICE':
           curr_order['device'] = str(row[1].value).strip().strip("\s")
           poro = curr_order['device']
           device_code = deviceDictionary[poro]
           #print poro
           curr_order['device_code'] = device_code

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

######## SPECIAL INSTRUCTIONS
       if row[0].value=='NOTES':
           nextrow = xl_sheet.row(i+1)
           curr_order['notes'] = nextrow[0].value
           orders.append(curr_order)

##**************************************ALL ORDERS HAVE BEEN PROCESSED*******************************************

keyPhrases = {   ### KEY PHRASES TO LOOK FOR IN SPECIAL INSTRUCTIONS
   'HOLDFORNOW':[ u"HOLD",u"FOR",u"NOW"]
}

# def longestSubstringFinder(string1, string2):
#     answer = []
#     len1, len2 = len(string1), len(string2)
#     for i in range(len1):
#         match = []
#         for j in range(len2):
#             if (i + j < len1 and string1[i + j] == string2[j]):
#                 match += string2[j]
#             else:
#                 if (len(match) > len(answer)):
#                     answer = match
#                 match = []
#     return answer
#
# print longestSubstringFinder("apple pie available", "apple pies")
# print longestSubstringFinder("apples", "appleses")
# print longestSubstringFinder("bapples", "cappleses")


#================================= CONTROL PANEL ========================================

#pyautogui.click(1046,743)
#pyautogui.click(100,100)
order_limit =100 #how many orders from start you want
time_delay = 1.5 #delay between most actions default: 1.5
short_delay = 0.25 #delay between quick actions default: 0.25
lag = 70 #delay for in and out of pt search default: 40


#=========================================================================================


#TODO make a report system here for flagging Orders of Interest
k=0
for i in orders:
    if i['po_num'] != '':
         if i['sub_order'] == 'new order8':
            print '\n', '\n', '\n'
            print '==============================================================================='
            print '                          ','ORDER' , k+1, ': Position' , k, '\n'
            print ((k+1) * 210)/60
            # pyautogui.typewrite(i['name'])
            # pyautogui.typewrite(['return'])
            # pyautogui.typewrite(i['shoesize'])
            # pyautogui.typewrite(['return'])
            print 'This is a', i['sub_order']+ '\n' + 'NAME ON ORDER =', i['name'] + '\n' + 'FIRST NAME FIELD =', i['FIRSTNAMECOMPLETE']
            print 'LAST NAME FIELD =', i['nameLast'] + '\n' + 'DOB =', i['dob']+ '\n' + 'OUTGROWTH =', i['outgrowth']
            print 'WEIGHT =', i['weight']
            print 'SHOE SIZE =', i['shoesize']
            print 'PO NUMBER =', i['po_num']
            print 'PREVIOUS PO# =',i['prev_po']
            print 'FOOT TO SCAN =', i['foot2scan'] + '\n' + 'PRIORITY =', i['priority'] + '\n' + 'QUANTITY =', i ['quantity']
            print 'NOTES =', i ['notes'] + '\n' + 'DEVICE =', i['device'] + '\n' + 'D.CODE =', i['device_code']
            print 'SAMEDAY SUBORDER =', i['sameday_suborder'] + '\n' + 'ORDER POSITION UP TO NOW =', i['counter']
            print 'SUBORDER TARGET =', i['suborder_target']
    k = k + 1
  # print('\x1b[5;30;42m' + 'Success!' + '\x1b[0m')


   #tokenizednotes = i["notes"].split(" ")



     ##  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<   CONSOLE OUT PUT ONLY    >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>##
    if i['po_num'] != '':
        if i['sub_order'] == 'new order88':
            if k <= order_limit:
               # start new order and collect wo#
               pyautogui.typewrite(['esc', 'esc', 'esc', 'esc', 'esc'], interval=0.3)
               pyautogui.typewrite(['return'])
               time.sleep(15)
               pyautogui.press(['f3'])
               time.sleep(5)
               pyautogui.typewrite(['tab', 'up', 'up', 'up', 'up', 'up'], interval=0.4)
               time.sleep(1)
               pyautogui.hotkey('ctrl', 'c')
               time.sleep(1)
               wonum_copied = (Tk().clipboard_get())
               i['wo_num'] = wonum_copied
               # select firefly as account
               #i['wo_num'] = str(k*3)
               pyautogui.hotkey('alt', 'g')
               time.sleep(time_delay)
               pyautogui.press(['0', '7', '4', '5', 'return'])
               time.sleep(time_delay)

               #SELECT CLINICIAN (MAIN)
               pyautogui.typewrite(['f6'])
               time.sleep(3)
               pyautogui.typewrite('martin mcgeough')
               time.sleep(time_delay)
               pyautogui.typewrite(['return', 'return', 'return'], interval=time_delay)

               #CREATE PATIENT CARD VIA SEARCH (MAIN)
               if i['sameday_suborder'] == 'yes':
                   pyautogui.typewrite(orders[i['suborder_target']]['pt_num'])

               else:
                   if i['sub_order'] == 'new order':
                       pyautogui.typewrite(['f6'])
                       time.sleep(lag)  # ENTERING PT SEARCH, LONG DELAY
                       # CREATING NEW CARD VIA SEARCHING
                       pyautogui.typewrite([i['FIRSTNAMECOMPLETE'], 'tab', i['nameLast']])
                       # pyautogui.typewrite(['tab'])
                       # pyautogui.typewrite(i["nameLast"])
                       pyautogui.hotkey('alt', 's')
                       pyautogui.hotkey('alt', 'c')
                       pyautogui.hotkey('alt', 's')
                       time.sleep(lag)  # SEARCHING FOR PT CREATES LONG DELAY
                       # NEEDS IMAGE RECOGNITION HERE TO SEE POPUP
                       # if match:
                       #     pyautogui.typewrite(['y'])
                       #     pyautogui.press('tab', 'tab')
                       # else:
                       #     pyautogui.hotkey('shift', 'f5')
                       #     pyautogui.press('f3')
                       #     pyautogui.typewrite(i['FIRSTNAMECOMPLETE'])
                       #     pyautogui.typewrite(['tab'])
                       #     pyautogui.typewrite(i['nameLast'])
                       #     pyautogui.press('tab')
                       pyautogui.typewrite(['y'])
                       time.sleep(1)
                       # entering pt info
                       pyautogui.typewrite(['tab', 'tab', 'm', 'tab'])  # gender
                       time.sleep(short_delay)
                       pyautogui.typewrite(i["dob"])
                       # pyautogui.typewrite(['tab', i['weight'], 'tab', i['shoesize'], 'tab', i['shoesize']],
                       #                     interval=0.25)
                       pyautogui.typewrite(['tab'])
                       pyautogui.typewrite(i['weight'])
                       pyautogui.typewrite(['tab'])
                       pyautogui.typewrite(i['shoesize'])
                       pyautogui.typewrite(['tab'])
                       pyautogui.typewrite(i['shoesize'])


                       time.sleep(time_delay)
                       if i['outgrowth'] == 'yes':
                           pyautogui.typewrite(['down', 'space'])
                           time.sleep(time_delay)
                       else:
                           pyautogui.typewrite(['down'])

                       pyautogui.typewrite(['esc', 'return'], interval=short_delay)
                       pyautogui.hotkey('ctrl', 'c')
                       #i['pt_num'] = str(k*2)
                       i['pt_num'] = Tk().clipboard_get()
                       time.sleep(time_delay)

                   else:
                       #pyautogui.typewrite(i[suborder_pt])
                       pyautogui.typewrite('000123123')#Susie Suborder place holder for human to do it
                       pyautogui.hotkey('enter')


               # IMPRESSION TYPE, FOOT TO SCAN (MAIN)
               pyautogui.typewrite(['tab'])
               time.sleep(time_delay)
               if i['sameday_suborder'] == 'yes':
                   if orders[i['suborder_target']]['foot2scan'] == 'Human Help!':
                       pyautogui.typewrite(['space', 'tab'])
                       time.sleep(short_delay)
                   else:
                       pyautogui.typewrite(['delete', 'tab', 'tab'])
                       pyautogui.typewrite(orders[i['suborder_target']]['foot2scan'])
                       pyautogui.typewrite(['tab'])
               else:
                   if i['sub_order'] == 'new order':
                       if i['foot2scan'] == 'Human Help!':
                           pyautogui.typewrite(['space', 'tab'])
                           time.sleep(short_delay)
                       else:
                           pyautogui.typewrite(['a', 'tab', 'tab'], interval=short_delay)
                           time.sleep(short_delay)
                           pyautogui.typewrite(i['foot2scan'])
                           pyautogui.typewrite(['tab'])

                   else:
                       pyautogui.typewrite(['tab', 'tab', 'tab'])
               # PURCHASE ORDER NUMBER (MAIN)
               pyautogui.typewrite(['tab'])
               pyautogui.typewrite(i['po_num'])

               # DEVICE SELECTION (MAIN)
               pyautogui.hotkey('alt', 'm')
               time.sleep(1.5)
               if i['sameday_suborder'] == 'yes':
                   pyautogui.typewrite(orders[i['suborder_target']]['wo_num'])
                   pyautogui.typewrite(['down'])
                   pyautogui.typewrite('changed device')
                   pyautogui.hotkey('alt', 'm')
                   pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=0.10)
                   pyautogui.typewrite(i['device_code'])
                   pyautogui.typewrite(['return', 'return'], interval=time_delay)
                   time.sleep(8)
               else:
                   if i['sub_order'] == 'new order':
                       pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=0.10)
                       pyautogui.typewrite(i['device_code'])
                       pyautogui.typewrite(['return', 'return'], interval=time_delay)
                       time.sleep(12)

                   if i['sub_order'] == 'changed':
                       pyautogui.typewrite(i['prev_po'])
                       pyautogui.typewrite(['down'])
                       pyautogui.typewrite('changed device')
                       pyautogui.hotkey('alt', 'm')
                       pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=0.10)
                       if i['device_code'] == '':
                           i['device_code']='43170ST01'
                       pyautogui.typewrite(i['device_code'])
                       pyautogui.typewrite(['return', 'return'], interval=0.10)
                       time.sleep(12)

                   if i['sub_order'] == 'duplicate':
                       pyautogui.typewrite(i['prev_po'])
                       pyautogui.typewrite(['down'])
                       pyautogui.typewrite('duplicate')
                       pyautogui.hotkey('alt', 'm')
                       pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=0.10)
                       if i['device_code'] == '':
                           i['device_code']='43170ST01'
                       pyautogui.typewrite(i['device_code'])
                       pyautogui.typewrite(['return', 'return'], interval=0.10)
                       time.sleep(12)

            # return to main screen
               pyautogui.hotkey('alt', 'g')
               time.sleep(2)
               pyautogui.typewrite(['right'])

            #rush or on time
               if i['sub_order'] == 'duplicate':
                   pyautogui.typewrite(['delete'])
               if i['sub_order'] == 'changed':
                   pyautogui.typewrite(['delete'])

               pyautogui.typewrite(['tab', 'tab', 'tab'])

               if i['priority'] == 'RRU On Time':
                   pyautogui.typewrite(['s', 'p', 'return', 'up', 'right', 'space'])
                   time.sleep(5)
               if i['priority'] == '3day Rush':
                   pyautogui.typewrite(['3', 'return', 'return', 'up', 'right', 'space'])
                   time.sleep(5)
               #pyautogui.typewrite(['f3', 'tab'], interval=time_delay)
               #end of order clean reset

            #QUANTITY 2 or more
               q_plus = i['quantity']
               while q_plus - 1 > 0:
                   q_plus = q_plus - 1
                   # start new order and collect wo#
                   pyautogui.press(['f3', 'tab', 'up', 'up', 'up'])
                   pyautogui.hotkey('ctrl', 'c')
                   #select firefly as account
                   pyautogui.hotkey('alt', 'g')
                   time.sleep(time_delay)
                   pyautogui.press(['0', '7', '4', '5', 'return'])
                   time.sleep(2)

                   # select clinician
                   pyautogui.typewrite(['f6'])
                   time.sleep(2)
                   pyautogui.typewrite('martin mcgeough')
                   time.sleep(time_delay)
                   pyautogui.typewrite(['return', 'return', 'return'], interval=time_delay)

                   #pick pt card by saved pt_num number
                   pyautogui.typewrite(i['pt_num'])
                   time.sleep(1.5)

                   # impression type, foot to scan
                   pyautogui.typewrite(['tab', 'delete', 'tab', 'tab'], interval=short_delay)
                   time.sleep(short_delay)
                   pyautogui.typewrite(i['foot2scan'])
                   pyautogui.typewrite(['tab'])

                   # po number
                   pyautogui.typewrite(['tab'])
                   pyautogui.typewrite(i['po_num'])

                   # DEVICE SELECTION
                   pyautogui.hotkey('alt', 'm')
                   time.sleep(1)
                   pyautogui.typewrite(i['wo_num'])
                   pyautogui.typewrite(['down'])
                   pyautogui.typewrite('duplicate')
                   pyautogui.hotkey('alt', 'm')
                   pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=0.10)
                   if i['device_code'] == '':
                       i['device_code'] = '43170ST01'
                   pyautogui.typewrite(i['device_code'])
                   pyautogui.typewrite(['return', 'return'], interval=0.10)
                   time.sleep(8)

                   pyautogui.hotkey('alt', 'g')# return to main screen
                   time.sleep(8)
                   pyautogui.typewrite(['right'])

                   # RUSH OR ONTIME SPECIFIED
                   pyautogui.typewrite(['tab', 'tab', 'tab'])
                   if i['priority'] == 'RRU On Time':
                       pyautogui.typewrite(['s', 'p', 'return', 'up', 'right', 'space'])
                       time.sleep(5)
                   if i['priority'] == '3day Rush':
                       pyautogui.typewrite(['3', 'return', 'return', 'up', 'right', 'space'])
                       time.sleep(5)
                   # pyautogui.typewrite(['f3', 'tab'], interval=time_delay)
                   #end of order clean reset

                #pyautogui.press(['esc', 'esc', 'enter', 'esc', 'esc', 'enter'], interval=0.25)

print '                            <[ [ [PROGRAM COMPLETE] ] ]>'


print len(orders)
















################################
################################
################################






# import xlrd
# import json
# import os
# import re
# import datetime
# import pyautogui
# import time
# import PIL
#
#
# # from PIL import Image
# # myImage = Image.open('hello.PNG')
# # print '======= START ======='
# # myImage.filename
# # pyautogui.locateOnScreen(myImage)
# # print '=======  END  ======='
# #
#
# deviceDictionary = {
#     #"FUNCTIONAL STANDARD":"STANDARD FUNCTIONAL", "EVA":"EVA", "SPORT STANDARD - NEOPRENE TO TOES":"43172ST01"
#
# "FIREFLY NHS FUNCTIONAL":"43170DM01", "FIREFLY NHS DRESS":"431750611", "FIREFLY NHS SPORT":"431750621",
# "FIREFLY SOCCER SPORT":"431750626", "FIREFLY SOCCER SPORT (DM)":"431750627", "FIREFLY SPORT IMPACT":"431750622",
# "FUNCTIONAL STANDARD":"43170ST01", "FUNCTIONAL DIRECT MILLED":"43170DM01", "STANDARD SLIMLINE":"43171LA01",
# "LOW HEEL CUP SLIMLINE":"43171LA11", "FLAT HEEL CUP":"43171LA21", "COBRA":"43171LA31", "MENS DRESS":"43171ME01",
# "SPORT STANDARD - NEOPRENE TO TOES":"43172ST01", "SPORT DIRECT MILLED - NEOPRENE TO TOES":"43172DM01",
# "SPORT DIRECT MILLED - VINYL TO METS":"43172DM02", "SPORT LOW PROFILE":"43172LP01", "SPORT SKI - ALPINE":"43172SI01",
# "SPORT SKI - NORDIC":"43172SI02", "SPORT SKI - SNOWBOARD":"43172SI03", "SPORT SKATE - HOCKEY":"43172SA01",
# "SPORT SKATE - FIGURE":"43172SA02", "MOLD STANDARD":"43173ST01", "MOLD LOW PROFILE":"43173LP11",
# "FIREFLY DIABETIC TRIDENSITY":"431750671", "FIREFLY RA FLEXIBLE MOLD":"431750681", "EVA":"43174EV01",
# "UCBL":"43174UC01", "ROBERTS WHITMAN":"43174RB01", "GAIT PLATE - INDUCE OUT-TOEING":"43174GP02",
# "GAIT PLATE - INDUCE IN-TOEING":"43174GP01", "":"STAND-IN"
# }
#
# print deviceDictionary["SPORT STANDARD - NEOPRENE TO TOES"]
# print deviceDictionary
#
#
# targetdir = os.path.join("firefly_orders", "orders")
# files = os.listdir(targetdir)
# xlsfiles = []
#
# def convertserialtodate(xlserial):
#     basedate = datetime.date(1900,1,1)
#     delta = datetime.timedelta(days=xlserial)
#     newdate = basedate + delta
#     newdate.strftime("%m%d%y")
#     output = newdate.strftime("%m%d%y")
#     return output
#
# print convertserialtodate(41875)
#
#
# for f in files:
#     if f.endswith("xls"):
#         xlsfiles.append(f)
#
#
# orders = []
# for xfile in xlsfiles:
#     filepath = os.path.join(targetdir,xfile)
#     #dowork
#
#     xl_workbook = xlrd.open_workbook(filepath)
#     xl_sheet = xl_workbook.sheet_by_index(1)
#     print ('Sheet name: %s' % xl_sheet.name)
#     print xlrd.xldate_as_datetime(42088,xl_workbook.datemode)
#
#
#
#
#     curr_order = None
#     for i in range (xl_sheet.nrows):
#         row = xl_sheet.row(i)
#         if row[0].value=="PATIENT NAME / CODE NO.":
#             curr_order = {}
#             curr_order["po_num"] = str(row[-1].value).strip().strip("\s").rstrip(".0")
#             curr_order["name"] = row[1].value.strip()
#             #finding all character strings, hyphen and apostraphy included
#             nameList = (re.findall("[a-zA-Z'-]+", row[1].value.strip()))
#             #finding if there is a pt serial number
#             nameNumber = (re.findall("[\d]+",row[1].value.strip()))
#             print nameList
#
#             curr_order["firstname"]=""
#             curr_order["nameLast"] = ""
#             if len(nameList) >= 1:
#                 curr_order["firstname"] = " ".join(nameList[0:-1])
#                 curr_order["nameLast"] = nameList[-1]
#
#             curr_order["nameNumber"] = ""
#             if len(nameNumber)==1:
#                 curr_order["nameNumber"] = nameNumber[0]
#
#             curr_order["FIRSTNAMECOMPLETE"] = curr_order["firstname"] + " " +curr_order["nameNumber"]
#
#
#
#
#         if row[0].value=="WEIGHT RANGE / SIZE OF FOOT / PRIORITY / TEMPLATE":
#             weightLen = (re.findall('[\d]+',row[1].value))
#             if len(weightLen)==4:
#                 curr_order["weight"] = weightLen[1]
#             if len(weightLen)==2:
#                 curr_order["weight"] = weightLen[0]
#
#             curr_order["shoesize"] = re.findall('[.\d]+',row[4].value)[0]
#             curr_order["priority"] = row[6].value
#
#         if row[0].value=="QUANTITY / SUBSEQUENT PAIR":
#             curr_order["quantity"] = row[1].value
#             #removing white space, extra numbers, carriage return
#             prev_po = str(row[4].value).strip().strip("\s").rstrip(".0")
#             curr_order["prev_po"] = prev_po
#             sub_order = row[7].value
#             if sub_order == "":
#                 sub_order = "new order"
#             elif sub_order == "CHANGED DEVICE (Select device and options)":
#                 sub_order = "changed"
#             elif sub_order == "DUPLICATE DEVICE (No change)":
#                 sub_order = "duplicate"
#             else:
#                 sub_order = "error"
#
#             curr_order["sub_order"] = sub_order
#             print curr_order["sub_order"]
#
#
#         if row[0].value=="OUTGROWTH PAIR / DOB":
#             curr_order["outgrowth"] = row[1].value.strip()
#             match = re.search(r'pair', curr_order["outgrowth"])
#             #ogConfirm = (re.findall("[a-zA-Z]"))
#             if match:
#                 curr_order["outgrowth"] = 1
#             else :
#                 curr_order["outgrowth"] = 0
#
#             if row[7].ctype == 3:
#                 curr_order["dob"] = xlrd.xldate_as_datetime(int(row[7].value),xl_workbook.datemode).strftime("%m%d%y")
#                 print row
#             else:
#                 curr_order["dob"]= ""
#
#         if row[0].value=="DEVICE":
#             curr_order["device"] = str(row[1].value).strip().strip("\s")
#             poro = curr_order["device"]
#             device_code = deviceDictionary[poro]
#             print poro
#             curr_order["device_code"] = device_code
#
#
#         if row[0].value=="FOOT SCANNED":
#             curr_order["both"] = row[1].value
#             curr_order["left"] = row[4].value
#             curr_order["right"] = row[7].value
#
#             if (not curr_order["both"] and not curr_order["right"] and not curr_order["left"]):
#                 curr_order["foot2scan"] = "SUBSEQUENT PAIR ORDER"
#             elif curr_order["both"] and (not curr_order["right"] and not curr_order["left"]) \
#                     or (curr_order["right"] and curr_order["left"] and not curr_order["both"] ) :
#                 curr_order["foot2scan"] = "BOTH"
#             elif curr_order["right"] and not curr_order["left"] and not curr_order["both"]:
#                 curr_order["foot2scan"] = "RIGHT"
#             elif curr_order["left"] and not curr_order["right"] and not curr_order["both"]:
#                 curr_order["foot2scan"] = "LEFT"
#             else:
#                 curr_order["foot2scan"] = "Human Help!"
#
#         if row[0].value=="NOTES":
#             nextrow = xl_sheet.row(i+1)
#             curr_order["notes"] = nextrow[0].value
#             orders.append(curr_order)
#
# keyPhrases = {
#     "HOLDFORNOW":[ u"HOLD",u"FOR",u"NOW"]
# }
#
# # def longestSubstringFinder(string1, string2):
# #     answer = []
# #     len1, len2 = len(string1), len(string2)
# #     for i in range(len1):
# #         match = []
# #         for j in range(len2):
# #             if (i + j < len1 and string1[i + j] == string2[j]):
# #                 match += string2[j]
# #             else:
# #                 if (len(match) > len(answer)):
# #                     answer = match
# #                 match = []
# #     return answer
# #
# # print longestSubstringFinder("apple pie available", "apple pies")
# # print longestSubstringFinder("apples", "appleses")
# # print longestSubstringFinder("bapples", "cappleses")
#
#
# #===== CONTROL PANEL ====================================================================
#
# #pyautogui.click(1046,743)
# pyautogui.click(100,100)
# order_limit =10
# time_delay = 1.5
# short_delay = 0.25
# lag = 40
#
# #=========================================================================================
#
#
# k=0
# for i in orders:
#     k=k+1
#     print""
#     print""
#     print""
#     print "                          ","ORDER" , k
#     #print json.dumps(i,indent=4)
#     print ""
#     print i["sub_order"]
#     print "NAME ON ORDER =", i["name"]
#     print "FIRST NAME FIELD =", i["FIRSTNAMECOMPLETE"]
#     print "LAST NAME FIELD =", i["nameLast"]
#     print "DOB =", i["dob"]
#     print "OUTGROWTH =", i["outgrowth"]
#     print "WEIGHT =", i["weight"]
#     print "SHOE SIZE =", i["shoesize"]
#     print "PO NUMBER =", i["po_num"]
#     print "PREVIOUS PO# =",i["prev_po"]
#     print "FOOT TO SCAN =", i["foot2scan"]
#     print "PRIORITY =", i["priority"]
#     print "QUANTITY =", i ["quantity"]
#     print "NOTES =", i ["notes"]
#     print "DEVICE =", i["device"]
#     print "D.CODE =", i["device_code"]
#     tokenizednotes = i["notes"].split(" ")
#         #############################   CONSOLE OUT PUT ONLY    ###############
#         if k <= order_limit:
#         # select firefly as account
#         pyautogui.hotkey('alt', 'g')
#         time.sleep(time_delay)
#         pyautogui.typewrite('0745')
#         time.sleep(3)
#         pyautogui.typewrite(['return'])
#         time.sleep(time_delay)
#
#         # select clinician
#         pyautogui.typewrite(['f6'])
#         time.sleep(3)
#         pyautogui.typewrite("Martin McGeough")
#         time.sleep(time_delay)
#         pyautogui.typewrite(['return', 'return', 'return'], interval=time_delay)
#
#         # create pt card
#         #if i["sub_order"] == True:
#         # if i['sub_order'] == ("duplicate"):
#         #     #pyautogui.typewrite('DUPE PT Maker')
#         #     pyautogui.hotkey('alt', 'g')
#         #     time.sleep(time_delay)
#         #     pyautogui.typewrite(['right', 'tab'])
#         #
#         # if i['sub_order'] == ("changed"):
#         #     #pyautogui.typewrite(' CHANGED PT Marker')
#         #     pyautogui.hotkey('alt', 'g')
#         #     time.sleep(time_delay)
#         #     pyautogui.typewrite(['right', 'tab'])
#
#         if i['sub_order'] == ("new order"):
#             pyautogui.typewrite(['f6'], interval=time_delay)
#             time.sleep(lag)
#             #pt name search
#             pyautogui.typewrite(i["FIRSTNAMECOMPLETE"])
#             time.sleep(time_delay)
#             pyautogui.typewrite(['tab'])
#             pyautogui.typewrite(i["nameLast"])
#             time.sleep(time_delay)
#             pyautogui.hotkey('alt', 's')
#             time.sleep(time_delay)
#             pyautogui.hotkey('alt', 'c')
#             time.sleep(time_delay)
#             pyautogui.hotkey('alt', 's')
#             time.sleep(lag)
#             pyautogui.typewrite(['y'])
#             time.sleep(time_delay)
#             # entering pt info
#             pyautogui.typewrite(['tab', 'tab', 'm', 'tab'], interval=time_delay) #gender
#             time.sleep(time_delay)
#             #pyautogui.typewrite('xxx')
#             pyautogui.typewrite(i["dob"])
#             # pyautogui.typewrite('xxx')
#             # if i["dob"] == True:
#             #     pyautogui.typewrite(i["dob"])
#             #     time.sleep(time_delay)
#
#             pyautogui.typewrite([i["dob"], 'tab', i["weight"], 'tab', i["shoesize"], 'tab', i["shoesize"]], interval=short_delay)
#             time.sleep(time_delay)
#
#             if i["outgrowth"] == 1:
#                 pyautogui.typewrite(['down', 'space'])
#                 time.sleep(time_delay)
#             else:
#                 pyautogui.typewrite(['down'])
#
#             pyautogui.typewrite(['esc', 'return'], interval=short_delay)
#             time.sleep(time_delay)
#
#
#         # impression type, foot to scan
#         pyautogui.typewrite(['tab'])
#         time.sleep(time_delay)
#         if i["sub_order"] == ("duplicate"):
#             pyautogui.typewrite(['delete', 'tab', 'tab', 'tab'], interval=short_delay)
#             time.sleep(time_delay)
#             #pyautogui.typewrite("MARKER dupe")
#         if i["sub_order"] == ("changed"):
#             pyautogui.typewrite(['delete', 'tab', 'tab', 'tab'], interval=short_delay)
#             time.sleep(time_delay)
#             #pyautogui.typewrite("MARKER change")
#
#         if i["sub_order"] == ("new order"):
#             pyautogui.typewrite(['a', 'tab', 'tab'], interval=short_delay)
#             time.sleep(time_delay)
#             if i["foot2scan"] == "RIGHT":
#                 pyautogui.typewrite(['r', 'tab'])
#                 time.sleep(time_delay)
#             if i["foot2scan"] == "LEFT":
#                 pyautogui.typewrite(['l', 'tab'])
#                 time.sleep(time_delay)
#             if i["foot2scan"] == "BOTH":
#                 pyautogui.typewrite(['b', 'tab'])
#                 time.sleep(time_delay)
#             if i["foot2scan"] == "Human Help!":
#                 pyautogui.typewrite(['space', 'tab'])
#                 time.sleep(time_delay)
#
#         # po number
#         pyautogui.typewrite(['tab'])
#         time.sleep(short_delay)
#         pyautogui.typewrite(i["po_num"])
#         time.sleep(short_delay)
#
#         # device selection tab
#         pyautogui.hotkey('alt', 'm')
#         time.sleep(2)
#
#         if i["sub_order"] == ("new order"):
#             #pyautogui.typewrite("  new order marker")
#             pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=0.10)
#             #pyautogui.typewrite(i["device"])
#             pyautogui.typewrite(i["device_code"])
#             pyautogui.typewrite(['return', 'return'], interval=time_delay)
#             time.sleep(8)
#
#     # if i["sub_order"] == "changed" or "duplicate":
#     #     pyautogui.typewrite(['down'])
#     #     time.sleep(time_delay)
#         if i["sub_order"] == "changed":
#             #pyautogui.typewrite("  changed order marker   Previous PO# : ")
#             pyautogui.typewrite(i["prev_po"])
#             pyautogui.typewrite(['down'])
#             pyautogui.typewrite(" changed device")
#             pyautogui.hotkey('alt', 'm')
#             pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=0.10)
#             if i["device_code"] == (""):
#                 i["device_code"]="43170ST01"
#             pyautogui.typewrite(i["device_code"])
#             pyautogui.typewrite(['return', 'return'], interval=0.10)
#             time.sleep(3)
#
#         if i["sub_order"] == "duplicate":
#             #pyautogui.typewrite("  duplicate order marker   Previous PO# : ")
#             pyautogui.typewrite(i["prev_po"])
#             pyautogui.typewrite(['down'])
#             pyautogui.typewrite(" duplicate")
#             pyautogui.hotkey('alt', 'm')
#             pyautogui.typewrite(['tab', 'esc', 'tab', 'esc'], interval=0.10)
#             if i["device_code"] == (""):
#                 i["device_code"]="43170ST01"
#             pyautogui.typewrite(i["device_code"])
#             pyautogui.typewrite(['return', 'return'], interval=0.10)
#             time.sleep(3)
#
#
#
#     # return to main screen
#         pyautogui.hotkey('alt', 'g')
#         time.sleep(time_delay)
#     #rush or on time
#
#         pyautogui.typewrite(['right'])
#
#         if i["sub_order"] == "duplicate":
#             pyautogui.typewrite(['delete'])
#
#         if i["sub_order"] == "changed":
#             pyautogui.typewrite(['delete'])
#
#         pyautogui.typewrite(['tab', 'tab', 'tab'])
#
#         if i["priority"] == "RRU On Time":
#             #pyautogui.typewrite(i["priority"])
#             pyautogui.typewrite(['s', 'p', 'return', 'up', 'right', 'space'])
#             time.sleep(5)
#         if i["priority"] == "3day Rush":
#             #pyautogui.typewrite(i["priority"])
#             pyautogui.typewrite(['3', 'return', 'return', 'up', 'right', 'space'])
#             time.sleep(5)
#         pyautogui.typewrite(['f3', 'tab'], interval=time_delay)
# #pyautogui.typewrite("finish")

# pyautogui.hotkey('alt','g')
    # time.sleep(1)
    # pyautogui.typewrite('FIRE')
    # time.sleep(1)
    # pyautogui.typewrite(['return'])
    # time.sleep(1)
    #









    # time.sleep(0.5)
    # pyautogui.typewrite(['y'])
    # time.sleep(0.5)
    # pyautogui.typewrite(['y','return','F6'])

    # pyautogui.typewrite(['return', 'return', 'a', 'return', 'return', 'L', 'return'])
    # time.sleep(1)
    # pyautogui.hotkey('alt', 'm')
    # time.sleep(1)
    # pyautogui.typewrite(['tab', 'return', 'return', 'return', 'return', 'return'])
    # time.sleep(6)
    # pyautogui.typewrite(['return'])
    # time.sleep(1)
    # pyautogui.hotkey('alt', 'g')
    # time.sleep(1)
    # pyautogui.typewrite(['right', 'down', 'down', 'down'])
    # time.sleep(1)
    # pyautogui.typewrite(['s', 'p'])
    # time.sleep(1)
    # pyautogui.typewrite(['return', 'up', 'right', 'space'])
    #
    #
    # exit()
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["name"])
    # time.sleep(1)
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["device"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["po_num"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["weight"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["shoesize"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["priority"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["foot2scan"])
    # pyautogui.typewrite(['return','return'])

   # print tokenizednotes

#    for keyPhrase in keyPhrases:
#        print tokenizednotes
#        print keyPhrases[keyPhrase]
#        print longestSubstringFinder(tokenizednotes, keyPhrases[keyPhrase])






# Print 1st row values and types
#
#from xlrd.sheet import ctype_text
#print('(Column #) type:value')
#for idx, cell_obj in enumerate(row):
  # cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
   # print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))


