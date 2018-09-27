import xlrd
import json
import os
import re
import datetime
import time
import pyautogui


deviceDictionary = {
    "FUNCTIONAL STANDARD":"STANDARD FUNCTIONAL",
    "EVA":"EVA"
}

print deviceDictionary

if "FUNCTIONAL STANDARD" in deviceDictionary:
    print "FOUND!!!!!!!!"
    print deviceDictionary["FUNCTIONAL STANDARD"]

targetdir = os.path.join("firefly_orders","orders")
files = os.listdir(targetdir)
xlsfiles = []

def convertserialtodate(xlserial):
    basedate = datetime.date(1900,1,1)
    delta = datetime.timedelta(days=xlserial)
    newdate = basedate + delta
    newdate.strftime("%m%d%y")
    output = newdate.strftime("%m%d%y")
    return output

print convertserialtodate(41875)


for f in files:
    if f.endswith("xls"):
        xlsfiles.append(f)


orders = []
for xfile in xlsfiles:
    filepath = os.path.join(targetdir,xfile)
    #dowork

    xl_workbook = xlrd.open_workbook(filepath)
    xl_sheet = xl_workbook.sheet_by_index(1)
    print ('Sheet name: %s' % xl_sheet.name)
    print xlrd.xldate_as_datetime(42088,xl_workbook.datemode)




    curr_order = None
    for i in range (xl_sheet.nrows):
        row = xl_sheet.row(i)
        if row[0].value=="PATIENT NAME / CODE NO.":
            curr_order = {}
            curr_order["po_num"] = str(row[-1].value).strip().strip("\s").rstrip(".0")
            curr_order["name"] = row[1].value.strip()
            #finding all character strings, hyphen and apostraphy included
            nameList = (re.findall("[a-zA-Z'-]+", row[1].value.strip()))
            #finding if there is a pt serial number
            nameNumber = (re.findall("[\d]+",row[1].value.strip()))
            print nameList
            #print nameLen[-1]
            #print len(nameLen)
            #print "max"
            #print max( nameLen)

            #extracts the first and last names from the name field
            #the last name is the last element of the nameList
            #the first name is the rest of the elements from nameList
            # curr_order["firstname"] = ""
            # for x in xrange(len(nameList)):
            #     if x == len(nameList) -1 :
            #         print "The Last Name is ", nameList[x]
            #         curr_order["nameLast"] = nameList[x]
            #     else:
            #         tName = curr_order["firstname"] + nameList[x] + " "
            #         curr_order["firstname"] = tName
            #         print curr_order["firstname"]
            curr_order["firstname"]=""
            curr_order["nameLast"] = ""
            if len(nameList)>1:
                curr_order["firstname"] = " ".join(nameList[0:-1])
                curr_order["nameLast"] = nameList[-1]

            curr_order["nameNumber"] = ""
            if len(nameNumber)==1:
                curr_order["nameNumber"] = nameNumber[0]

            curr_order["FIRSTNAMECOMPLETE"] = curr_order["firstname"] + " " +curr_order["nameNumber"]




        if row[0].value=="WEIGHT RANGE / SIZE OF FOOT / PRIORITY / TEMPLATE":
            weightLen = (re.findall('[\d]+',row[1].value))
            if len(weightLen)==4:
                curr_order["weight"] = weightLen[1]
            if len(weightLen)==2:
                curr_order["weight"] = weightLen[0]

            curr_order["shoesize"] = re.findall('[.\d]+',row[4].value)[0]
            curr_order["priority"] = row[6].value

        if row[0].value=="QUANTITY / SUBSEQUENT PAIR":
            curr_order["quantity"] = row[1].value
            #removing white space, extra numbers, carriage return
            prev_po = str(row[4].value).strip().strip("\s").rstrip(".0")
            curr_order["prev_po"] = prev_po
            curr_order["sub_order"] = row[7].value


        if row[0].value=="OUTGROWTH PAIR / DOB":
            curr_order["outgrowth"] = row[1].value
            if row[7].ctype == 3:
                curr_order["dob"] = xlrd.xldate_as_datetime(int(row[7].value),xl_workbook.datemode).strftime("%m%d%y")
                print row
            else:
                curr_order["dob"]= ""

        if row[0].value=="DEVICE":
            curr_order["device"] = row[1].value
#selecting the right row that has the foot to scan info you need
        if row[0].value=="FOOT SCANNED":
            curr_order["both"] = row[1].value
            curr_order["left"] = row[4].value
            curr_order["right"] = row[7].value

            if (not curr_order["both"] and not curr_order["right"] and not curr_order["left"]):
                curr_order["foot2scan"] = "SUBSEQUENT PAIR ORDER"
            elif curr_order["both"] and (not curr_order["right"] and not curr_order["left"]) \
                    or (curr_order["right"] and curr_order["left"] and not curr_order["both"] ) :
                curr_order["foot2scan"] = "BOTH"
            elif curr_order["right"] and not curr_order["left"] and not curr_order["both"]:
                curr_order["foot2scan"] = "RIGHT"
            elif curr_order["left"] and not curr_order["right"] and not curr_order["both"]:
                curr_order["foot2scan"] = "LEFT"
            else:
                curr_order["foot2scan"] = "Human Help!"








        if row[0].value=="NOTES":
            nextrow = xl_sheet.row(i+1)
            curr_order["notes"] = nextrow[0].value
            orders.append(curr_order)

keyPhrases = {
    "HOLDFORNOW":[ u"HOLD",u"FOR",u"NOW"]
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



pyautogui.click(100,100)




k=0
for i in orders:
    k=k+1
    print""
    print""
    print""
    print "                          ","ORDER" , k
    #print json.dumps(i,indent=4)
    print ""
    print "NAME ON ORDER =", i["name"]
    print "FIRST NAME FIELD =", i["FIRSTNAMECOMPLETE"]
    print "LAST NAME FIELD =", i["nameLast"]
    print "DOB =", i["dob"]
    print "SUBSEQUENT =", i["sub_order"]
    print "OUTGROWTH =", i["outgrowth"]
    print "WEIGHT =", i["weight"]
    print "SHOE SIZE =", i["shoesize"]
    print "PO NUMBER =", i["po_num"]
    print "PREVIOUS PO# =", i["prev_po"]
    print "FOOT TO SCAN =", i["foot2scan"]
    print "PRIORITY =", i["priority"]
    print "QUANTITY =", i["quantity"]
    print "NOTES =", i["notes"]
    tokenizednotes = i["notes"].split(" ")

    time_delay = 0
    # create new order
    pyautogui.typewrite(['return'])
    pyautogui.typewrite("Order Start Marker   ")
    pyautogui.typewrite(i["priority"])
    pyautogui.typewrite("xxx")
    # pyautogui.typewrite(['f3'])
    time.sleep(time_delay)
    pyautogui.typewrite(['tab'])
    time.sleep(time_delay)

    # select firefly as account
    pyautogui.hotkey('alt', 'g')
    time.sleep(time_delay)
    pyautogui.typewrite('FIRE')
    time.sleep(time_delay)
    pyautogui.typewrite(['return'])
    time.sleep(time_delay)

    # select clinician
    pyautogui.typewrite(['tab', 'f6'], interval=time_delay)
    pyautogui.typewrite("Martin McGeough")
    time.sleep(time_delay)
    pyautogui.typewrite(['return', 'return'], interval=time_delay)

    # create pt card
    if i["sub_order"] == True:
        pyautogui.typewrite("existing pt")
    else:
        pyautogui.typewrite(['tab', 'f6'], interval=time_delay)
        # time.sleep(40)
        pyautogui.typewrite(i["FIRSTNAMECOMPLETE"])
        time.sleep(time_delay)
        pyautogui.typewrite(['tab'])
        pyautogui.typewrite(i["nameLast"])
        time.sleep(time_delay)
        pyautogui.hotkey('alt', 's')
        time.sleep(time_delay)
        pyautogui.hotkey('alt', 'c')
        time.sleep(time_delay)
        pyautogui.hotkey('alt', 's')
        # time.sleep(40)
        pyautogui.typewrite(['y'])
        time.sleep(time_delay)

        # entering pt info
        pyautogui.typewrite(['tab', 'tab', 'm', 'tab'], interval=time_delay)  # gender
        time.sleep(time_delay)
        if i["dob"] == True:
            pyautogui.typewrite(i["dob"])
            time.sleep(time_delay)

        pyautogui.typewrite(['tab', i["weight"], 'tab', i["shoesize"], 'tab', i["shoesize"], 'return'],
                            interval=time_delay)
        time.sleep(time_delay)

        if i["outgrowth"] == True:
            pyautogui.typewrite(['space'])
            time.sleep(time_delay)

        pyautogui.typewrite(['esc', 'return'], interval=time_delay)
        time.sleep(time_delay)

    # impression type, foot to scan
    pyautogui.typewrite(['tab'])
    time.sleep(time_delay)
    if i["foot2scan"] == "SUBSEQUENT PAIR ORDER":
        pyautogui.typewrite(['delete', 'tab', 'tab'], interval=time_delay)
        time.sleep(time_delay)
    else:
        pyautogui.typewrite(['a', 'tab', 'tab'], interval=time_delay)
        time.sleep(time_delay)
        if i["foot2scan"] == "RIGHT":
            pyautogui.typewrite(['r', 'tab'], interval=time_delay)
            time.sleep(time_delay)
        if i["foot2scan"] == "LEFT":
            pyautogui.typewrite(['l', 'tab'], interval=time_delay)
            time.sleep(time_delay)
        if i["foot2scan"] == "BOTH":
            pyautogui.typewrite(['b', 'tab'], interval=time_delay)
            time.sleep(time_delay)
        if i["foot2scan"] == "Human Help!":
            pyautogui.typewrite(['space', 'tab'], interval=time_delay)
            time.sleep(time_delay)

    # po number
    pyautogui.typewrite(['tab'])
    time.sleep(time_delay)
    pyautogui.typewrite(i["po_num"])
    time.sleep(time_delay)

    # device selection tab
    pyautogui.hotkey('alt', 'm')
    time.sleep(time_delay)

    if i["sub_order"] == ("new order"):
        pyautogui.typewrite("  new order marker")
        pyautogui.typewrite(['tab', 'esc', 'tab', 'esc', "43170ST01", 'return', 'return'], interval=time_delay)
        time.sleep(time_delay)

        # if i["sub_order"] == "changed" or "duplicate":
        #     pyautogui.typewrite(['down'])
        #     time.sleep(time_delay)
    if i["sub_order"] == "changed":
        pyautogui.typewrite("  changed order marker   Previous PO# : ")
        pyautogui.typewrite(i["prev_po"])
        pyautogui.typewrite(['down'])
        pyautogui.typewrite(" changed device")
        time.sleep(time_delay)
    if i["sub_order"] == "duplicate":
        pyautogui.typewrite("  duplicate order marker   Previous PO# : ")
        pyautogui.typewrite(i["prev_po"])
        pyautogui.typewrite(['down'])
        pyautogui.typewrite(" duplicate")
        time.sleep(time_delay)

        # return to main screen
    pyautogui.hotkey('alt', 'g')
    time.sleep(time_delay)
    # rush or on time
    if i["priority"] == True:
        pyautogui.typewrite(['right', 'tab', 'tab', 'tab'])
        pyautogui.typewrite(i["priority"])
        time.sleep(time_delay)
        if i["priority"] == "RRU On Time":
            pyautogui.typewrite(i["priority"])
            pyautogui.typewrite(['s', 'p'])
            time.sleep(time_delay)
            pyautogui.typewrite(['return'])
            time.sleep(time_delay)
            pyautogui.typewrite(['up', 'right', 'space'])
            time.sleep(time_delay)
        elif i["priority"] == "3day Rush":
            pyautogui.typewrite(i["priority"])
            pyautogui.typewrite(['3'])
            time.sleep(time_delay)
            pyautogui.typewrite(['return', 'return'])
            time.sleep(time_delay)
            pyautogui.typewrite(['up', 'right', 'space'])
            time.sleep(time_delay)
pyautogui.typewrite("finish")

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
    #
    #
    #
    #
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["name"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["po_num"])
    # #time.sleep(1)
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["device"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["weight"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["shoesize"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["priority"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["foot2scan"])
    # pyautogui.typewrite(['return'])
    # pyautogui.typewrite(i["notes"])
    # pyautogui.typewrite(['return', 'return'])




   # print tokenizednotes

    # for keyPhrase in keyPhrases:
    #     print tokenizednotes
    #     print keyPhrases[keyPhrase]
    #     print longestSubstringFinder(tokenizednotes, keyPhrases[keyPhrase])

print ""
print "Total Orders Processed: ", len(orders)
print len(orders)


# Print 1st row values and types
#
#from xlrd.sheet import ctype_text
#print('(Column #) type:value')
#for idx, cell_obj in enumerate(row):
  # cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
   # print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))


