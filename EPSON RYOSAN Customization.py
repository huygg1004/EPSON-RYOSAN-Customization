import sys
import glob, os
import shutil
import configparser
import numpy as np
import pyx12
import pyx12.x12context
import keyboard
import xml.etree.ElementTree as ET
import logging
from datetime import datetime
from asammdf import MDF
from sys import argv
import pandas as pd
import csv
from datetime import date
from io import BytesIO

###Email Sending
import smtplib
import mimetypes
from email.message import EmailMessage



exec_path = os.path.dirname(os.path.abspath(__file__))
config = configparser.ConfigParser()
config.read(exec_path + "\config.ini")

INPUT_PATH = config.get('INPUT', 'PATH')
TEMP_PATH = config.get('TEMP', 'PATH')
OUTPUT_PATH = config.get('OUTPUT', 'PATH')
ARCHIVE_PATH = config.get('ARCHIVE', 'PATH')
XFAILED_PATH = config.get('XFAILED', 'PATH')
REPORTED_PATH = config.get('REPORTED', 'PATH')
EVENTLOG_PATH = config.get('EVENTLOG', 'PATH')
DSEVTLOG_PATH = config.get('DSEVTLOG', 'PATH')

INPUT_FILE_EXT = ".dat"
TMP_FILE_EXT = ".tmp"
OUTPUT_FILE_EXT = ".xml"

logging.basicConfig(level=logging.INFO, handlers=[logging.FileHandler(
    config.get('EVENTLOG', 'PATH') + 'Inbound_ToXMLTranslation_HK_Python-' + datetime.today().strftime(
        '%Y%m%d') + '.log')], format='%(asctime)s %(shortcode)s [%(levelname).3s] %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')
old_factory = logging.getLogRecordFactory()


def record_factory(*args, **kwargs):
    record = old_factory(*args, **kwargs)
    record.shortcode = "USC"
    return record


logging.setLogRecordFactory(record_factory)


def PrintMessage(msg):
    infoMsg = msg
    logging.info(infoMsg)
    print(infoMsg)
    return


def PrintWarning(warnMsg):
    logging.info(warnMsg)
    print(warnMsg)
    return


def PrintError(errMsg):
    exceptionErrorMsg = "[ERROR] " + "Exception error: ", sys.exc_info()[0], " occurred."
    logging.info(exceptionErrorMsg)
    print(exceptionErrorMsg)

    errMsg = "[ERROR] " + errMsg
    logging.info(errMsg)
    print(errMsg)
    return


def GetCurrentDatePath(path):
    current_time = datetime.today()
    year = current_time.year
    month = current_time.month
    day = current_time.day

    path = os.path.join(path, str(year))
    isExist = os.path.exists(path)
    if isExist == False:
        os.mkdir(path)

    path = os.path.join(path, str(month).zfill(2))
    isExist = os.path.exists(path)
    if isExist == False:
        os.mkdir(path)

    path = os.path.join(path, str(day).zfill(2))
    isExist = os.path.exists(path)
    if isExist == False:
        os.mkdir(path)
    return path


def MoveFile(srcFilePath, dstFilePath):
    if os.path.exists(dstFilePath):
        os.remove(dstFilePath)

    shutil.move(srcFilePath, dstFilePath)
    return


def CopyFile(srcFilePath, dstFilePath):
    if os.path.exists(dstFilePath):
        os.remove(dstFilePath)

    shutil.copyfile(srcFilePath, dstFilePath)
    return


def IsStrip(data):
    if not data is None:
        data = data.strip()
        return data
    else:
        return ""


def AddExtra(element, parent, code, label, content):
    # Adding <EXTRA> under element
    EXTRA = ET.SubElement(element, "EXTRA")
    EXTRA.set("PARENT", parent)
    # Adding subtags under <EXTRA>
    CODE = ET.SubElement(EXTRA, "CODE").text = code
    LABEL = ET.SubElement(EXTRA, "LABEL").text = label
    CONTENT = ET.SubElement(EXTRA, "CONTENT").text = content
    return


def AddRefnote(element, parent, code, label, contentList):
    # Adding <REFNOTE> under element
    REFNOTE = ET.SubElement(element, "REFNOTE")
    REFNOTE.set("PARENT", parent)
    # Adding subtags under <REFNOTE>
    CODE = ET.SubElement(REFNOTE, "CODE").text = code
    LABEL = ET.SubElement(REFNOTE, "LABEL").text = label
    for content in contentList:
        CONTENT = ET.SubElement(REFNOTE, "CONTENT").text = content
    return


def ReorderAttributes(root):
    for el in root.iter():
        attrib = el.attrib
        if len(attrib) > 1:
            if el.tag == "DOC824":
                # adjust attribute order, e.g. by sorting
                attribs = sorted(attrib.items())
                item1 = attribs[1]
                item2 = attribs[2]
                attribs[1] = item2
                attribs[2] = item1
                attrib.clear()
                attrib.update(attribs)
    return


PrintMessage("Start US Custom EDI355 translation...")

for file in os.scandir(INPUT_PATH):
    if file.is_file():

        PrintMessage("Start to process file: " + file.path)

        from pathlib import Path

        filePath = Path(file)
        name = filePath.name
        filenameWithoutExtension = filePath.stem
        ext = filePath.suffix
        fileSize = os.path.getsize(file)
        inputTmpFileTesting = os.path.join(TEMP_PATH, name)

        if ext == INPUT_FILE_EXT:
            # get data from EDI X12 file
            lineNo = 0

            CUST_PO_LINE = CUST_PO = PO_DATE = CUST_NO = SHIP_TO_NO = MATERIAL = CUST_MATERIAL = QUANTITY = UOM = UNIT_PRICE = CURRENCY = DELIVERY_DATE = EXPECTED_DELIVDATE = SALES_EMPLOYEE = SALES_GROUP = END_USER = ""


            try:
                inputTmpFile = os.path.join(TEMP_PATH, name)
                filePath.rename(inputTmpFile)

                dat_file_values = []
                with open(inputTmpFile, newline='') as games:
                    fileReader = csv.reader(games, delimiter='\t')
                    for element in fileReader:
                        dat_file_values = element[:]
                        print(element)

                today = date.today()

                missing_values = []

                CUST_PO_LINE = dat_file_values[0]
                if CUST_PO_LINE == "":
                    missing_values.append("Cust PO Line")

                CUST_PO = dat_file_values[1]
                if CUST_PO == "":
                    missing_values.append("Cust PO")


                PO_DATE = dat_file_values[2]
                if PO_DATE == "":
                    missing_values.append("PO Date")


                CUST_NO = dat_file_values[3]
                if CUST_NO == "":
                    missing_values.append("Cust No")


                SHIP_TO_NO = dat_file_values[4]
                if SHIP_TO_NO == "":
                    missing_values.append("Ship-To No")


                MATERIAL = dat_file_values[5]
                if MATERIAL == "":
                    missing_values.append("Material")


                QUANTITY = dat_file_values[7]
                if QUANTITY == "":
                    missing_values.append("Quantity")


                UOM = dat_file_values[8]
                if UOM == "":
                    missing_values.append("UOM")


                UNIT_PRICE = dat_file_values[9]
                if UNIT_PRICE == "":
                    missing_values.append("Unit Price")


                CURRENCY = dat_file_values[10]
                if CURRENCY == "":
                    missing_values.append("Currency")


                DELIVERY_DATE = dat_file_values[11]
                if DELIVERY_DATE == "":
                    missing_values.append("Delivery Date")


                EXPECTED_DELIVDATE = dat_file_values[13]
                if EXPECTED_DELIVDATE == "":
                    missing_values.append("Expected Delivery")


                SALES_EMPLOYEE = dat_file_values[17]
                if SALES_EMPLOYEE == "":
                    missing_values.append("Sales Employee")


                SALES_GROUP = dat_file_values[len(dat_file_values)-3]
                if SALES_GROUP == "":
                    missing_values.append("Sales Group")


                END_USER = dat_file_values[18]
                if END_USER == "":
                    missing_values.append("End User")


                if len(missing_values) > 0:
                    raise Exception


                # created root xml tag <DOCS850>
                DOCS850 = ET.Element('DOCS850')

                # Adding a subtag <DOC850> under <DOCS850>
                DOC850 = ET.SubElement(DOCS850, "DOC850")
                DOC850.set("AGENCY", "M")
                DOC850.set("VERSION", "XML10")
                DOC850.set("SET", "850")

                # Adding subtags under <DOC850>
                SENDER = ET.SubElement(DOC850, "SENDER").text = "RYOSAN"
                RECEIVER = ET.SubElement(DOC850, "RECEIVER").text = "EPSON"
                MSGSENDDATE = ET.SubElement(DOC850, "MSGSENDDATE").text = str(today).replace('-', '')

                # Adding <HEADER> under <DOC850>
                HEADER = ET.SubElement(DOC850, "HEADER")
                # Adding subtags under <HEADER>
                DOCPURPOSE = ET.SubElement(HEADER, "DOCPURPOSE").text = "00"
                ORDERTYPE = ET.SubElement(HEADER, "ORDERTYPE").text = "SA"
                ORDERNUM = ET.SubElement(HEADER, "ORDERNUM").text = CUST_PO

                #Adding DATEINFO under <HEADER>
                DATEINFO_HEADER = ET.SubElement(HEADER, "DATEINFO")
                #Adding DATETYPE and DATE under DATEINFO HEADER
                DATETYPE_dateinfo_header = ET.SubElement(DATEINFO_HEADER, "DATETYPE").text = "004"
                DATE_dateinfo_header = ET.SubElement(DATEINFO_HEADER, "DATE").text = PO_DATE

                #Adding SHIPINFO under HEADER
                SHIPINFO_HEADER = ET.SubElement(HEADER, "SHIPINFO")
                #Adding SHIPSCH under SHIP INFO header
                SHIPSCH_SHIPINFO_HEADER = ET.SubElement(SHIPINFO_HEADER, "SHIPSCH")
                SHIPDATE_SHIPSCH_SHIPINFO_HEADER = ET.SubElement(SHIPSCH_SHIPINFO_HEADER, "SHIPDATE").text = " "
                #Adding SHIPLOC under SHIPINFO
                SHIPLOC_SHIPINFO_HEADER = ET.SubElement(SHIPINFO_HEADER, "SHIPLOC")
                LOCTYPE_SHIPLOC_SHIPINFO_HEADER = ET.SubElement(SHIPLOC_SHIPINFO_HEADER, "LOCTYPE").text = "ST"
                ID_SHIPLOC_SHIPINFO_HEADER = ET.SubElement(SHIPLOC_SHIPINFO_HEADER, "ID").text = SHIP_TO_NO
                #Adding SHIPSPEC under SHIPINFO
                SHIPSPEC_SHIPINFO_HEADER = ET.SubElement(SHIPINFO_HEADER, "SHIPSPEC")
                SHIPMODE_SHIPSPEC_SHIPINFO_HEADER = ET.SubElement(SHIPSPEC_SHIPINFO_HEADER, "SHIPMODE").text = ""
                EXPQTY_SHIPSPEC_SHIPINFO_HEADER = ET.SubElement(SHIPSPEC_SHIPINFO_HEADER, "EXPQTY").text = ""



                # Adding <DETAIL> under <DOC824>
                DETAIL = ET.SubElement(DOC850, "DETAIL")

                # Adding <SEQUENCE> under <DETAIL>
                SEQUENCE_DETAIL = ET.SubElement(DETAIL, "SEQUENCE").text = CUST_PO_LINE.lstrip('0')
                STYLE = ET.SubElement(DETAIL, "STYLE").text = MATERIAL
                ATTRCODE1 = ET.SubElement(DETAIL, "ATTRCODE1").text = CUST_NO
                QUANTITY = ET.SubElement(DETAIL, "QUANTITY").text = QUANTITY
                QTYUOM = ET.SubElement(DETAIL, "QTYUOM")
                UNITPRICE = ET.SubElement(DETAIL, "UNITPRICE").text = UNIT_PRICE

                # Adding <SHIPINFO> under <DETAIL>
                SHIPINFO_DETAIL = ET.SubElement(DETAIL, "SHIPINFO")
                SHIPINFO_DETAIL.set("PARENT", "DETAIL")
                # Adding SHIPSCH under <SHIPINFO>
                SHIPSCH_SHIPINFO_DETAIL = ET.SubElement(SHIPINFO_DETAIL, "SHIPSCH")
                SHIPDATE_SHIPSCH_SHIPINFO_DETAIL = ET.SubElement(SHIPSCH_SHIPINFO_DETAIL, "SHIPDATE").text = DELIVERY_DATE
                EXPDELVY_SHIPSCH_SHIPINFO_DETAIL = ET.SubElement(SHIPSCH_SHIPINFO_DETAIL, "EXPDELVY").text = EXPECTED_DELIVDATE
                #Adding SHIPLOC under SHIPINFO
                SHIPLOC_SHIPINFO_DETAIL = ET.SubElement(SHIPINFO_DETAIL, "SHIPLOC")
                LOCTYPE_SHIPLOC_SHIPINFO_DETAIL = ET.SubElement(SHIPLOC_SHIPINFO_DETAIL, "LOCTYPE").text = "ST"
                ID_SHIPLOC_SHIPINFO_DETAIL = ET.SubElement(SHIPLOC_SHIPINFO_DETAIL, "ID").text = SHIP_TO_NO
                #Adding SHIPSPEC under SHIPINFO
                SHIPSPEC_SHIPINFO_DETAIL = ET.SubElement(SHIPINFO_DETAIL, "SHIPSPEC")
                SHIPMODE_SHIPSPEC_SHIPINFO_DETAIL = ET.SubElement(SHIPSPEC_SHIPINFO_DETAIL, "SHIPMODE")
                EXPQTY_SHIPSPEC_SHIPINFO_DETAIL = ET.SubElement(SHIPSPEC_SHIPINFO_DETAIL, "EXPQTY")

                AddExtra(DETAIL, "DETAIL", "999VP", "Material", MATERIAL)
                AddExtra(DETAIL, "DETAIL", "999CUR", "CURRENCY", CURRENCY)
                AddExtra(DETAIL, "DETAIL", "999SE", "Sales Employee", SALES_EMPLOYEE)
                AddExtra(DETAIL, "DETAIL", "999SG", "Sales Group", SALES_GROUP)
                # AddExtra(DETAIL, "DETAIL", "98EC", "End User", END_USER)

                # Adding <PARTY> 98EC under element
                PARTY = ET.SubElement(DETAIL, "PARTY")
                PARTY.set("PARENT", "DETAIL")
                CODE = ET.SubElement(PARTY, "CODE").text = "98EC"
                LABEL = ET.SubElement(PARTY, "LABEL").text = "End User"
                ID = ET.SubElement(PARTY, "ID").text = END_USER


                #Add EXTRA DOC850
                from datetime import datetime
                now = datetime.now()
                current_time = now.strftime("%H:%M")

                AddExtra(DOC850, "DOC850", "999ISADT", "EDI Interchange Date", str(today).replace('-', ''))
                AddExtra(DOC850, "DOC850", "999ISATM", "EDI Interchange Time", current_time.replace(':',''))
                # AddExtra(DOC850, "DOC850", "98BY", "Cust No", CUST_NO)

                # Adding <PARTY> 98BY under element
                PARTY = ET.SubElement(DOC850, "PARTY")
                PARTY.set("PARENT", "DOC850")
                CODE = ET.SubElement(PARTY, "CODE").text = "98BY"
                LABEL = ET.SubElement(PARTY, "LABEL").text = "Cust No"
                ID = ET.SubElement(PARTY, "ID").text = CUST_NO



                # Converting the xml data to byte object,
                # for allowing flushing data to file
                # stream
                tree = ET.ElementTree(DOCS850)

                f = BytesIO()
                tree.write(f, encoding='utf-8', xml_declaration=True)
                # tree.write("hello-file.xml", encoding='utf-8', xml_declaration=True, method='xml')
                # b_xml = ET.tostring(DOCS850)
                b_xml = f.getvalue()
                # print(b_xml)


                # save the data to XML file
                outputTmpFile = os.path.join(TEMP_PATH, filenameWithoutExtension + TMP_FILE_EXT)
                with open(outputTmpFile, "wb") as f:
                    f.write(b_xml)

                # move output tmp file to output folder
                dstFilePath = os.path.join(OUTPUT_PATH, filenameWithoutExtension + OUTPUT_FILE_EXT)
                CopyFile(outputTmpFile, dstFilePath)
                PrintMessage("The output XML824 file was generated in the output folder, " + dstFilePath)

                # archive output xml file
                os.remove(outputTmpFile)
                archivePath = GetCurrentDatePath(ARCHIVE_PATH)
                dstFilePath = os.path.join(archivePath, filenameWithoutExtension + OUTPUT_FILE_EXT)
                # MoveFile(outputTmpFile, dstFilePath)
                # PrintMessage("The output XML824 file was moved to the archive folder, " + archivePath)

                # archive source edi file
                dstFilePath = os.path.join(archivePath, name)
                MoveFile(inputTmpFile, dstFilePath)
                PrintMessage("The source EDI355 file was moved to the archive folder, " + dstFilePath)

                PrintMessage(filePath.name + " was translated successfully!")

            except Exception as e:
                PrintMessage("Exception Message: " + str(e))
                xfailedFilePath = os.path.join(XFAILED_PATH, name)
                MoveFile(inputTmpFile, xfailedFilePath)
                PrintError(filePath.name + " was failed in translation!")

                ####Sending failure email

                message = EmailMessage()
                sender = "example@gmail.com"
                recipient = "example@gmail.com, example@tradelinkone.com"
                message['From'] = sender
                message['To'] = recipient
                message['Subject'] = 'TLT ALERT: Epson Ryosan PO to XML850 translation failure'
                #Email Body
                body = "Dear Recipients," + "\n" + "\n"
                body+="The PO message is failed to translate due to below reason:" + "\n" + "\n"
                body+="PO Number: " + CUST_PO + "\n" + "\n"
                missing_columns = ""
                for missing_column in missing_values:
                    missing_columns += missing_column + " "
                body+=f"Reason of failure: Mandatory Field, {missing_columns} data is missing" +"\n"
                body+="\n"
                body+=f"Filename: {filenameWithoutExtension}"


                message.set_content(body)
                # mime_type, _ = mimetypes.guess_type('something.pdf')
                # mime_type, mime_subtype = mime_type.split('/')
                mime_type = 'application'
                mime_subtype = 'dat'
                with open(f'{xfailedFilePath}', 'rb') as file:
                    message.add_attachment(file.read(),
                                           maintype=mime_type,
                                           subtype=mime_subtype,
                                           filename=f'{filenameWithoutExtension}.dat')
                print(message)
                mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
                mail_server.set_debuglevel(1)
                mail_server.login("example@gmail.com", 'password')
                mail_server.send_message(message)
                mail_server.quit()
# end for

PrintMessage("End US Custom EDI355 translation...")
keyboard.press_and_release('escape')
