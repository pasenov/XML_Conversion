import tkinter
from tkinter import *
from tkinter import filedialog,StringVar
from tkinter.ttk import Frame, Button, Style
from tkinter import Tk
from tkinter.filedialog import askdirectory
import tkinter.font as font
import xlsxwriter
import sys
import os
import os.path
from os import path
import csv
import pathlib
from decimal import Decimal
from pathlib import Path
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import numpy as np
import xml.etree.cElementTree as ET
import xml.dom.minidom
import re
from datetime import datetime, timedelta
import openpyxl
from pathlib import Path


def polyCBKRFunc(folderIn, folderOut, fileXLSX, fileIn):


    fileName = fileIn.split('.')[0]
    fileNew = fileName + '_new.txt'
    fileOut = fileName + '.xml'

    with open(folderIn/fileIn) as fFirst:
            first_line = fFirst.readline()
            
    ROhm_value = first_line.split('\t')[1]

    with open(folderIn/fileIn, 'r') as fin:
            dataIn = fin.read().splitlines(True)
    with open(folderIn/fileNew, 'w') as fout:
            fout.writelines(dataIn[2:])
            

    with open(folderIn/fileNew) as f:
            lines = (line for line in f if not line.startswith('#'))
            FData = np.loadtxt(lines, delimiter='\t', skiprows=0)
            
    voltageArr = FData[:, 0]
    currentArr = FData[:, 1]
    temperatureArr = FData[:, 2]
    airTemperatureArr = FData[:, 3]
    RHArr = FData[:, 4]

    dayData = fileIn.split('_')[9]
    if (int(dayData) < 10):
            dayData = '0' + dayData
    monthData = fileIn.split('_')[10]
    if (int(monthData) < 10):
            monthData = '0' + monthData
    yearData = fileIn.split('_')[11]
    a14 = fileIn.split('_')[12]
    hourData = a14.split('h')[0]
    b2 = a14.split('h')[1]
    minuteData = b2.split('m')[0]
    if (int(minuteData) < 10):
            minuteData = '0' + minuteData
    m2 = b2.split('m')[1]
    secondData = m2.split('s')[0]
    if (int(secondData) < 10):
            secondData = '0' + secondData
    runBeginTimestamp_value = yearData + '-' + monthData + '-' + dayData + ' ' + hourData + ':' + minuteData + ':' + secondData

    datetime_original = datetime(year=int(yearData), month=int(monthData), day=int(dayData), hour = int(hourData), minute = int(minuteData), second = int(secondData))
    time_delta = timedelta(hours=0, minutes=0, seconds=1, microseconds=0)
    datetime_new = datetime_original + time_delta

    n1 = fileIn.split('_')[1]
    n2 = fileIn.split('_')[2]
    n3 = fileIn.split('_')[3]
    n4 = fileIn.split('_')[4]
    n5 = fileIn.split('_')[5]
    if (n5 == 'E'):
            n5a = n5 + 'E'
    elif (n5 == 'W'):
            n5a = n5 + 'W'
    nL = n1 + '_' + n2 + '_' + n3 + '_' + n4 + '_' + n5a

    if ((n3 == '2-S') or (n3 == '2S')):
            kp1 = '2S'
    elif (n3 == 'PSP'):
            kp1 = 'PSP'
    elif (n3 == 'PSS'):
            kp1 = 'PSS'
    kp = kp1 + ' Halfmoon ' + n5

    n7 = fileIn.split('_')[7]
    if (n7 == 'L'):
            pos = 'Left'
    elif (n7 == 'R'):
            pos = 'Right'
            
    n6 = fileIn.split('_')[6]
    if (n6 == 'flute1'):
            flute = 'PQC1'
            flutePos = '1'
    elif (n6 == 'flute2'):
            flute = 'PQC2'
            flutePos = '2'
    elif (n6 == 'flute3'):
            flute = 'PQC3'
            flutePos = '3'
    elif (n6 == 'flute4'):
            flute = 'PQC4'
            flutePos = '4'
            
    n8 = fileIn.split('_')[8]
    if (n8 == 'polyCBKR'):
            struct = 'CBKR_POLY'
            waitTime = '0.200'
            extTabNam = 'TEST_SENSOR_IV'
            extTabNam2 = 'HALFMOON_IV_PAR'
            nameTest = 'Tracker Halfmoon IV Test'
            nameTest2 = 'Tracker Halfmoon IV Parameters'
            versionMeas = 'IV_measurement-004'


            
    m_encoding = 'UTF-8'
    m_standalone = 'yes'

    root = ET.Element("ROOT")
    header = ET.SubElement(root, "HEADER")
    type1 = ET.SubElement(header, "TYPE")
    extensionTableName = ET.SubElement(type1, "EXTENSION_TABLE_NAME").text = "HALFMOON_METADATA"
    name = ET.SubElement(type1, "NAME").text = "Tracker Halfmoon Metadata"
    run = ET.SubElement(header, "RUN", mode ="SEQUENCE_NUMBER", sequence ="TRK_OT_RUN_SEQ")

    runType = ET.SubElement(run, "RUN_TYPE").text = "PQC"
    location = ET.SubElement(run, "LOCATION").text = "Perugia"
    initiatedByUser = ET.SubElement(run, "INITIATED_BY_USER").text = "Patrick Asenov"
    runBeginTimestamp = ET.SubElement(run, "RUN_BEGIN_TIMESTAMP").text = runBeginTimestamp_value
    commentDescription = ET.SubElement(run, "COMMENT_DESCRIPTION").text = "\n\n   "

    data_set = ET.SubElement(root, "DATA_SET")
    commentDescription2 = ET.SubElement(data_set, "COMMENT_DESCRIPTION").text = "Metadata with flute and structure"
    version = ET.SubElement(data_set, "VERSION").text = "v2"

    part = ET.SubElement(data_set, "PART")
    nameLabel = ET.SubElement(part, "NAME_LABEL").text = nL
    kindOfPart = ET.SubElement(part, "KIND_OF_PART").text = kp

    data = ET.SubElement(data_set, "DATA")
    kindOfHMSetID = ET.SubElement(data, "KIND_OF_HM_SET_ID").text = pos
    kindOfHMFluteID = ET.SubElement(data, "KIND_OF_HM_FLUTE_ID").text = flute
    kindOfHMStructID = ET.SubElement(data, "KIND_OF_HM_STRUCT_ID").text = struct
    kindOfHMConfigID = ET.SubElement(data, "KIND_OF_HM_CONFIG_ID").text = "Standard"

    procedureType = ET.SubElement(data, "PROCEDURE_TYPE").text = 'CBKR-polySi'
    fileName = ET.SubElement(data, "FILE_NAME").text = fileIn
    equipment = ET.SubElement(data, "EQUIPMENT").text = "PQC_HM_POSITION " + flutePos
    waitingTimeS = ET.SubElement(data, "WAITING_TIME_S").text = waitTime
    tempSetDegC = ET.SubElement(data, "TEMP_SET_DEGC").text = '20.'
    avTempDegC = ET.SubElement(data, "AV_TEMP_DEGC").text = '20.000'

    childDataSet = ET.SubElement(data_set, "CHILD_DATA_SET")
    header2 = ET.SubElement(childDataSet, "HEADER")
    type2 = ET.SubElement(header2, "TYPE")
    extensionTableName = ET.SubElement(type2, "EXTENSION_TABLE_NAME").text = extTabNam
    name2 = ET.SubElement(type2, "NAME").text = nameTest
    dataset2 = ET.SubElement(childDataSet, "DATA_SET")
    commentDescription3 = ET.SubElement(dataset2, "COMMENT_DESCRIPTION").text = "Test"
    version2 = ET.SubElement(dataset2, "VERSION").text = versionMeas
    partnew2 = ET.SubElement(dataset2, "PART")
    nameLabel2 = ET.SubElement(partnew2, "NAME_LABEL").text = nL
    kindOfPart2 = ET.SubElement(partnew2, "KIND_OF_PART").text = kp

    for i in range(voltageArr.size):
            voltageNum = voltageArr[i]
            voltage = str(voltageNum)
            currentNum = (1E9)*currentArr[i]
            currentNum = round(currentNum, 3)
            current = str(currentNum)
            temperatureNum = temperatureArr[i]
            temperature = str(temperatureNum)
            airTemperatureNum = airTemperatureArr[i]
            airTemperature = str(airTemperatureNum)
            RHNum = RHArr[i]
            RH = str(RHNum)


            datetime_new = datetime_new + time_delta
            data2 = ET.SubElement(dataset2, "DATA")
            time = ET.SubElement(data2, "TIME").text = str(datetime_new)
            volts = ET.SubElement(data2, "VOLTS").text = voltage
            currntNamp = ET.SubElement(data2, "CURRNT_NAMP").text = current
            tempDegC = ET.SubElement(data2, "TEMP_DEGC").text = temperature
            airTempDegC = ET.SubElement(data2, "AIR_TEMP_DEGC").text = airTemperature
            RHPrcnt = ET.SubElement(data2, "RH_PRCNT").text = RH


    childDataSet2 = ET.SubElement(dataset2, "CHILD_DATA_SET")
    header3 = ET.SubElement(childDataSet2, "HEADER")
    type3 = ET.SubElement(header3, "TYPE")
    extensionTableName2 = ET.SubElement(type3, "EXTENSION_TABLE_NAME").text = extTabNam2
    name3 = ET.SubElement(type3, "NAME").text = nameTest2
    dataset3 = ET.SubElement(childDataSet2, "DATA_SET")
    commentDescription4 = ET.SubElement(dataset3, "COMMENT_DESCRIPTION").text = "Test"
    version3 = ET.SubElement(dataset3, "VERSION").text = versionMeas
    partnew3 = ET.SubElement(dataset3, "PART")
    nameLabel3 = ET.SubElement(partnew3, "NAME_LABEL").text = nL
    kindOfPart3 = ET.SubElement(partnew3, "KIND_OF_PART").text = kp
    data3 = ET.SubElement(dataset3, "DATA")

    wb_obj = openpyxl.load_workbook(folderIn/fileXLSX) 
    sheet = wb_obj.active
    RcOhm = ET.SubElement(data3, "RC_OHM").text = ROhm_value
    ROhm = ET.SubElement(data3, "R_OHM").text = ROhm_value

    dom = xml.dom.minidom.parseString(ET.tostring(root))
    xml_string = dom.toprettyxml()
    part1, part2 = xml_string.split('?>')

    with open(folderIn/fileIn, 'r') as fin1:
            fin1.close()

    with open(folderOut/fileOut, 'w') as fout1:
            fout1.write(part1 + ' encoding=\"{}\"'.format(m_encoding) + ' standalone=\"{}\"?>\n'.format(m_standalone)  + part2)
            fout1.close()

    os.remove(folderIn/fileNew)

    return(folderOut/fileOut)
