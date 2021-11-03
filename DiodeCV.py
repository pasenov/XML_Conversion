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

def DiodeCVFunc(folderIn, folderOut, fileXLSX, fileIn):


    fileName = fileIn.split('.')[0]
    fileNew = fileName + '_new.txt'
    fileOut = fileName + '.xml'

    with open(folderIn/fileIn, 'r') as fin:
            dataIn = fin.read().splitlines(True)
    with open(folderIn/fileNew, 'w') as fout:
            fout.writelines(dataIn[1:])
            

    with open(folderIn/fileNew) as f:
            lines = (line for line in f if not line.startswith('#'))
            FData = np.loadtxt(lines, delimiter='\t', skiprows=0)
            
    voltageArr = FData[:, 0]
    capacitanceArr = FData[:, 1]
    resistanceArr = FData[:, 2]
    temperatureArr = FData[:, 3]
    airTemperatureArr = FData[:, 4]
    RHArr = FData[:, 5]

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
            
    if ((n3 == '2-S') or (n3 == '2S')):
            kp1 = '2S'
            n3a = '2-S'
    elif (n3 == 'PSP'):
            kp1 = 'PSP'
            n3a = 'PSP'
    elif (n3 == 'PSS'):
            kp1 = 'PSS'
            n3a = 'PSS'
    kp = kp1 + ' Halfmoon ' + n5
    nL = n1 + '_' + n2 + '_' + n3a + '_' + n4 + '_' + n5a

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
    if (n8 == 'DiodeCV'):
            struct = 'DIODE_HALF'
            waitTime = '0.500'
            extTabNam = 'TEST_SENSOR_CV'
            extTabNam2 = 'HALFMOON_CV_PAR'
            nameTest = 'Tracker Halfmoon CV Test'
            nameTest2 = 'Tracker Halfmoon CV Parameters'
            versionMeas = 'CV_measurement-004'

            
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
    kindOfHMConfigID = ET.SubElement(data, "KIND_OF_HM_CONFIG_ID").text = "Not Used"

    procedureType = ET.SubElement(data, "PROCEDURE_TYPE").text = 'DiodeCV'
    fileName = ET.SubElement(data, "FILE_NAME").text = fileIn
    equipment = ET.SubElement(data, "EQUIPMENT").text = "PQC_HM_POSITION " + flutePos
    waitingTimeS = ET.SubElement(data, "WAITING_TIME_S").text = waitTime
    tempSetDegC = ET.SubElement(data, "TEMP_SET_DEGC").text = '20.'
    avTempDegC = ET.SubElement(data, "AV_TEMP_DEGC").text = '20.000'
    corrOpenPfrd = ET.SubElement(data, "CORR_OPEN_PFRD").text = '0.2'

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
            capacitanceNum = (1E12)*capacitanceArr[i]
            capacitanceNum = round(capacitanceNum, 3)
            capacitance = str(capacitanceNum)
            resistanceNum = (1E-6)*(resistanceArr[i])
            resistanceNum = round(resistanceNum, 3)
            resistance = str(resistanceNum)
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
            capctncPfrd = ET.SubElement(data2, "CAPCTNC_PFRD").text = capacitance
            resstncMohm = ET.SubElement(data2, "RESSTNC_MOHM").text = resistance
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
    Vd = sheet["B16"].value
    Vd = round(Vd, 3)
    VdV = ET.SubElement(data3, "VD_V").text = str(Vd)
    CapMin = (min(capacitanceArr))*(1E12)
    CapMin = round(CapMin, 3)
    CminPfrd = ET.SubElement(data3, "CMIN_PFRD").text = str(CapMin)
    tox = (1E-3)*3.9*8.854*16900/CapMin
    tox = round(tox, 3)
    DUm = ET.SubElement(data3, "D_UM").text = str(tox)
    capacitanceArrSq = np.square(capacitanceArr)
    capacitanceArrSqRec = np.reciprocal(capacitanceArrSq)
    dy = np.diff(capacitanceArrSqRec,1)
    dx = np.diff(voltageArr,1)
    yfirst = (-1)*dy/dx
    Na = (1E-18)*2/(1.6*(1E-19)*8.854*(1E-12)*11.68*2.5*2.5*(1E-6)*2.5*2.5*(1E-6)*yfirst[0])
    Na = round(Na, 3)
    NA1 = ET.SubElement(data3, "NA").text = str(Na)
    Rho = 290*(1E-6)*290*(1E-6)/(10*2*8.854*(1E-12)*11.68*483.78*(1E-4)*Vd)
    Rho = round(Rho, 3)
    RhoKohmcm = ET.SubElement(data3, "RHO_KOHMCM").text = str(Rho)



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
