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
            kp1 = 'PS-p'
            n3a = 'PSP'
    elif (n3 == 'PSS'):
            kp1 = 'PS-s'
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

    Vd = round(Vd, 0)

    if Vd == 100:
        Rho = 8405
    elif Vd == 101:
        Rho = 8322
    elif Vd == 102:
        Rho = 8240
    elif Vd == 103:
        Rho = 8160
    elif Vd == 104:
        Rho = 8082
    elif Vd == 105:
        Rho = 8005
    elif Vd == 106:
        Rho = 7929
    elif Vd == 107:
        Rho = 7855
    elif Vd == 108:
        Rho = 7782
    elif Vd == 109:
        Rho = 7711
    elif Vd == 110:
        Rho = 7641
    elif Vd == 111:
        Rho = 7572
    elif Vd == 112:
        Rho = 7504
    elif Vd == 113:
        Rho = 7438
    elif Vd == 114:
        Rho = 7373
    elif Vd == 115:
        Rho = 7309
    elif Vd == 116:
        Rho = 7246
    elif Vd == 117:
        Rho = 7184
    elif Vd == 118:
        Rho = 7123
    elif Vd == 119:
        Rho = 7063
    elif Vd == 120:
        Rho = 7004
    elif Vd == 121:
        Rho = 6946
    elif Vd == 122:
        Rho = 6889
    elif Vd == 123:
        Rho = 6833
    elif Vd == 124:
        Rho = 6778
    elif Vd == 125:
        Rho = 6724
    elif Vd == 126:
        Rho = 6670
    elif Vd == 127:
        Rho = 6618
    elif Vd == 128:
        Rho = 6566
    elif Vd == 129:
        Rho = 6515
    elif Vd == 130:
        Rho = 6465
    elif Vd == 131:
        Rho = 6416
    elif Vd == 132:
        Rho = 6367
    elif Vd == 133:
        Rho = 6319
    elif Vd == 134:
        Rho = 6272
    elif Vd == 135:
        Rho = 6226
    elif Vd == 136:
        Rho = 6180
    elif Vd == 137:
        Rho = 6135
    elif Vd == 138:
        Rho = 6090
    elif Vd == 139:
        Rho = 6047
    elif Vd == 140:
        Rho = 6003
    elif Vd == 141:
        Rho = 5961
    elif Vd == 142:
        Rho = 5919
    elif Vd == 143:
        Rho = 5878
    elif Vd == 144:
        Rho = 5837
    elif Vd == 145:
        Rho = 5796
    elif Vd == 146:
        Rho = 5757
    elif Vd == 147:
        Rho = 5718
    elif Vd == 148:
        Rho = 5679
    elif Vd == 149:
        Rho = 5641
    elif Vd == 150:
        Rho = 5603
    elif Vd == 151:
        Rho = 5566
    elif Vd == 152:
        Rho = 5529
    elif Vd == 153:
        Rho = 5493
    elif Vd == 154:
        Rho = 5458
    elif Vd == 155:
        Rho = 5422
    elif Vd == 156:
        Rho = 5388
    elif Vd == 157:
        Rho = 5353
    elif Vd == 158:
        Rho = 5320
    elif Vd == 159:
        Rho = 5286
    elif Vd == 160:
        Rho = 5253
    elif Vd == 161:
        Rho = 5220
    elif Vd == 162:
        Rho = 5188
    elif Vd == 163:
        Rho = 5157
    elif Vd == 164:
        Rho = 5125
    elif Vd == 165:
        Rho = 5094
    elif Vd == 166:
        Rho = 5063
    elif Vd == 167:
        Rho = 5033
    elif Vd == 168:
        Rho = 5003
    elif Vd == 169:
        Rho = 4973
    elif Vd == 170:
        Rho = 4933
    elif Vd == 171:
        Rho = 4915
    elif Vd == 172:
        Rho = 4887
    elif Vd == 173:
        Rho = 4858
    elif Vd == 174:
        Rho = 4830
    elif Vd == 175:
        Rho = 4803
    elif Vd == 176:
        Rho = 4775
    elif Vd == 177:
        Rho = 4749
    elif Vd == 178:
        Rho = 4722
    elif Vd == 179:
        Rho = 4695
    elif Vd == 180:
        Rho = 4669
    elif Vd == 181:
        Rho = 4644
    elif Vd == 182:
        Rho = 4618
    elif Vd == 183:
        Rho = 4593
    elif Vd == 184:
        Rho = 4568
    elif Vd == 185:
        Rho = 4543
    elif Vd == 186:
        Rho = 4519
    elif Vd == 187:
        Rho = 4495
    elif Vd == 188:
        Rho = 4471
    elif Vd == 189:
        Rho = 4447
    elif Vd == 190:
        Rho = 4424
    elif Vd == 191:
        Rho = 4400
    elif Vd == 192:
        Rho = 4378
    elif Vd == 193:
        Rho = 4355
    elif Vd == 194:
        Rho = 4332
    elif Vd == 195:
        Rho = 4310
    elif Vd == 196:
        Rho = 4288
    elif Vd == 197:
        Rho = 4266
    elif Vd == 198:
        Rho = 4245
    elif Vd == 199:
        Rho = 4224 
    elif Vd == 200:
        Rho = 4202
    elif Vd == 201:
        Rho = 4182
    elif Vd == 202:
        Rho = 4161
    elif Vd == 203:
        Rho = 4140
    elif Vd == 204:
        Rho = 4120
    elif Vd == 205:
        Rho = 4100
    elif Vd == 206:
        Rho = 4080
    elif Vd == 207:
        Rho = 4060
    elif Vd == 208:
        Rho = 4041
    elif Vd == 209:
        Rho = 4021
    elif Vd == 210:
        Rho = 4002
    elif Vd == 211:
        Rho = 3983
    elif Vd == 212:
        Rho = 3965
    elif Vd == 213:
        Rho = 3946
    elif Vd == 214:
        Rho = 3928
    elif Vd == 215:
        Rho = 3909
    elif Vd == 216:
        Rho = 3891
    elif Vd == 217:
        Rho = 3873
    elif Vd == 218:
        Rho = 3855
    elif Vd == 219:
        Rho = 3838
    elif Vd == 220:
        Rho = 3820
    elif Vd == 221:
        Rho = 3803
    elif Vd == 222:
        Rho = 3786
    elif Vd == 223:
        Rho = 3769
    elif Vd == 224:
        Rho = 3752
    elif Vd == 225:
        Rho = 3735
    elif Vd == 226:
        Rho = 3719
    elif Vd == 227:
        Rho = 3703
    elif Vd == 228:
        Rho = 3686
    elif Vd == 229:
        Rho = 3670
    elif Vd == 230:
        Rho = 3654
    elif Vd == 231:
        Rho = 3638
    elif Vd == 232:
        Rho = 3623
    elif Vd == 233:
        Rho = 3607
    elif Vd == 234:
        Rho = 3592
    elif Vd == 235:
        Rho = 3577
    elif Vd == 236:
        Rho = 3561
    elif Vd == 237:
        Rho = 3546
    elif Vd == 238:
        Rho = 3531
    elif Vd == 239:
        Rho = 3517
    elif Vd == 240:
        Rho = 3502
    elif Vd == 241:
        Rho = 3488
    elif Vd == 242:
        Rho = 3473
    elif Vd == 243:
        Rho = 3459
    elif Vd == 244:
        Rho = 3445
    elif Vd == 245:
        Rho = 3431
    elif Vd == 246:
        Rho = 3417
    elif Vd == 247:
        Rho = 3403
    elif Vd == 248:
        Rho = 3389
    elif Vd == 249:
        Rho = 3375
    elif Vd == 250:
        Rho = 3362
    elif Vd == 251:
        Rho = 3349
    elif Vd == 252:
        Rho = 3335
    elif Vd == 253:
        Rho = 3322
    elif Vd == 254:
        Rho = 3309
    elif Vd == 255:
        Rho = 3296
    elif Vd == 256:
        Rho = 3283
    elif Vd == 257:
        Rho = 3270
    elif Vd == 258:
        Rho = 3258
    elif Vd == 259:
        Rho = 3245
    elif Vd == 260:
        Rho = 3233
    elif Vd == 261:
        Rho = 3220
    elif Vd == 262:
        Rho = 3208
    elif Vd == 263:
        Rho = 3196
    elif Vd == 264:
        Rho = 3184
    elif Vd == 265:
        Rho = 3172
    elif Vd == 266:
        Rho = 3160
    elif Vd == 267:
        Rho = 3148
    elif Vd == 268:
        Rho = 3136
    elif Vd == 269:
        Rho = 3124
    elif Vd == 270:
        Rho = 3113
    elif Vd == 271:
        Rho = 3101
    elif Vd == 272:
        Rho = 3090
    elif Vd == 273:
        Rho = 3079
    elif Vd == 274:
        Rho = 3067
    elif Vd == 275:
        Rho = 3056
    elif Vd == 276:
        Rho = 3045
    elif Vd == 277:
        Rho = 3034
    elif Vd == 278:
        Rho = 3023
    elif Vd == 279:
        Rho = 3013
    elif Vd == 280:
        Rho = 3002
    elif Vd == 281:
        Rho = 2991
    elif Vd == 282:
        Rho = 2980
    elif Vd == 283:
        Rho = 2970
    elif Vd == 284:
        Rho = 2959
    elif Vd == 285:
        Rho = 2949
    elif Vd == 286:
        Rho = 2939
    elif Vd == 287:
        Rho = 2929
    elif Vd == 288:
        Rho = 2918
    elif Vd == 289:
        Rho = 2908
    elif Vd == 290:
        Rho = 2898
    elif Vd == 291:
        Rho = 2888
    elif Vd == 292:
        Rho = 2878
    elif Vd == 293:
        Rho = 2869
    elif Vd == 294:
        Rho = 2859
    elif Vd == 295:
        Rho = 2849
    elif Vd == 296:
        Rho = 2839
    elif Vd == 297:
        Rho = 2830
    elif Vd == 298:
        Rho = 2820
    elif Vd == 299:
        Rho = 2811
    elif Vd == 300:
        Rho = 2802
    elif Vd == 301:
        Rho = 2792
    elif Vd == 302:
        Rho = 2783
    elif Vd == 303:
        Rho = 2773
    elif Vd == 304:
        Rho = 2765
    elif Vd == 305:
        Rho = 2756
    elif Vd == 306:
        Rho = 2747
    elif Vd == 307:
        Rho = 2738
    elif Vd == 308:
        Rho = 2729
    elif Vd == 309:
        Rho = 2720
    elif Vd == 310:
        Rho = 2711
    elif Vd == 311:
        Rho = 2703
    elif Vd == 312:
        Rho = 2694
    elif Vd == 313:
        Rho = 2685
    elif Vd == 314:
        Rho = 2677
    elif Vd == 315:
        Rho = 2668
    elif Vd == 316:
        Rho = 2660
    elif Vd == 317:
        Rho = 2651
    elif Vd == 318:
        Rho = 2643
    elif Vd == 319:
        Rho = 2635
    elif Vd == 320:
        Rho = 2627
    elif Vd == 321:
        Rho = 2618
    elif Vd == 322:
        Rho = 2610
    elif Vd == 323:
        Rho = 2602
    elif Vd == 324:
        Rho = 2594
    elif Vd == 325:
        Rho = 2586
    elif Vd == 326:
        Rho = 2578
    elif Vd == 327:
        Rho = 2570
    elif Vd == 328:
        Rho = 2562
    elif Vd == 329:
        Rho = 2555
    elif Vd == 330:
        Rho = 2547
    elif Vd == 331:
        Rho = 2539
    elif Vd == 332:
        Rho = 2532
    elif Vd == 333:
        Rho = 2524
    elif Vd == 334:
        Rho = 2516
    elif Vd == 335:
        Rho = 2509
    elif Vd == 336:
        Rho = 2501
    elif Vd == 337:
        Rho = 2494
    elif Vd == 338:
        Rho = 2487
    elif Vd == 339:
        Rho = 2479
    elif Vd == 340:
        Rho = 2472
    elif Vd == 341:
        Rho = 2465
    elif Vd == 342:
        Rho = 2458
    elif Vd == 343:
        Rho = 2450
    elif Vd == 344:
        Rho = 2443
    elif Vd == 345:
        Rho = 2436
    elif Vd == 346:
        Rho = 2429
    elif Vd == 347:
        Rho = 2422
    elif Vd == 348:
        Rho = 2415
    elif Vd == 349:
        Rho = 2408
    elif Vd == 350:
        Rho = 2401
    elif Vd == 351:
        Rho = 2395
    elif Vd == 352:
        Rho = 2388
    elif Vd == 353:
        Rho = 2381
    elif Vd == 354:
        Rho = 2374
    elif Vd == 355:
        Rho = 2368
    elif Vd == 356:
        Rho = 2361
    elif Vd == 357:
        Rho = 2354
    elif Vd == 358:
        Rho = 2348
    elif Vd == 359:
        Rho = 2341
    elif Vd == 360:
        Rho = 2335
    elif Vd == 361:
        Rho = 2328
    elif Vd == 362:
        Rho = 2322
    elif Vd == 363:
        Rho = 2315
    elif Vd == 364:
        Rho = 2309
    elif Vd == 365:
        Rho = 2303
    elif Vd == 366:
        Rho = 2296
    elif Vd == 367:
        Rho = 2290
    elif Vd == 368:
        Rho = 2284
    elif Vd == 369:
        Rho = 2278
    elif Vd == 370:
        Rho = 2272
    elif Vd == 371:
        Rho = 2265
    elif Vd == 372:
        Rho = 2259
    elif Vd == 373:
        Rho = 2253
    elif Vd == 374:
        Rho = 2247
    elif Vd == 375:
        Rho = 2241
    elif Vd == 376:
        Rho = 2235
    elif Vd == 377:
        Rho = 2229
    elif Vd == 378:
        Rho = 2224
    elif Vd == 379:
        Rho = 2218
    elif Vd == 380:
        Rho = 2212
    elif Vd == 381:
        Rho = 2206
    elif Vd == 382:
        Rho = 2200
    elif Vd == 383:
        Rho = 2195
    elif Vd == 384:
        Rho = 2189
    elif Vd == 385:
        Rho = 2183
    elif Vd == 386:
        Rho = 2177
    elif Vd == 387:
        Rho = 2172
    elif Vd == 388:
        Rho = 2166
    elif Vd == 389:
        Rho = 2161
    elif Vd == 390:
        Rho = 2155
    elif Vd == 391:
        Rho = 2150
    elif Vd == 392:
        Rho = 2144
    elif Vd == 393:
        Rho = 2139
    elif Vd == 394:
        Rho = 2133
    elif Vd == 395:
        Rho = 2128
    elif Vd == 396:
        Rho = 2122
    elif Vd == 397:
        Rho = 2117
    elif Vd == 398:
        Rho = 2112
    elif Vd == 399:
        Rho = 2107
    elif Vd == 400:
        Rho = 2101
    else:
        Rho = -1000

    Rho = Rho/1000
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
