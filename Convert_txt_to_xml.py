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
import zipfile
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
from capacitor import capacitorFunc
from FET import FETFunc
from MOS import MOSFunc
from strip import stripFunc
from Poly import PolyFunc
from pstop import pstopFunc
from DielectricBreakdown import DielectricBreakdownFunc
from GCD import GCDFunc
from LinewidthStrip import LinewidthStripFunc
from LinewidthPolyMeander import LinewidthPolyMeanderFunc
from Linewidthpstop import LinewidthpstopFunc
from BulkCross import BulkCrossFunc
from DiodeCV import DiodeCVFunc
from DiodeIV import DiodeIVFunc
from MetalClover import MetalCloverFunc
from pBridge import pBridgeFunc
from pCross import pCrossFunc
from MetalMeanderChain import MetalMeanderChainFunc
from GCD05 import GCD05Func
from stripCBKR import stripCBKRFunc
from polyCBKR import polyCBKRFunc
from nChain import nChainFunc
from pChain import pChainFunc
from polyChain import polyChainFunc

rootdir = os.getcwd()

dir_list = next(os.walk('.'))[1]
for folder in dir_list:
    folderName = str(folder)
    if (folderName != '__pycache__'):
        oldPath = Path(folderName)
        newPathString = 'Converted_' + folderName
        
        for file in os.listdir(oldPath):
            fileEx = os.fsdecode(file)
            if fileEx.endswith(".xlsx"):
                xlsxName = str(fileEx)

        if path.exists(oldPath):
            newPath = Path(newPathString)
            pathlib.Path(newPath).mkdir(parents=True, exist_ok=True) 

            for file in os.listdir(oldPath):
                fileCurr = os.fsdecode(file)
                
                if fileCurr.endswith(".txt"):

                    #Capacitor
                    keywordsCapacitor = ['flute1_L_Capacitor', 'flute1_R_Capacitor']
                    for keyword in keywordsCapacitor:
                        if keyword in fileCurr:
                            fileOutCapacitor = capacitorFunc(oldPath, newPath, xlsxName, fileCurr)

                    #FET
                    keywordsFET = ['flute1_L_FET', 'flute1_R_FET']
                    for keyword in keywordsFET:
                        if keyword in fileCurr:
                            fileOutFET = FETFunc(oldPath, newPath, xlsxName, fileCurr)
                    #MOS
                    keywordsMOS = ['flute1_L_MOS', 'flute1_R_MOS']
                    for keyword in keywordsMOS:
                        if keyword in fileCurr:
                            fileOutMOS = MOSFunc(oldPath, newPath, xlsxName, fileCurr)

                    #strip
                    keywordsnPlus = ['flute1_L_n+', 'flute1_R_n+']
                    for keyword in keywordsnPlus:
                        if keyword in fileCurr:
                            fileOutstrip = stripFunc(oldPath, newPath, xlsxName, fileCurr)


                    #Poly
                    keywordsPoly = ['flute1_L_Poly', 'flute1_R_Poly']
                    for keyword in keywordsPoly:
                        if keyword in fileCurr:
                            fileOutPoly = PolyFunc(oldPath, newPath, xlsxName, fileCurr)

                    #pstop
                    keywordspstop = ['flute1_L_pstop', 'flute1_R_pstop']
                    for keyword in keywordspstop:
                        if keyword in fileCurr:
                            fileOutpstop = pstopFunc(oldPath, newPath, xlsxName, fileCurr)


                    #Dielectric breakdown
                    keywordsDielectric = ['flute2_L_Dielectric', 'flute2_R_Dielectric']
                    for keyword in keywordsDielectric:
                        if keyword in fileCurr:
                            fileOutDielectricBreakdown = DielectricBreakdownFunc(oldPath, newPath, xlsxName, fileCurr)

                    #GCD
                    keywordsGCD = ['flute2_L_GCD', 'flute2_R_GCD']
                    for keyword in keywordsGCD:
                        if keyword in fileCurr:
                            fileOutGCD = GCDFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Linewidth strip
                    keywordsnPlusLinewidth = ['flute2_L_n+_linewidth', 'flute2_R_n+_linewidth']
                    for keyword in keywordsnPlusLinewidth:
                        if keyword in fileCurr:
                            fileOutLinewidthStrip = LinewidthStripFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Linewidth Poly Meander
                    keywordsPolyMeander = ['flute2_L_PolyMeander', 'flute2_R_PolyMeander']
                    for keyword in keywordsPolyMeander:
                        if keyword in fileCurr:
                            fileOutLinewidthPolyMeander = LinewidthPolyMeanderFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Linewidth p-stop
                    keywordspstopLinewidth = ['flute2_L_pstopLinewidth', 'flute2_R_pstopLinewidth']
                    for keyword in keywordspstopLinewidth:
                        if keyword in fileCurr:
                            fileOutLinewidthpstop = LinewidthpstopFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Bulk cross
                    keywordsBulkCross = ['flute3_L_BulckCross', 'flute3_R_BulckCross']
                    for keyword in keywordsBulkCross:
                        if keyword in fileCurr:
                            fileOutBulkCross = BulkCrossFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Diode CV
                    keywordsDiodeCV = ['flute3_L_DiodeCV', 'flute3_R_DiodeCV']
                    for keyword in keywordsDiodeCV:
                        if keyword in fileCurr:
                            fileOutDiodeCV = DiodeCVFunc(oldPath, newPath, xlsxName, fileCurr)
                   
                    #Diode IV
                    keywordsDiodeIV = ['flute3_L_DiodeIV', 'flute3_R_DiodeIV']
                    for keyword in keywordsDiodeIV:
                        if keyword in fileCurr:
                            fileOutDiodeIV = DiodeIVFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Metal clover
                    keywordsMetalClover = ['flute3_L_MetalCover', 'flute3_R_MetalCover']
                    for keyword in keywordsMetalClover:
                        if keyword in fileCurr:
                            fileOutMetalClover = MetalCloverFunc(oldPath, newPath, xlsxName, fileCurr)

                    #p-Bridge
                    keywordspPlusBridge = ['flute3_L_p+Bridge', 'flute3_R_p+Bridge']
                    for keyword in keywordspPlusBridge:
                        if keyword in fileCurr:
                            fileOutpBridge = pBridgeFunc(oldPath, newPath, xlsxName, fileCurr)
                   
                    #p-Cross
                    keywordspPlusCross = ['flute3_L_p+Cross', 'flute3_R_p+Cross']
                    for keyword in keywordspPlusCross:
                        if keyword in fileCurr:
                            fileOutpCross = pCrossFunc(oldPath, newPath, xlsxName, fileCurr)

                    #Metal Meander Chain
                    keywordsMetalMeanderChain = ['L_flute3_Metal_Meander_Chain', 'R_flute3_Metal_Meander_Chain']
                    for keyword in keywordsMetalMeanderChain:
                        if keyword in fileCurr:
                            fileOutMetalMeanderChain = MetalMeanderChainFunc(oldPath, newPath, xlsxName, fileCurr)
                    
                    #GCD05
                    keywordsGCD05 = ['flute4_L_GCD', 'flute4_R_GCD']
                    for keyword in keywordsGCD05:
                        if keyword in fileCurr:
                            fileOutGCD05 = GCD05Func(oldPath, newPath, xlsxName, fileCurr)

                    #strip CBKR
                    keywordsnPlusCBKR = ['flute4_L_n+CBKR', 'flute4_R_n+CBKR']
                    for keyword in keywordsnPlusCBKR:
                        if keyword in fileCurr:
                            fileOutstripCBKR = stripCBKRFunc(oldPath, newPath, xlsxName, fileCurr)

                    #poly CBKR
                    keywordspolyCBKR = ['flute4_L_polyCBKR', 'flute4_R_polyCBKR']
                    for keyword in keywordspolyCBKR:
                        if keyword in fileCurr:
                            fileOutpolyCBKR = polyCBKRFunc(oldPath, newPath, xlsxName, fileCurr)

                    #n-chain
                    keywordsnPlusChain = ['L_flute4_n+_Chain', 'R_flute4_n+_Chain']
                    for keyword in keywordsnPlusChain:
                        if keyword in fileCurr:
                            fileOutnChain = nChainFunc(oldPath, newPath, xlsxName, fileCurr)

                    #p-chain
                    keywordspPlusChain = ['L_flute4_p+_Chain', 'R_flute4_p+_Chain']
                    for keyword in keywordspPlusChain:
                        if keyword in fileCurr:
                            fileOutpChain = pChainFunc(oldPath, newPath, xlsxName, fileCurr)

                    #poly Chain
                    keywordsPolyChain = ['L_flute4_Poly_Chain', 'R_flute4_Poly_Chain']
                    for keyword in keywordsPolyChain:
                        if keyword in fileCurr:
                            fileOutpolyChain = polyChainFunc(oldPath, newPath, xlsxName, fileCurr)


        zipFileName = newPathString + '.zip'

        zf = zipfile.ZipFile(zipFileName, "w")
        for dirname, subdirs, files in os.walk(newPath):
            zf.write(dirname)
            for filename in files:
                zf.write(os.path.join(dirname, filename))
        zf.close()
