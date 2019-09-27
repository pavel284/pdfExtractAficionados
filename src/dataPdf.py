# -*- coding: utf-8 -*-
"""
"""

#Ing. Paul Avendaño Elizondo


import PyPDF2
import re
import PySimpleGUI as sg
#import pandas as pd
from openpyxl import Workbook, load_workbook 
import os

def listaAcuerdos (folder1):
    acuerdosList = []
    for path in os.listdir(folder1):
        acuerdoFiles = os.path.join(folder1, path)
        if os.path.isfile(acuerdoFiles):
            acuerdosList.append(acuerdoFiles)
           
    return acuerdosList

#function to save files
def saveFile(path,str):
    savePdf = open(path,'w')
    savePdf.write(str)
    savePdf.close()

#function to use regex
def patterns (searching,doc):
    result = re.findall(searching,doc)
    return result

#function to extract the first element of every match
def unique (lista):
    element = lista[0]
    return element

def PDF (source):
    # pdf file object
    pdf = open(source, 'rb')
    # pdf reader object
    pdfDoc = PyPDF2.PdfFileReader(pdf)
    # number of pages in pdf
    pagesPdf = pdfDoc.numPages
    #declare variables for while
    page = 0
    textPdf =""
    #read every page and extact text for every page
    while page < pagesPdf:
        # a page object
        pageObj = pdfDoc.getPage(page)
        #extract text from a page PDF
        text = pageObj.extractText()
        #concat the strings
        textPdf += text
        page +=1
    return textPdf    

def isCategorie (categoria):
    clase = ""
    if categoria =="Clase C":
        clase = "NOVICIO"
    elif categoria == "Clase B":
        clase = "INTERMEDIO"
    elif categoria == "Clase A":
        clase = "SUPERIOR"
    return clase    

def isCompany (patternJurid,textPdf,cedula):
    if patterns(patternJurid,textPdf):
        idAsocia = patterns(patternJurid,textPdf)
        cedJuridica = unique(idAsocia).replace("-","")
        return cedJuridica        
    else:
        data = cedula
        return data

def isAdult(patternName, patternMenor, textPdf):
    if patterns(patternName,textPdf):
        names = patterns (patternName,textPdf)
        name = unique(names)
        nameObj1 = name.replace("Permisionario\n","")
        nombre = nameObj1.replace("\nC","")
        return nombre
    elif patterns(patternMenor,textPdf):
        names = patterns (patternMenor,textPdf)
        name = unique(names)
        nombre = name.replace("menor","")
        return nombre

def saveExc (finalData, pathSave): 
    wb = Workbook()
    ws = wb.active
    for row in finalData:
        ws.append(row)
        wb.save(pathSave)

def fillData (dataTest1, dataTest2, patternInd, patternIcb, textPdf):
    listaData = []
    if patterns(patternInd,textPdf) and patterns(patternIcb,textPdf):
        categorie = patterns(patternClass,textPdf)
        callSign = patterns (patternInd,textPdf)
        
        #unique data
        categoria = unique(categorie)
        
        #define categorie
        clase = isCategorie (categoria)
        indicativo = callSign[1]
        callSignCb = patterns (patternIcb,textPdf)
        
        #unique data
        indiCb = unique(callSignCb)
        dataTest1[6] = clase
        dataTest1[22] = indicativo
        dataTest2[6] = "BANDA CIUDADANA"
        dataTest2[22] = indiCb
        
        listaData.append(dataTest1)
        listaData.append(dataTest2)
        print(dataTest1)
        print(dataTest2)
        return listaData
    
    elif re.findall(patternInd,textPdf):
        categorie = patterns(patternClass,textPdf)
        callSign = patterns (patternInd,textPdf)
        #unique data
        categoria = unique(categorie)
        #define categorie
        clase = isCategorie(categoria)
        indicativo = callSign[1]
        dataTest1[6] = clase
        dataTest1[22] = indicativo
        
        listaData.append(dataTest1)
        print(dataTest1)
        return listaData
    
    elif re.findall(patternIcb,textPdf):
        callSignCb = patterns (patternIcb,textPdf)
        #unique data
        indiCb = unique(callSignCb)
        
        dataTest2[6] = "BANDA CIUDADANA"
        dataTest2[22] = indiCb
        
        listaData.append(dataTest2)
        print(dataTest2)
      
        return listaData


#regex to import all the coincedences
patternName = "[a-zA-Z]ermisionari[a-z]+\s[a-zA-Zá-úÁ-ÚñÑ\s]+C"
patternMenor = "menor\s[a-zA-z-á-úÁ-ÚñÑ\s]+"
patternNum = "[0-9]+\-[0-9]+\-TEL\-MICITT"
patternId = "[a-zA-Z]édula\s[a-zá-ú\]+\sNº\s[0-9-\s]+"
patternCed = "[0-9]\-[0-9][0-9][0-9][0-9]\-[0-9]+"
patternJurid = "3-002-[0-9]+"
patternClass = "[a-zA-Z]lase\s[A-C]"
patternCb = "[a-zA-Z]anda\s[a-zA-Z]iudadana"
patternInd = "TI[0-9][A-Z]+"
patternIcb = "TEA[0-9][A-Z]+"
patternPer = "[a-zA-Z]ermisionari[a-z]+\s"

folder = sg.PopupGetFolder("Seleccione la carpeta de sus archivos PDF")
if folder is None:
    raise SystemExit

listaData = []

acuerdosList = listaAcuerdos (folder)
cont = 0
while cont < len(acuerdosList):
    path = acuerdosList[cont]
    textPdf = PDF (acuerdosList[cont])
    #save my coincedences lists
    numAcuerdo = patterns (patternNum,textPdf)
    names = patterns (patternName,textPdf)
    ced = patterns (patternCed,textPdf)

    categoria = ""
    indicativo = ""
    bandaCiudadana = ""
    indiCb = ""

    #unique data
    nombre = isAdult (patternName, patternMenor,textPdf)
    cedula = unique(ced).replace("-","")
    acuerdo = unique(numAcuerdo)

    #data I want in my Excel file
    data = [cedula,nombre,acuerdo,categoria,indicativo,bandaCiudadana,indiCb]
    dataTest1 = ["",nombre," ",acuerdo," "," ","Cat"," "," "," "," "," "," "," "," "," "," "," "," "," "," "," ","","6","80"]
    dataTest2= ["",nombre," ",acuerdo," "," ","Cat"," "," "," "," "," "," "," "," "," "," "," "," "," "," "," ","","6","81"]

    # fill data with person ID or Company ID
    dataTest1[0] = isCompany (patternJurid,textPdf,cedula)
    dataTest2[0] = isCompany(patternJurid,textPdf,cedula)
    
    # fill the lines according to the categorie
    listaData.append(fillData(dataTest1,dataTest2,patternInd,patternIcb,textPdf))
    
    cont += 1

finalData=[]
for n in listaData:
    for x in n:
        finalData.append(x)

sg.Popup('Seleccione una carpeta y escriba nombre del archivo a guardar')
pathSave = sg.PopupGetFile('Guardar en formato Excel',default_path='.xlsx', default_extension='.xlsx', save_as=True, file_types=(("Excel", '.xlsx'),), no_window=False, font=None, no_titlebar=False, grab_anywhere=False)


saveExc (finalData, pathSave)

################################################# 
# -*- coding: utf-8 -*-

