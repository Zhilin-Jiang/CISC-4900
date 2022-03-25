import PyPDF4 as pypdf

from datetime import datetime
import time
import os
import glob
import csv
import argparse
from PyPDF4.generic import PdfObject


# ----------------------------------------
# Function definitions
# ----------------------------------------
# allow user supplied path, if not supplied
# default path is current working directory
def main(path, out_path, filename):

    # get all the files in current working directory
    files_in_directory = glob.glob(os.path.join(path, '*.pdf'))

    # start extracting from files. using list comprehension for readability
    output = [extract_students(file) for file in files_in_directory]
    
    ComScien = []
    InfSystem = []
    MulComp = []

    # save extracted data to file with today as timestamp
    save_as = "_report_{}.csv".format(datetime.today().strftime('%Y-%m-%d'))
    save_to_record(output, save_as)

    [extract_major(file,ComScien,InfSystem,MulComp) for file in files_in_directory]

    save_to_CS(ComScien,'Computer Science.csv')
    save_to_IS(InfSystem,'Information Systems.csv')
    save_to_MC(MulComp,'Multimedia Computing.csv')


# Extract Student's Information


def extract_students(file):
    # retrieve the timestamp and filename
    document_info = {"Timestamp": time.ctime(
        os.path.getmtime(file)), "Filename": file.split("/")[-1]}

    pdfobject = open(file, 'rb')
    pdf = pypdf.PdfFileReader(pdfobject)
    #Extract field data from pdf which contains interactive form
    data = pdf.getFields()
    dataDict = []
    if isinstance(data, dict):
        dataDict = {k: v.get('/V', '') for (k, v) in data.items()}
    result = {}
    result.update({"First Name": dataDict.get("studentfname")})
    result.update({"Last Name": dataDict.get("studentlname")})
    result.update({"Emplid": dataDict.get("studentemplid")})
    result.update({"Expected_Graduation": dataDict.get("studentsemester")})
    result.update({"Student Major": dataDict.get("studentmajor")})
    result.update(document_info)
    return result

#extract course list for each major
def extract_major(file,ComScien,InfSystem,MulComp):
    pdfFile = open(file,'rb')
    Readpdf = pypdf.PdfFileReader(pdfFile)
    Pdfdata = Readpdf.getFields()
    dataDict1 = []
    if isinstance(Pdfdata, dict):
        dataDict1 = {k: v.get('/V', '') for (k, v) in Pdfdata.items()}

    if dataDict1.get('studentmajor') == 'Computer Science':
            ComScien.append(extract_CS(file))

    elif dataDict1.get('studentmajor') == 'Information Systems':
            InfSystem.append(extract_IS(file))

    else:
            MulComp.append(extract_MC(file))


def extract_CS(file):
    FormFile = open(file,'rb')
    pdfRead = pypdf.PdfFileReader(FormFile)
    Formdata = pdfRead.getFields()
    if isinstance(Formdata, dict):
        dataDict2 = {k: v.get('/V', '') for (k, v) in Formdata.items()}
    ComSci = {}
    ComSci.update({'First Name': dataDict2.get('studentfname')})
    ComSci.update({'Last Name': dataDict2.get('studentlname')})
    ComSci.update({'CUNY ID': dataDict2.get('studentemplid')})
    ComSci.update({'Email': dataDict2.get('studentemail')})
    ComSci.update({'Final Semester':dataDict2.get('studentsemester')})
    # CISC 1115 OR CISC 1170
    if dataDict2.get('1115instructor') == '':
        ComSci.update(
            {'CISC 1115/CISC 1170-Instructor': dataDict2.get('1170instructor')})
        ComSci.update(
            {'CISC 1115/CISC1170-Grade': dataDict2.get('1170grade')})
    else:
        ComSci.update(
            {'CISC 1115/CISC 1170-Instructor': dataDict2.get('1115instructor')})
        ComSci.update(
            {'CISC 1115/CISC1170-Grade': dataDict2.get('1115grade')})
    ComSci.update({'CISC 2210-Instructor':dataDict2.get('2210instructor')})
    ComSci.update({'CISC 2210-Grade': dataDict2.get('2210grade')})
    ComSci.update({'CISC 3115-Instructor':dataDict2.get('3115instructor')})
    ComSci.update({'CISC 3115-Grade': dataDict2.get('3115grade')})
    ComSci.update({'CISC 3130-Instructor':dataDict2.get('3130instructor')})
    ComSci.update({'CISC 3130-Grade': dataDict2.get('3130grade')})
    ComSci.update({'CISC 3140-Instructor':dataDict2.get('3140instructor')})
    ComSci.update({'CISC 3140-Grade': dataDict2.get('3140grade')})
    ComSci.update({'CISC 3142-Instructor':dataDict2.get('3142instructor')})
    ComSci.update({'CISC 3142-Grade': dataDict2.get('3142grade')})
    ComSci.update({'CISC 3220-Instructor':dataDict2.get('3220instructor')})
    ComSci.update({'CISC 3220-Grade': dataDict2.get('3220grade')})
    # CISC 3305 OR CISC 3310
    if dataDict2.get('3305instructor') == '':
        ComSci.update(
            {'CISC 3305/CISC 3310-Instructor':dataDict2.get('3310instructor')})
        ComSci.update(
            {'CISC 3305/CISC3310-Grade': dataDict2.get('3310grade')})
    else:
        ComSci.update(
            {'CISC 3305/CISC 3310-Instructor': dataDict2.get('3305instructor')})
        ComSci.update(
            {'CISC 3305/CISC3310-Grade': dataDict2.get('3305grade')})
    ComSci.update({'CISC 3320-Instructor':dataDict2.get('3320instructor')})
    ComSci.update({'CISC 3320-Grade': dataDict2.get('3320grade')})
    # CISC 4900 OR CISC 5001
    if dataDict2.get('4900instructor') == '':
        ComSci.update(
            {'CISC 4900/CISC 5001-Instructor':dataDict2.get('5001instructor')})
        ComSci.update(
            {'CISC 4900/CISC 5001-Grade': dataDict2.get('5001grade')})
    else:
        ComSci.update(
            {'CISC 4900/CISC 5001-Instructor': dataDict2.get('4900instructor')})
        ComSci.update(
            {'CISC 4900/CISC 5001-Grade': dataDict2.get('4900grade')})
    #CISC 2820W OR Phil 3318W
    if dataDict2.get('2820Winstructor') == '':
        ComSci.update(
                {'CISC 2820W/Phil 3318W-Instructor':dataDict2.get('P3318Winstructor')})
        ComSci.update(
                {'CISC 2820W/Phil 3318W-Grade': dataDict2.get('P3318Wgrade')})
    else:
        ComSci.update(
                {'CISC 2820W/Phil 3318W-Instructor': dataDict2.get('2820Winstructor')})
        ComSci.update(
                {'CISC 2820W/Phil 3318W-Grade': dataDict2.get('2820Wgrade')})
    ComSci.update({'Math 1201-Instructor':dataDict2.get('M1201instructor')})
    ComSci.update({'Math 1201-Grade': dataDict2.get('M1201grade')})
    ComSci.update({'Math 1206-Instructor': dataDict2.get('M1206instructor')})
    ComSci.update({'Math 1206-Grade': dataDict2.get('M1206grade')})
    ComSci.update({'Math 1211-Instructor': dataDict2.get('M1211instructor')})
    ComSci.update({'Math 1211-Grade': dataDict2.get('M1211grade')})
    # MATH 2501 0R MATH 3501
    if dataDict2.get('M2501instructor') == '':
        ComSci.update(
            {'Math 2501/Math 3501-Instructor': dataDict2.get('M3501instructor')})
        ComSci.update(
            {'Marh 2501/Math 3501-Grade': dataDict2.get('M3501grade')})
    else:
        ComSci.update(
            {'Math 2501/Math 3501-Instructor': dataDict2.get('M2501instructor')})
        ComSci.update(
            {'Marh 2501/Math 3501-Grade': dataDict2.get('M2501grade')})
    ComSci.update({'Additional CIS Course1-Instructor':dataDict2.get('a1instructor')})
    ComSci.update({'Additional CIS Course1-Grade': dataDict2.get('a1grade')})
    ComSci.update({'Additional CIS Course2-Instructor': dataDict2.get('a2instructor')})
    ComSci.update({'Additional CIS Course2-Grade': dataDict2.get('a2grade')})
    ComSci.update({'Additional CIS Course3-Instuctor': dataDict2.get('a3instructor')})
    ComSci.update({'Additional CIS Course3-Grade': dataDict2.get('a3grade')})
    ComSci.update({'Other1-Instructor':dataDict2.get('other1instructor')})
    ComSci.update({'Other1-Grade': dataDict2.get('other1grade')})
    ComSci.update({'Other2-Instructor': dataDict2.get('other2instructor')})
    ComSci.update({'Other2-Grade': dataDict2.get('other2grade')})
    ComSci.update({'Other3-Instructor': dataDict2.get('other3instructor')})
    ComSci.update({'Other3-Grade': dataDict2.get('other3grade')})
    ComSci.update({'Other4-Instructor': dataDict2.get('other4instructor')})
    ComSci.update({'Other4-Grade': dataDict2.get('other4grade')})
    ComSci.update({'Other5-Instructor': dataDict2.get('other5instructor')})
    ComSci.update({'Other5-Grade': dataDict2.get('other5grade')})
    return ComSci

def extract_IS(file):
    FormFile1 = open(file, 'rb')
    pdfRead1 = pypdf.PdfFileReader(FormFile1)
    Formdata1 = pdfRead1.getFields()
    if isinstance(Formdata1, dict):
        dataDict3 = {k: v.get('/V', '') for (k, v) in Formdata1.items()}
    InfSys = {}
    InfSys.update({'First Name': dataDict3.get('studentfname')})
    InfSys.update({'Last Name': dataDict3.get('studentlname')})
    InfSys.update({'CUNY ID': dataDict3.get('studentemplid')})
    InfSys.update({'Email': dataDict3.get('studentemail')})
    InfSys.update({'Final Semester': dataDict3.get('studentsemester')})
    # CISC 1115 OR CISC 1170
    if dataDict3.get('1115instructor') == '':
        InfSys.update(
            {'CISC 1115/CISC 1170-Instructor': dataDict3.get('1170instructor')})
        InfSys.update(
            {'CISC 1115/CISC 1170-Grade': dataDict3.get('1170grade')})
    else:
        InfSys.update(
            {'CISC 1115/CISC 1170-Instructor': dataDict3.get('1115instructor')})
        InfSys.update(
            {'CISC 1115/CISC 1170-Grade': dataDict3.get('1115grade')})
    InfSys.update({'CISC 3115-Instructor': dataDict3.get('3115instructor')})
    InfSys.update({'CISC 3115-Grade': dataDict3.get('3115grade')})
    InfSys.update({'CISC 3130-Instructor': dataDict3.get('3130instructor')})
    InfSys.update({'CISC 3130-Grade': dataDict3.get('3130grade')})
    InfSys.update({'CISC 3810-Instructor': dataDict3.get('3180instructor')})
    InfSys.update({'CISC 3810-Grade': dataDict3.get('3180grade')})
    # CISC 4900 OR CISC 5001
    if dataDict3.get('4900instructor') == '':
        InfSys.update(
            {'CISC 4900/CISC 5001-Instructor': dataDict3.get('5001instructor')})
        InfSys.update(
            {'CISC 4900/CISC 5001-Grade': dataDict3.get('5001grade')})
    else:
        InfSys.update(
            {'CISC 4900/CISC 5001-Instructor': dataDict3.get('4900instructor')})
        InfSys.update(
            {'CISC 4900/CISC 5001-Grade': dataDict3.get('4900grade')})
    InfSys.update({'Other CIS Course1-Instructor': dataDict3.get('a1instructor')})
    InfSys.update({'Other CIS Course1-Grade': dataDict3.get('a1grade')})
    InfSys.update({'Other CIS Course2-Instructor': dataDict3.get('a2instructor')})
    InfSys.update({'Other CIS Course2-Grade': dataDict3.get('a2grade')})
    InfSys.update({'Other CIS Course3-Instuctor': dataDict3.get('a3instructor')})
    InfSys.update({'Other CIS Course3-Grade': dataDict3.get('a3grade')})
    #CISC 2820W OR Phil 3318W
    if dataDict3.get('2820Winstructor') == '':
        InfSys.update(
            {'CISC 2820W/Phil 3318W-Instructor': dataDict3.get('3318Winstructor')})
        InfSys.update(
            {'CISC 2820W/Phil 3318W-Grade': dataDict3.get('3318grade')})
    else:
        InfSys.update(
            {'CISC 2820W/Phil 3318W-Instructor': dataDict3.get('2820Winstructor')})
        InfSys.update(
            {'CISC 2820W/Phil 3318W-Grade': dataDict3.get('2820Wgrade')})
    # Business 3420 OR CISC 1500
    if dataDict3.get('3420instructor') == '':
        InfSys.update(
            {'Business 3420/CISC 1590-Instructor':dataDict3.get('1590instructor')})
        InfSys.update(
            {'Business 3420/CISC 1590-Grade': dataDict3.get('1590grade')})
    else:
        InfSys.update(
            {'Business 3420/CISC 1590-Instructor':dataDict3.get('3420instructor')}
        )
        InfSys.update(
            {'Business 3420/CISC 1590-Grade':dataDict3.get('3420grade')}
        )
    # Business 3430 OR CISC 2531
    if dataDict3.get('3430instructor') == '':
        InfSys.update(
            {'Business 3430/CISC 2531-Instructor':dataDict3.get('2531instructor')})
        InfSys.update(
            {'Business 3430/CISC 2531-Grade':dataDict3.get('2531grade')}
        )
    else:
        InfSys.update(
            {'Business 3430/CISC 2531-Instructor':dataDict3.get('3430instructor')}
        )
        InfSys.update(
            {'Business 3430/CISC 2531-Grade':dataDict3.get('3430grade')}
        )
    # Business 3120 OR CISC 1530 OR Business 3432 OR CISC 2532
    if dataDict3.get('3120instructor') != '':
        InfSys.update(
            {'Business 3120/CISC 1530/Business 3432/CISC 2532-Instructor':dataDict3.get('3120instructor')}
        )
        InfSys.update(
            {'Business 3120/CISC 1530/Business 3432/CISC 2532-Grade':
                dataDict3.get('3120grade')}
        )
    elif dataDict3.get('1530instructor') != '':
        InfSys.update(
            {'Business 3120/CISC 1530/Business 3432/CISC 2532-Instructor':
                dataDict3.get('1530instructor')})
        InfSys.update(
            {'Business 3120/CISC 1530/Business 3432/CISC 2532-Grade':
                dataDict3.get('1530grade')}
        )
    elif dataDict3.get('3432instructor') != '':
        InfSys.update(
            {'Business 3120/CISC 1530/Business 3432/CISC 2532-Instructor':
                dataDict3.get('3432instructor')}
        )
        InfSys.update(
            {'Business 3120/CISC 1530/Business 3432/CISC 2532-Grade':
                dataDict3.get('3432grade')}
        )
    else:
        InfSys.update(
            {'Business 3120/CISC 1530/Business 3432/CISC 2532-Instructor':
                dataDict3.get('2532instructor')}
        )
        InfSys.update(
            {'Business 3120/CISC 1530/Business 3432/CISC 2532-Grade':
                dataDict3.get('2532grade')}
        )
    # Business 4202W OR CISC 1580W
    if dataDict3.get('4202Winstructor') == '':
        InfSys.update(
            {'Business 4202W/CISC 1580W-Instructor':dataDict3.get('1580Winstructor')}
        )
        InfSys.update(
            {'Business 4202W/CISC 1580W-Grade':dataDict3.get('1580Wgrade')}
        )
    else:
        InfSys.update(
            {'Business 4202W/CISC 1580W-Instructor':dataDict3.get('4202Winstructor')}
        )
        InfSys.update(
            {'Business 4202W/CISC 1580W-Grade':dataDict3.get('4202Wgrade')}
        )
    # Economics 2100 OR Business 2100
    if dataDict3.get('2100ainstructor') == '':
        InfSys.update(
            {'Economics 2100/Business 2100-Instructor':dataDict3.get('2100binstructor')}
        )
        InfSys.update(
            {'Economics 2100/Business 2100-Grade':dataDict3.get('2100bgrade')}
        )
    else:
        InfSys.update(
            {'Economics 2100/Business 2100-Instructor':dataDict3.get('2100ainstructor')}
        )
        InfSys.update(
            {'Economics 2100/Business 2100-Grade':dataDict3.get('2100agrade')}
        )
    # Economics 2200 OR Business 2200
    if dataDict3.get('2200ainstructor') == '':
        InfSys.update(
            {'Economics 2200/Business 2200-Instructor':
                dataDict3.get('2200binstructor')}
        )
        InfSys.update(
            {'Economics 2200/Business 2200-Grade': dataDict3.get('2200bgrade')}
        )
    else:
        InfSys.update(
            {'Economics 2200/Business 2200-Instructor':
                dataDict3.get('2200ainstructor')}
        )
        InfSys.update(
            {'Economics 2200/Business 2200-Grade': dataDict3.get('2200agrade')}
        )
    InfSys.update({'Business 3200-Instructor':dataDict3.get('3200instructor')})
    InfSys.update({'Business 3200-Grade': dataDict3.get('3200grade')})
    InfSys.update({'Finance 3310-Instructor':dataDict3.get('3310instructor')})
    InfSys.update({'Finance 3310-Grade': dataDict3.get('3310grade')})
    InfSys.update({'Accounting 2001-Instructor':dataDict3.get('2001instructor')})
    InfSys.update({'Accounting 2001-Grade': dataDict3.get('2001grade')})
    # Business 3400 OR Econ 3400 OR Math 2501 OR Math 3501 OR Psy 3400 OR Businsess 3410 OR Math 1201 OR Business 3421 OR CISC 2590
    if dataDict3.get('3400ainstructor') != '':
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor':dataDict3.get('3400ainstructor')})
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade': dataDict3.get('3400agrade')})
    elif dataDict3.get('3400binstructor') != '':
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor': dataDict3.get('3400binstructor')})
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade': dataDict3.get('3400bgrade')})
    elif dataDict3.get('2501instructor') != '':
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor': dataDict3.get('2501instructor')})
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade': dataDict3.get('2501grade')})
    elif dataDict3.get('3501instructor') != '':
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor': dataDict3.get('3501instructor')})
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade': dataDict3.get('3501grade')})
    elif dataDict3.get('3400cinstructor') == '':
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor': dataDict3.get('3400cinstructor')})
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade': dataDict3.get('3400cgrade')})
    elif dataDict3('3410instructor') == '':
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor': dataDict3.get('3410instructor')})
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade': dataDict3.get('3410grade')})
    elif dataDict3('1201instructor') == '':
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor': dataDict3.get('1201instructor')})
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade': dataDict3.get('1201grade')})
    elif dataDict3('3421instructor') == '':
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor': dataDict3.get('3421ainstructor')})
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade': dataDict3.get('3421agrade')})
    elif dataDict3('2590instructor') == '':
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor': dataDict3.get('2590instructor')})
        InfSys.update(
            {'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade': dataDict3.get('2590agrade')})
    InfSys.update({'Other1-Instructor': dataDict3.get('other1instructor')})
    InfSys.update({'Other1-Grade': dataDict3.get('other1grade')})
    InfSys.update({'Other2-Instructor': dataDict3.get('other2instructor')})
    InfSys.update({'Other2-Grade': dataDict3.get('other2grade')})
    InfSys.update({'Other3-Instructor': dataDict3.get('other3instructor')})
    InfSys.update({'Other3-Grade': dataDict3.get('other3grade')})
    InfSys.update({'Other4-Instructor': dataDict3.get('other4instructor')})
    InfSys.update({'Other4-Grade': dataDict3.get('other4grade')})
    InfSys.update({'Other5-Instructor': dataDict3.get('other5instructor')})
    InfSys.update({'Other5-Grade': dataDict3.get('other5grade')})
    return InfSys
    
 




def extract_MC(file):
    FormFile2 = open(file, 'rb')
    pdfRead2 = pypdf.PdfFileReader(FormFile2)
    Formdata2 = pdfRead2.getFields()
    if isinstance(Formdata2, dict):
        dataDict3 = {k: v.get('/V', '') for (k, v) in Formdata2.items()}
    MulComp = {}
    MulComp.update({'First Name': dataDict3.get('studentfname')})
    MulComp.update({'Last Name': dataDict3.get('studentlname')})
    MulComp.update({'CUNY ID': dataDict3.get('studentemplid')})
    MulComp.update({'Email': dataDict3.get('studentemail')})
    MulComp.update({'Final Semester': dataDict3.get('studentsemester')})
    # CISC 1115 OR CISC 1170
    if dataDict3.get('1115instructor') == '':
        MulComp.update(
            {'CISC 1115/CISC 1170-Instructor': dataDict3.get('1170instructor')})
        MulComp.update(
            {'CISC 1115/CISC 1170-Grade': dataDict3.get('1170grade')})
    else:
        MulComp.update(
            {'CISC 1115/CISC 1170-Instructor': dataDict3.get('1115instructor')})
        MulComp.update(
            {'CISC 1115/CISC 1170-Grade': dataDict3.get('1115grade')})
    MulComp.update({'CISC 1600-Instructor':dataDict3.get('1600instructor')})
    MulComp.update({'CISC 1600-Grade': dataDict3.get('1600grade')})
    MulComp.update({'CISC 2210-Instructor':dataDict3.get('2210instructor')})
    MulComp.update({'CISC 2210-Grade': dataDict3.get('2210grade')})
    MulComp.update({'CISC 2820W-Instructor':dataDict3.get('2820Winstructor')})
    MulComp.update({'CISC 2820W-Grade': dataDict3.get('2820Wgrade')})
    MulComp.update({'CISC 3115-Instructor':dataDict3.get('3115instructor')})
    MulComp.update({'CISC 3115-Grade': dataDict3.get('3115grade')})
    MulComp.update({'CISC 3130-Instructor': dataDict3.get('3130instructor')})
    MulComp.update({'CISC 3130-Grade': dataDict3.get('3130grade')})
    MulComp.update({'CISC 3220-Instructor': dataDict3.get('3220instructor')})
    MulComp.update({'CISC 3220-Grade': dataDict3.get('3220grade')})
    MulComp.update({'CISC 3620-Instructor': dataDict3.get('3620instructor')})
    MulComp.update({'CISC 3620-Grade': dataDict3.get('3620grade')})
    MulComp.update({'CISC 3630-Instructor': dataDict3.get('3630instructor')})
    MulComp.update({'CISC 3630-Grade': dataDict3.get('3630grade')})
    # CISC 4900 OR CISC 5001
    if dataDict3.get('4900instructor') == '':
        MulComp.update(
            {'CISC 4900/CISC 5001-Instructor': dataDict3.get('5001instructor')})
        MulComp.update(
            {'CISC 4900/CISC 5001-Grade': dataDict3.get('5001grade')})
    else:
        MulComp.update(
            {'CISC 4900/CISC 5001-Instructor': dataDict3.get('4900instructor')})
        MulComp.update(
            {'CISC 4900/CISC 5001-Grade': dataDict3.get('4900grade')})
    MulComp.update({'Math 1006-Instructor':dataDict3.get('M1006instructor')})
    MulComp.update({'Math 1006-Grade': dataDict3.get('M1006grade')})
    # Math 1011 OR Math 1012
    if dataDict3.get('M1011instructor') == '':
        MulComp.update(
            {'Math 1011/Math 1012-Instructor':dataDict3.get('M1012instructor')})
        MulComp.update(
            {'Math 1011/Math 1012-Grade': dataDict3.get('M1012grade')})
    else:
        MulComp.update(
            {'Math 1011/Math 1012-Instructor': dataDict3.get('M1011instructor')})
        MulComp.update(
            {'Math 1011/Math 1012-Grade': dataDict3.get('M1011grade')})
    MulComp.update({'Math 1201-Instructor':dataDict3.get('M1201instructor')})
    MulComp.update({'Math 1201-Grade': dataDict3.get('M1201grade')})
    # Math 1711 OR Math 1206
    if dataDict3.get('M1711instructor') == '':
        MulComp.update(
            {'Math 1711/Math 1206-Instructor': dataDict3.get('M1206instructor')})
        MulComp.update(
            {'Math 1711/Math 1206-Grade': dataDict3.get('M1206grade')})
    else:
        MulComp.update(
            {'Math 1711/Math 1206-Instructor': dataDict3.get('M1711instructor')})
        MulComp.update(
            {'Math 1711/Math 1206-Grade': dataDict3.get('M1711grade')})
    # Math 1716 OR Math 2501
    if dataDict3.get('M1716instructor') == '':
        MulComp.update(
            {'Math 1716/Math 2501-Instructor': dataDict3.get('M2501instructor')})
        MulComp.update(
            {'Math 1716/Math 2501-Grade': dataDict3.get('M2501grade')})
    else:
        MulComp.update(
            {'Math 1716/Math 2501-Instructor': dataDict3.get('M1716instructor')})
        MulComp.update(
            {'Math 1716/Math 2501-Grade': dataDict3.get('M1716grade')})
    MulComp.update({'Additional CIS Course1-Instructor': dataDict3.get('a1instructor')})
    MulComp.update({'Additional CIS Course1-Grade': dataDict3.get('a1grade')})
    MulComp.update({'Additional CIS Course2-Instructor': dataDict3.get('a2instructor')})
    MulComp.update({'Additional CIS Course2-Grade': dataDict3.get('a2grade')})
    MulComp.update({'Additional CIS Course3-Instuctor': dataDict3.get('a3instructor')})
    MulComp.update({'Additional CIS Course3-Grade': dataDict3.get('a3grade')})
    MulComp.update({'Elective Class1-Instructor':dataDict3.get('b1instructor')})
    MulComp.update({'Elective Class1-Grade': dataDict3.get('b1grade')})
    MulComp.update({'Elective Class2-Instructor': dataDict3.get('b2instructor')})
    MulComp.update({'Elective Class2-Grade': dataDict3.get('b2grade')})
    MulComp.update({'Other1-Instructor': dataDict3.get('other1instructor')})
    MulComp.update({'Other1-Grade': dataDict3.get('other1grade')})
    MulComp.update({'Other2-Instructor': dataDict3.get('other2instructor')})
    MulComp.update({'Other2-Grade': dataDict3.get('other2grade')})
    MulComp.update({'Other3-Instructor': dataDict3.get('other3instructor')})
    MulComp.update({'Other3-Grade': dataDict3.get('other3grade')})
    MulComp.update({'Other4-Instructor': dataDict3.get('other4instructor')})
    MulComp.update({'Other4-Grade': dataDict3.get('other4grade')})
    MulComp.update({'Other5-Instructor': dataDict3.get('other5instructor')})
    MulComp.update({'Other5-Grade': dataDict3.get('other5grade')})




    return MulComp

# saves extracted data to file
def save_to_record(data, output_filename):
    with open(output_filename, 'w', newline='') as output_file:
        head = ['Timestamp', 'Filename',  'Emplid',
                'First Name', 'Last Name', 'Expected_Graduation', 'Student Major']
        dict_writer = csv.DictWriter(output_file, head)
        dict_writer.writeheader()
        for data1 in data:
            dict_writer.writerow(data1)

    print("Saved results to {}".format(output_filename))

# save cs student's course information
def save_to_CS(data, output_filename):
    with open(output_filename, 'w', newline='') as output_file:

        head = ['First Name','Last Name','CUNY ID','Email','Final Semester','CISC 1115/CISC 1170-Instructor',
                'CISC 1115/CISC1170-Grade','CISC 2210-Instructor','CISC 2210-Grade','CISC 3115-Instructor','CISC 3115-Grade',
                'CISC 3130-Instructor','CISC 3130-Grade','CISC 3140-Instructor','CISC 3140-Grade','CISC 3142-Instructor','CISC 3142-Grade',
                'CISC 3220-Instructor','CISC 3220-Grade','CISC 3305/CISC 3310-Instructor','CISC 3305/CISC3310-Grade','CISC 3320-Instructor','CISC 3320-Grade',
                'CISC 4900/CISC 5001-Instructor','CISC 4900/CISC 5001-Grade','CISC 2820W/Phil 3318W-Instructor','CISC 2820W/Phil 3318W-Grade',
                'Math 1201-Instructor','Math 1201-Grade','Math 1206-Instructor','Math 1206-Grade','Math 1211-Instructor','Math 1211-Grade','Math 2501/Math 3501-Instructor','Marh 2501/Math 3501-Grade',
                'Additional CIS Course1-Instructor', 'Additional CIS Course1-Grade', 'Additional CIS Course2-Instructor','Additional CIS Course2-Grade',
                'Additional CIS Course3-Instuctor','Additional CIS Course3-Grade','Other1-Instructor','Other1-Grade','Other2-Instructor','Other2-Grade','Other3-Instructor','Other3-Grade',
                'Other4-Instructor','Other4-Grade','Other5-Instructor','Other5-Grade' ]
        dict_writer = csv.DictWriter(output_file, head)
        dict_writer.writeheader()
        for data1 in data:
            dict_writer.writerow(data1)
    
            
        
            

    print("Saved results to {}".format(output_filename))

# save multimedia student's course information
def save_to_MC(data, output_filename):
    with open(output_filename, 'w', newline='') as output_file:

        head = ['First Name', 'Last Name', 'CUNY ID', 'Email', 'Final Semester', 'CISC 1115/CISC 1170-Instructor','CISC 1115/CISC 1170-Grade','CISC 1600-Instructor','CISC 1600-Grade',
                'CISC 2210-Instructor', 'CISC 2210-Grade', 'CISC 2820W-Instructor', 'CISC 2820W-Grade', 'CISC 3115-Instructor', 'CISC 3115-Grade', 'CISC 3130-Instructor', 'CISC 3130-Grade',
                'CISC 3220-Instructor', 'CISC 3220-Grade', 'CISC 3620-Instructor', 'CISC 3620-Grade', 'CISC 3630-Instructor', 'CISC 3630-Grade', 'CISC 4900/CISC 5001-Instructor', 'CISC 4900/CISC 5001-Grade',
                'Math 1006-Instructor', 'Math 1006-Grade', 'Math 1011/Math 1012-Instructor', 'Math 1011/Math 1012-Grade', 'Math 1201-Instructor', 'Math 1201-Grade','Math 1711/Math 1206-Instructor','Math 1711/Math 1206-Grade',
                'Math 1716/Math 2501-Instructor', 'Math 1716/Math 2501-Grade', 'Additional CIS Course1-Instructor', 'Additional CIS Course1-Grade', 'Additional CIS Course2-Instructor', 'Additional CIS Course2-Grade',
                'Additional CIS Course3-Instuctor', 'Additional CIS Course3-Grade','Elective Class1-Instructor','Elective Class1-Grade','Elective Class2-Instructor','Elective Class2-Grade','Other1-Instructor','Other1-Grade',
                'Other2-Instructor','Other2-Grade','Other3-Instructor','Other3-Grade','Other4-Instructor','Other4-Grade','Other5-Instructor','Other5-Grade' ]
        dict_writer = csv.DictWriter(output_file, head)
        dict_writer.writeheader()
        for data1 in data:
            dict_writer.writerow(data1)

    print("Saved results to {}".format(output_filename))


# save Information System student's course information
def save_to_IS(data, output_filename):
    with open(output_filename, 'w', newline='') as output_file:

        head = ['First Name', 'Last Name', 'CUNY ID', 'Email', 'Final Semester', 'CISC 1115/CISC 1170-Instructor', 'CISC 1115/CISC 1170-Grade','CISC 3115-Instructor', 'CISC 3115-Grade', 'CISC 3130-Instructor', 'CISC 3130-Grade', 'CISC 3810-Instructor', 'CISC 3810-Grade',
                'CISC 4900/CISC 5001-Instructor', 'CISC 4900/CISC 5001-Grade', 'Other CIS Course1-Instructor', 'Other CIS Course1-Grade', 'Other CIS Course2-Instructor', 'Other CIS Course2-Grade', 
                'Other CIS Course3-Instuctor', 'Other CIS Course3-Grade', 'CISC 2820W/Phil 3318W-Instructor', 'CISC 2820W/Phil 3318W-Grade','Business 3420/CISC 1590-Instructor','Business 3420/CISC 1590-Grade','Business 3430/CISC 2531-Instructor','Business 3430/CISC 2531-Grade',
                'Business 3120/CISC 1530/Business 3432/CISC 2532-Instructor', 'Business 3120/CISC 1530/Business 3432/CISC 2532-Grade', 'Business 4202W/CISC 1580W-Instructor', 'Business 4202W/CISC 1580W-Grade', 'Economics 2100/Business 2100-Instructor', 'Economics 2100/Business 2100-Grade',
                'Economics 2200/Business 2200-Instructor', 'Economics 2200/Business 2200-Grade', 'Business 3200-Instructor', 'Business 3200-Grade', 'Finance 3310-Instructor', 'Finance 3310-Grade', 'Accounting 2001-Instructor', 'Accounting 2001-Grade',
                'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Instructor', 'Business 3400/Economics 3400/Math 2501/Math 3501/Psychology 3400/Business 3410/Math 1201/Business 3421/CISC 2590-Grade',
                'Other1-Instructor', 'Other1-Grade','Other2-Instructor', 'Other2-Grade', 'Other3-Instructor', 'Other3-Grade', 'Other4-Instructor', 'Other4-Grade', 'Other5-Instructor', 'Other5-Grade']
        dict_writer = csv.DictWriter(output_file, head)
        dict_writer.writeheader()
        for data1 in data:
            dict_writer.writerow(data1)
    print("Saved results to {}".format(output_filename))




# ----------------------------------------
# Program entry point
# ----------------------------------------
# Accept user supplied arguments
# -d  allows user to specify directory to perform extraction

parser = argparse.ArgumentParser(description='Text goes here.')
parser.add_argument('-d', '--directory', metavar='D', required=False, default=os.getcwd(),
                    help="directory path where files are located", action="store")
parser.add_argument('-o', '--output-directory', metavar='O', required=False,
                    action="store", help="output directory")
parser.add_argument('-f', '--filename', default="_report", required=False, action="store",
                    help='give a filename to output. otherwise it will be _report_YYYY-MM-DD.csv')

args = parser.parse_args()

# start from main method
# main(path=args.directory,
#      out_path=vars(args).get("output-directory", ""),
#      filename=args.filename)

main("", "", "")
