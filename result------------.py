
# Reading an excel file using Python
import io
import xlrd
from configparser import ConfigParser
import re
import csv
import sqlite3
import psycopg2
import sys
from termcolor import colored, cprint
import os

missing = []
n = 0


try:
    myConnection = psycopg2.connect( host="localhost", port="8085", user="postgres", password="dream@3030.com", dbname="db_muERPv7p2" )
except psycopg2.DatabaseError:
    sys.exit('Failed to connect to database')


#f= open("resultlog.txt","w+")

def getProgramCode():
          cur = myConnection.cursor()
          sql_query = """SELECT "programmeCode", pro_name, "pro_shortName" FROM tbl_c_programme"""
          cur.execute(sql_query, ())
          query_results = cur.fetchall()
          for row in  query_results:
             print (row)
          return 1

def UpdateDataToDatabase(id,moduleRegi, offeredID):

    try:
                    cur = myConnection.cursor()
                    sql_update_query = """Update x_tbl_tmp_marks_eee set "moduleRegistrationNo" = %s where "studentID" = %s AND "offeredModuleID"=%s"""
                    cur.execute(sql_update_query, (moduleRegi, id, offeredID))
                    myConnection.commit()
                    count = cur.rowcount
                    print(count, "Record Updated successfully ")
    except (Exception, psycopg2.DatabaseError) as error:
            #if(myConnection):
                with io.open("resultlog.txt", mode='a', encoding='utf-8') as f:
                    f.write(str(error))
                    f.write("\n")
                print("Failed to insert record into tbl_tmp_marks table", error)


def getModuleRegistrationByOfferedID(id, offeredID):
            cur = myConnection.cursor()
            sql_query = """SELECT m."moduleRegistrationID" FROM
                             public.tbl_q_termadmission t,
                             public.tbl_s_moduleregistration m
                           WHERE
                             t."termAdmissionID" = m."termAdmissionID" AND
                             t."studentID" = %s AND
                             m."offeredModuleID" =%s"""
        #    print(sql_query)
            cur.execute(sql_query, (id, offeredID))
            #query_results = cur.fetchall()
            row = cur.fetchone()
            if row == None:
                    return None
            else:
                    return row[0]

def getRemainingId():
          cur = myConnection.cursor()
          sql_query = """ SELECT "studentID","courseCode","offeredModuleID" FROM x_tbl_tmp_marks_eee"""
          cur.execute(sql_query, ())
          query_results = cur.fetchall()
          for row in  query_results:
             #print(row[0], row[2])
             moduleRegistrationID = getModuleRegistrationByOfferedID(row[0], row[2]);
             if moduleRegistrationID is not None:
                 print(row[0], getModuleRegistrationByOfferedID(row[0], row[2]))
                 UpdateDataToDatabase(row[0],moduleRegistrationID, row[2])

def getCourseOfferedNo(prog, batch, sec, code):
            cur = myConnection.cursor()
            sql_query = """SELECT
                          o."offeredModuleID"
                        FROM
                          public.tbl_h_offeredmodule o
                        WHERE
                          o."programmeCode" = %s AND
                          o."batchName" = %s AND
                          o."sectionName" = %s AND
                          o."moduleCode" = %s AND
                          o.ofm_year = 2020 AND
                          o.ofm_term = 1"""
            #print(sql_query)
            cur.execute(sql_query, (str(prog), int(batch), sec, code))
            #query_results = cur.fetchall()
            row = cur.fetchone()
            if row == None:
                    return None
            else:
                    return row[0]


def getModuleRegistration(id, code):
            cur = myConnection.cursor()
            sql_query = """SELECT
                              m."moduleRegistrationID"
                            FROM
                              public.tbl_q_termadmission t,
                              public.tbl_s_moduleregistration m,
                              public.tbl_h_offeredmodule o
                            WHERE
                              t."termAdmissionID" = m."termAdmissionID" AND
                              m."offeredModuleID" = o."offeredModuleID" AND
                              t."studentID" = %s AND
                              t.tra_term = 1 AND
                              t.tra_year = 2020 AND
                              o."moduleCode" = %s """
        #    print(sql_query)
            cur.execute(sql_query, (id, code))
            #query_results = cur.fetchall()
            row = cur.fetchone()
            if row == None:
                    return None
            else:
                    return row[0]


def checkDuplicate(data):
    d = []
    index = 0
    for i in range(0,len(data)):
        for j in range(i+1,len(data)):
            if j <= len(data):
                if data[i][0] == data[j][0]:
                    d.append(data[j][0])
                    #index += 1
    return d

def SaveDataToDatabase(data, n, prog,batch,sec, code, title, name, offeredID):
    checkDuplicate(data)
    try:

    #    f.write('hi there\n')
        if offeredID != None:
            count = 0
            for i in range(0,len(data)):
                    #print(data[i][0] ,prog, batch, sec,code, data[i][1], data[i][2], data[i][3], data[i][4],title,name)
                    postgres_insert_query = """INSERT INTO x_tbl_tmp_marks_eee(
                                                     "studentID","programCode", "batchName", "sectionNo",
                                                     "courseCode", att, ct, mid, final,absent,title,name, "offeredModuleID","moduleRegistrationNo")
                                             VALUES ( %s, %s, %s,%s, %s,%s, %s, %s,%s, %s,%s,%s,%s,%s)"""
                    record_to_insert = (data[i][0] ,prog, batch, sec,code, data[i][1], data[i][2], data[i][3], data[i][4],data[i][5],title,name, offeredID,data[i][7])
                    cur = myConnection.cursor()
                    cur.execute(postgres_insert_query, record_to_insert)
                    myConnection.commit()
                    count += 1
            print("*******************",count," Record inserted successfully**************************\n\n\n\n")
        else:
            print("**********************************************************************************")
            print("Can't save data due to offeredID is missing")
            print("**********************************************************************************\n")
    except (Exception, psycopg2.DatabaseError) as error:
            #if(myConnection):
                with io.open("resultlog.txt", mode='a', encoding='utf-8') as f:
                    f.write(str(error))
                    f.write("\n")
                print("Failed to insert record into tbl_tmp_marks table", error)


# Give the location of the file
def displayData(data):
    print("======================================================================")
    print("Student ID   Att. CT. MID. FInal")
    print("======================================================================")

    for i in range(0,len(data)):
    #    print()
        for j in range(0,8):
             print(data[i][j], end = "    ")
        print()
    print("======================================================================")
    print(i, "no of record has been read.")

def chaekVal(user_input):

    try:
        val = float(user_input)
    except ValueError:
        try:
            val = float(user_input)
        except ValueError:
            val = 0.0
    return val

def isnumber(x):
    return re.match("-?[0-9]+([.][0-9]+)?$", x) is not None

def XLSXDataProcess(sheet):

    data=[]

    sheet.cell_value(0, 0)

    #fileName = input("Enter File Name:")
    temp=[]
    offer = 0
    mis = 0
    count = 0
#    try:
    line_count = 0

    for i in range(sheet.nrows):
    #    recordCount +=1
    #    print(sheet.cell_value(i, 0),sheet.cell_value(i, 1),sheet.cell_value(i, 2), sheet.cell_value(i, 3),sheet.cell_value(i, 4), sheet.cell_value(i, 5), sheet.cell_value(i, 6))
        line_count +=1
        if line_count == 1:
                array = re.findall(r'[0-9]+', sheet.cell_value(i, 0))
                year = array[0].strip()
                print("Year:", year )
        elif line_count == 3:
                prog = int(sheet.cell_value(i, 1))
                print("Program",prog)
        elif line_count == 4:
                batch =  int(sheet.cell_value(i, 1))
                print("Batch",batch)
        elif line_count == 5:
                array = sheet.cell_value(i, 1)
                sec = array[0].strip()
                print("Sec",sec)
        elif line_count == 6:
                code = sheet.cell_value(i, 1)
                print("Code",code)
        elif line_count == 7:
                title = sheet.cell_value(i, 1)
                print("Title",title)
        elif line_count == 8:
                name = sheet.cell_value(i, 1)
                print("Name",name)
        elif line_count == 10:
                offeredID = getCourseOfferedNo(prog, batch, sec, code)
        elif line_count >=12 :
                idNumRegex = re.compile(r'\d\d\d-\d\d\d-\d\d\d')
                mo = idNumRegex.search(sheet.cell_value(i, 0))
                try:
                    if mo != None:
                        x =  mo.group()
                        if len(x) == 11:
                            id =  sheet.cell_value(i, 0).strip()
                            moduleRegi = getModuleRegistration(id, code)
                            att = chaekVal(sheet.cell_value(i, 1))
                            ct =  chaekVal(sheet.cell_value(i, 2))
                            mid = chaekVal(sheet.cell_value(i, 3))
                            sixty = chaekVal(sheet.cell_value(i, 4))
                            _final = sheet.cell_value(i, 5)

                            if isinstance(_final,str) == False:
                                absent = False
                                final = chaekVal(sheet.cell_value(i, 5))
                            else:
                                if _final.isdigit() == True:
                                    absent = False
                                    final = chaekVal(sheet.cell_value(i, 5))
                                else:
                                    absent = True
                                    final = 0.0

                            total = chaekVal(sheet.cell_value(i, 6))
                            temp=[]

                            temp.append(id)
                            temp.append(att)
                            temp.append(ct)
                            temp.append(mid)
                            temp.append(final)
                            temp.append(absent)
                            temp.append(offeredID)
                            temp.append(moduleRegi)
                            data.append(temp)
                            count +=1
                        else:
                            print("Invalid id")
                            continue
                except:
                        print("Error reading")

    #getCourseOfferedNo(2020,1,'115','CSE-469', 39)
    print("Preview data:")
    displayData(data)
    if len(checkDuplicate(data)) > 0:
        print("*************************************************************************************")
        print("Some duplicate data exist in the list, ",checkDuplicate(data))
        print("*************************************************************************************")
    #choice = input("Do you want to save it (yes / not):")
    #if choice.lower() == "yes" or choice.lower() == "y":
    SaveDataToDatabase(data, count, prog, batch, sec, code, title, name,offeredID)
#            missingId()
    #@else:
    #        print("You have canceled")

    #        print("Invalid choice")
    print(f'Processed {line_count} lines.')
#    print("Program:", prog)
#    print("Batch:", batch)
#    print("Section", sec)
#    print("Course Code:",code)
#    print("Course Title:",title)
#    displayData()
    #except:
    #        print("Bad file name!")

            #conn.commit()
            #for i in range(0,3):

    #print("missing", missing)
    #print("Processed data")
    #displayData()
#CSVDataProcess()
#def tes():


fileCount = 0

path = 'd:\\py\\results\\temp\\'
files = []

for r, d, f in os.walk(path):
	for file in f:
		if '.xls' in file:
			files.append(os.path.join(r, file))

for f in files:
	try:
		print("===============================================================")
		print(f)
		print("===============================================================")
		loc = (f)
		wb = xlrd.open_workbook(loc)
		sheet = wb.sheet_by_index(0)
		XLSXDataProcess(sheet)
	except:
		print(f)
		loc = (f)
		wb = xlrd.open_workbook(loc)
		sheet = wb.sheet_by_index(0)
		XLSXDataProcess(sheet)
#f.close()
#getRemainingId()
for f in files:
    fileCount += 1
print("Total File Processed:", fileCount,"\n")
myConnection.close()
