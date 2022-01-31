from configparser import ConfigParser
import re
import csv
import sqlite3
import psycopg2
import sys
from termcolor import colored, cprint

data=[]
missing = []
n = 0
try:
    myConnection = psycopg2.connect( host="localhost", port="5432", user="postgres", password="dream@3030.com", dbname="db_muERPv7p0" )
except psycopg2.DatabaseError:
    sys.exit('Failed to connect to database')

def settings():
    host = input("Enter HOST:")
    dbname = input("Enter DB NAME:")
    password = input("Enter password:")
    myConnection.close()
    myConnection = psycopg2.connect( host=host, port="5432", user="postgres", password=password, dbname=dbname )
    print("Connected!")


def getProgramCode():
          cur = myConnection.cursor()
          sql_query = """SELECT "programmeCode", pro_name, "pro_shortName" FROM tbl_c_programme"""
          cur.execute(sql_query, ())
          query_results = cur.fetchall()
          for row in  query_results:
             print (row)
          return 1

def getCourseOfferedNo(year,term, batch, sec):

        prog = input("Enter Programme Code:")
        code = input("Enter Course Code:")
        try:
            cur = myConnection.cursor()
            sql_query = """SELECT DISTINCT
                          o."offeredModuleID",
                          m.mod_name
                        FROM
                          public.tbl_h_offeredmodule o,
                          public.tbl_e_module m
                        WHERE
                          o."moduleCode" = m."moduleCode" AND
                          o."syllabusCode" = m."syllabusCode" AND
                          o."moduleCode" = %s AND
                          o."programmeCode" = %s AND
                          o.ofm_year = %s AND
                          o.ofm_term = %s AND
                          o."batchName" = %s AND
                          o."sectionName" = %s"""
        #    print(sql_query)
            cur.execute(sql_query, (code, str(prog), year,term, batch,sec))
            #query_results = cur.fetchall()
            row = cur.fetchone()
            if row == None:
                    print("Course not found!")
                    return 0
            else:
                    print('================================================================================')
                    print (row[0])
                    print('================================================================================')
                    return row[0]

        except:
            print("Invalid Programme Code")

def doQuery(sid, year, term, offeredID):

    cur = myConnection.cursor()

    #Id = input('Enter location: ')
    #cur.execute( "SELECT * FROM tbl_o_student WHERE 'programmeCode' = '111'" )
    sql_query = """SELECT
                  m."moduleRegistrationID",
                  m."termAdmissionID",
                  m."offeredModuleID",
                  m.reg_date,
                  m.reg_attendence,
                  m."reg_classTest",
                  m.reg_midterm,
                  m."markingSchemeID",
                  m.reg_type,
                  m."reg_finalExamReg",
                  m."reg_finalExamRegDate",
                  m."reg_suppleExamReg",
                  m."reg_suppleExamRegDate",
                  m.reg_status
                FROM
                  public.tbl_q_termadmission t,
                  public.tbl_s_moduleregistration m
                WHERE
                  t."termAdmissionID" = m."termAdmissionID" AND
                  t."studentID" = %s AND
                  t.tra_year = %s AND
                  t.tra_term = %s AND
                  m."offeredModuleID" = %s """
#SELECT * FROM tbl_q_termadmission WHERE "studentID" = %s

    cur.execute(sql_query, (sid,year,term,offeredID))
    query_results = cur.fetchall()
    #print(query_results[0][0])
    if len(query_results) == 0:
        return 0
    else:
        return query_results

def getCourseOfferIno():
    with open('resultsheet.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                array = re.findall(r'[0-9]+', row[0])
                year = array[0]
            elif line_count == 2:
                prog = row[1]
            elif line_count == 3:
                batch = row[1]
            elif line_count == 4:
                sec = row[1]
            elif line_count == 5:
                code = row[1]
            elif line_count == 6:
                title = row[1]
            elif line_count == 7:
                fname = row[1]
            else:
                phoneNumRegex = re.compile(r'\d\d\d-\d\d\d-\d\d\d')
                mo = phoneNumRegex.search(row[0])
            line_count += 1
    #print (year, prog, batch, sec, code, title, fname)
    #courseName = getCourseOfferedNo(  year, 1, batch, sec)
    #print(courseName)
    return getCourseOfferedNo(  year, 1, batch, sec)

def missingId():
    with open('missing_id.csv', 'w', newline='') as file:
        fieldnames = ['Student_id']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for i in range(0, len(missing)):
            writer.writerow({'Student_id': missing[i]})

def displayData():
    print("======================================================================")
    print("Student ID       Att.      CT.       MID.         FInal")
    print("======================================================================")

    for i in range(0,len(data)):
    #    print()
        for j in range(0,5):
             print(data[i][j], end = "   ")
        print()
    print("======================================================================")

def displayMissingId():
    print("======================================================================")
    for i in range(0,len(missing)):
             print(missing[i])
    print()
    print("======================================================================")

def SaveDataToDatabase(data, n, prog,batch,sec, code):
    for i in range(1,len(data)):
    #    print()
            postgres_insert_query = """INSERT INTO tbl_tmp_marks(
                                             "studentID","programCode", "batchName", "sectionNo",
                                             "courseCode", att, ct, mid, final)
                                     VALUES ( %s, %s, %s,%s, %s,%s, %s, %s, %s)"""
            record_to_insert = (data[i][0] ,prog, batch, sec,code, data[i][1], data[i][2], data[i][3], data[i][4])
            cur = myConnection.cursor()
            cur.execute(postgres_insert_query, record_to_insert)
            myConnection.commit()

#            sql_update_query = """UPDATE tbl_s_moduleregistration
#                                  SET reg_attendence=%s, "reg_classTest"=%s, reg_midterm=%s WHERE "moduleRegistrationID"=%s"""
#            cur = myConnection.cursor()
#            cur.execute(sql_update_query, (data[i][2], data[i][3], data[i][4] ,data[i][0]))

            #sql_update_query_final = """UPDATE tbl_u_exammarks SET emr_mark=%s WHERE "moduleRegistrationID"=%s"""

            #cur = myConnection.cursor()
            #cur.execute(sql_update_query_final, (data[i][5],data[i][0]))
            #myConnection.commit()
    print("Record inserted successfully")

#doQuery()


#conn = sqlite3.connect('exam-final.sqlite')
#cur = conn.cursor()

#cur.execute('''
#DROP TABLE IF EXISTS Results''')

#cur.execute('''
#CREATE TABLE Results (id TEXT, att INTEGER, ct INTEGER, mid INTEGER, final INTEGER

def CSVDataProcess():

    #fileName = input("Enter File Name:")
    fileName = "resultsheet.csv"
    temp=[]

    mis = 0
    count = 0
#    try:
    with open(fileName) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            #courseOfferedID = getCourseOfferIno()
            for row in csv_reader:
                line_count += 1
                if line_count == 1:
                    array = re.findall(r'[0-9]+', row[0])
                    year = array[0].strip()
                elif line_count == 3:
                    prog = row[1].strip()
                elif line_count == 4:
                    array = re.findall(r'[0-9]+', row[1])
                    batch = array[0].strip()
                elif line_count == 5:
                    sec = row[1].strip()
                elif line_count == 6:
                    code = row[1].strip()
                elif line_count == 7:
                    title = row[1].strip()
                elif line_count == 8:
                    fname = row[1].strip()
                else:
                    idNumRegex = re.compile(r'\d\d\d-\d\d\d-\d\d\d')
                    mo = idNumRegex.search(row[0])

                    try:
                        if mo != None:
                            x =  mo.group()
                            if len(x) == 11:
                                    id = row[0].strip()
                                    att = row[1].strip()
                                    ct = row[2].strip()
                                    mid = row[3].strip()
                                    sixty = row[4].strip()
                                    final = row[5].strip()
                                    total = row[6].strip()
                                    temp=[]

                                    temp.append(id)
                                    temp.append(att)
                                    temp.append(ct)
                                    temp.append(mid)
                                    temp.append(final)
                                    data.append(temp)
                                    count +=1

                        #if len(x) == 11:

                    except:
                        print("Invalid id")
                        continue
    print("Program:", prog)
    print("Batch:", batch)
    print("Section", sec)
    print("Course Code:",code)
    print("Course Title:",title)
    displayData()
    print("Preview data:")
    print("Is it okey. Yes to continue or do you want to input program code and course code?")
    c = choice = input("Do you want to save it (yes / not):")
    try:
            if choice.lower() == "yes" or choice.lower() == "y":
                       prog = input("Enter Programme Code:")
                       code = input("Enter Course Code:")
    except:
            print("Invalid choice")
    displayData()
    choice = input("Do you want to save it (yes / not):")
    if choice.lower() == "yes" or choice.lower() == "y":
            SaveDataToDatabase(data, count, prog, batch, sec, code)
            missingId()
    else:
            print("You have canceled")

    #        print("Invalid choice")
    n = count
    print(f'Processed {line_count} lines.')
    #except:
    #        print("Bad file name!")

            #conn.commit()
            #for i in range(0,3):

    #print("missing", missing)
    #print("Processed data")
    #displayData()

while True:
    print("*******************************************")
    print('*  1. Display Program List                *')
    print('*  2. Read Data                           *')
    print('*  3. Process Data                        *')
    print('*  4. Missing ids                         *')
    print('*  5. Settings                            *')

    print('*  0. Exit                                *')
    print("*******************************************")
    CSVDataProcess()
    choice = input("Enter your choice:")
    if choice == "0" : break
    #elif choice == "1": getProgramCode()
    #elif choice == "2": displayData()
    #elif choice == "3": CSVDataProcess()
    #elif choice == "4": displayMissingId();
    #elif choice == "5": settings();
    #else:
#        print("Invalid input")


myConnection.close()
#cur.close()
