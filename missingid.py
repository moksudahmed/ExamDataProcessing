
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


def UpdateDataToDatabase(id,moduleRegi, offeredID):

    try:
                    cur = myConnection.cursor()
                    sql_update_query = """Update x_tbl_tmp_marks_eng set "moduleRegistrationNo" = %s where "studentID" = %s AND "offeredModuleID"=%s"""
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
          sql_query = """ SELECT "studentID","courseCode","offeredModuleID" FROM x_tbl_tmp_marks_eng"""
          cur.execute(sql_query, ())
          query_results = cur.fetchall()
          for row in  query_results:
             #print(row[0], row[2])
             moduleRegistrationID = getModuleRegistrationByOfferedID(row[0], row[2]);
             if moduleRegistrationID is not None:
                 print(row[0], getModuleRegistrationByOfferedID(row[0], row[2]))
                 UpdateDataToDatabase(row[0],moduleRegistrationID, row[2])


getRemainingId()
myConnection.close()
