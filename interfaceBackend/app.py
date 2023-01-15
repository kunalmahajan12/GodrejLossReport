from flask import Flask,request,jsonify,send_file
from flask_cors import CORS, cross_origin

import os
import mysql.connector
import datetime
from datetime import timedelta
import io
import xlsxwriter
import dateutil.parser
from string import Template

downloadDirectory="/Users/kunalmahajan/College 4th year/freelance project/Interface/interfaceBackend/Reports"

class DeltaTemplate(Template):
    delimiter = '%'


def strfdelta(tdelta, fmt):
    d = {}
    l = {'D': 86400, 'H': 3600, 'M': 60, 'S': 1}
    rem = int(tdelta.total_seconds())

    for k in ( 'D', 'H', 'M', 'S' ):
        if "%{}".format(k) in fmt:
            d[k], rem = divmod(rem, l[k])

    t = DeltaTemplate(fmt)
    return t.substitute(**d)


mydb=mysql.connector.connect(host='localhost',
                            user='root',
                            passwd="MS16ct40$",
                            database='Godrej',
                            auth_plugin="mysql_native_password"
                            )

cursor=mydb.cursor(buffered=True)

app = Flask(__name__)
cors = CORS(app, resources={r"*": {"origins": "*"}})
app.config['CORS_HEADERS'] = 'Content-Type'

@app.route("/")
def home():
    print("Hello flask")
    return "Hello flask"

def ReportGenerator(machine,spt,stt,ReportName,StatusTag,AlarmTag,st):
    print(spt.strftime('%x %X'),stt)
    shiftcount=-1
    query=f"""select startTime, stopTime, timediff(startTime,stopTime) as duration from (
        select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
        from {machine}_floattable where tagIndex = {StatusTag}
    ) time where val=0 and Status <> 'U' and startTime is NOT NULL and stopTime between '{spt}' and '{stt}';"""
    cursor.execute(query)
    result=cursor.fetchall()
    print(result)
    workbook= xlsxwriter.Workbook(f"Reports/{ReportName}")
    worksheet=workbook.add_worksheet(f"{machine}")
    worksheet.write(0,0,'Date')
    worksheet.write(0,1,'Shift')
    worksheet.write(0,2,'Loss Type')
    worksheet.write(0,3,'Loss Reason')
    worksheet.write(0,4,'SKU')
    worksheet.write(0,5,'Stop Date/Time')
    worksheet.write(0,6,'Start Date/Time')
    worksheet.write(0,7,'Duration')
    idx=0
    shift=""
    for row in result:
        if (st=="var"):
            tup=row[1]-spt
            shiftsize= timedelta(hours=8, minutes=0, seconds=0)
            counter=int(tup/shiftsize)
            if (shiftcount==-1):
                shiftcount=counter
            elif (shiftcount==counter):
                shiftcount=counter
            elif (shiftcount!=counter):
                idx+=1
                shiftcount=counter
            
            
            if (shiftcount%3==0):
                shift="A"
            elif (shiftcount%3==1):
                shift="B"
            else:
                shift="C"
        else:
            shift=st


        query=f"Select {machine}_floattable.DateAndTime, {machine}_floattable.val, {machine}_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from {machine}_floattable left join `fault reason mapping new` on {machine}_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= '{row[1]}' and TagIndex = {AlarmTag} and val <> 0"
        cursor.execute(query)
        Description=cursor.fetchone()
        cursor.execute("select oee_stringtable.val from oee_stringtable where tagIndex=9 and DateAndTime <= (%s) order by DateAndTime desc limit 1",(Description[0],))
        sku=cursor.fetchone()
        if (Description[0]-row[1]>timedelta(hours=0, minutes=10, seconds=0)):
            continue
        worksheet.write(idx+1,0,Description[0].strftime('%x %X'))
        worksheet.write(idx+1,1,shift)
        worksheet.write(idx+1,2,Description[4])
        worksheet.write(idx+1,3,Description[3])
        worksheet.write(idx+1,4,sku[0])
        worksheet.write(idx+1,5,row[1].strftime('%x %X'))
        worksheet.write(idx+1,6,row[0].strftime('%x %X'))
        worksheet.write(idx+1,7,strfdelta(row[2],'%H:%M:%S'))
        idx+=1
    workbook.close()
    path=downloadDirectory+f"/{ReportName}"
    return send_file(path)

def summary(machine,ReportType,Shift,Date,Week,Year,Month,statusTag,alarmTag):
    if (ReportType=="Shift"):
        if (Shift=="ShiftA"):
            shift='A'
            spt=Date+timedelta(hours=6, minutes=0, seconds=0)
            stt=Date+timedelta(hours=14, minutes=0, seconds=0)
        elif (Shift == "ShiftB"):
            shift='B'
            spt=Date+timedelta(hours=14, minutes=0, seconds=0)
            stt=Date+timedelta(hours=22, minutes=0, seconds=0)
        else:
            shift='C'
            spt=Date+timedelta(hours=22, minutes=0, seconds=0)
            stt=Date+timedelta(hours=30, minutes=0, seconds=0)
        d=Date.date
        ReportName=f"Report_{machine}_{shift}_{d}.xlsx"
        spt= spt.replace(tzinfo=None)
        return ReportGenerator(machine,spt,stt,ReportName,statusTag,alarmTag,shift)
    elif(ReportType=="Daily"):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        ReportName=f"Report_{machine}_{Date}.xlsx"
        return ReportGenerator(machine, spt,stt,ReportName,statusTag,alarmTag,"var")
    elif (ReportType=="Weekly"):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        ReportName=f"Report_{machine}_{Year}_Week{Week}.xlsx"
        return ReportGenerator(machine,spt,stt,ReportName,statusTag,alarmTag,"var")
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        ReportName=f"Report_{Month}{Year}.xlsx"
        return ReportGenerator(machine,spt,stt,ReportName,statusTag,alarmTag,"var")
    return ""

@app.route("/lossReport", methods=["POST"])
@cross_origin(origin='*',headers=['Content-Type','Authorization'])
def report():
    input_json = request.get_json(force=True)
    Line = input_json['Line']
    Report=input_json['Report']
    ReportType=input_json['ReportType']
    Machine=input_json['Machine']
    Date=input_json['Date']
    Shift=input_json['Shift']
    Week=input_json['Week']
    Month=input_json['Month']
    Year=input_json['Year']
    Date=dateutil.parser.isoparse(Date)
    Date+= timedelta(hours=5,minutes=30,seconds=0)
    Year=dateutil.parser.isoparse(Year)
    Year+= timedelta(hours=5,minutes=30,seconds=0)
    Year=Year.year
    Month=dateutil.parser.isoparse(Month)
    Month+= timedelta(hours=5,minutes=30,seconds=0)
    Month=Month.month
    print(Month)
    Week=int(Week)
    if (Machine=='Banding1'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,4,11)
    elif (Machine=='Banding2'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,4,11)
    elif (Machine=='Cutter3'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,3,10)
    elif (Machine=='Cutter4'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,4,11)
    elif (Machine=='Stamper3'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,4,11)
    elif (Machine=='Stamper4'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,4,11)
    elif (Machine=='Wrapper5'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,4,11)
    elif (Machine=='Wrapper6'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,4,11)
    elif (Machine=='Wrapper7'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,4,11)
    elif (Machine=='Wrapper8'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,4,11)
    elif (Machine=='Mpc'):
        return summary(Machine.lower(),ReportType,Shift,Date,Week,Year,Month,1,8)
    # return jsonify(dictToReturn)
    
    # except FileNotFoundError:
    #     abort(404)
if __name__ == "__main__":
    app.run(debug=True)