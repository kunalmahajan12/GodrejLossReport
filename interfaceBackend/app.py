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


def Banding1(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from banding1_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_banding1_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select banding1_floattable.DateAndTime, banding1_floattable.val, banding1_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from banding1_floattable left join `fault reason mapping new` on banding1_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_banding1_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from banding1_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_banding1_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("banding1")
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
            cursor.execute("Select banding1_floattable.DateAndTime, banding1_floattable.val, banding1_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from banding1_floattable left join `fault reason mapping new` on banding1_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_banding1_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from banding1_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_banding1_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("banding1")
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
            cursor.execute("Select banding1_floattable.DateAndTime, banding1_floattable.val, banding1_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from banding1_floattable left join `fault reason mapping new` on banding1_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_banding1_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from banding1_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_banding1_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("banding1")
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
            cursor.execute("Select banding1_floattable.DateAndTime, banding1_floattable.val, banding1_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from banding1_floattable left join `fault reason mapping new` on banding1_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_banding1_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def Banding2(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from banding2_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_banding2_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select banding2_floattable.DateAndTime, banding2_floattable.val, banding2_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from banding2_floattable left join `fault reason mapping new` on banding2_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_banding2_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from banding2_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_banding2_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("banding2")
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
            cursor.execute("Select banding2_floattable.DateAndTime, banding2_floattable.val, banding2_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from banding2_floattable left join `fault reason mapping new` on banding2_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_banding2_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from banding2_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_banding2_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("banding2")
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
            cursor.execute("Select banding2_floattable.DateAndTime, banding2_floattable.val, banding2_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from banding2_floattable left join `fault reason mapping new` on banding2_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_banding2_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from banding2_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_banding2_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("banding2")
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
            cursor.execute("Select banding2_floattable.DateAndTime, banding2_floattable.val, banding2_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from banding2_floattable left join `fault reason mapping new` on banding2_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_banding2_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def Cutter3(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from cutter3_floattable where tagIndex = 3 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_Cutter3_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select cutter3_floattable.DateAndTime, cutter3_floattable.val, cutter3_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from cutter3_floattable left join `fault reason mapping new` on cutter3_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 10 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_Cutter3_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from cutter3_floattable where tagIndex = 3 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_cutter3_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("cutter3")
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
            cursor.execute("Select cutter3_floattable.DateAndTime, cutter3_floattable.val, cutter3_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from cutter3_floattable left join `fault reason mapping new` on cutter3_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 10 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_Cutter3_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from cutter3_floattable where tagIndex = 3 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_cutter3_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("cutter3")
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
            cursor.execute("Select cutter3_floattable.DateAndTime, cutter3_floattable.val, cutter3_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from cutter3_floattable left join `fault reason mapping new` on cutter3_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 10 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_Cutter3_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from cutter3_floattable where tagIndex = 3 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_cutter3_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("cutter3")
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
            cursor.execute("Select cutter3_floattable.DateAndTime, cutter3_floattable.val, cutter3_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from cutter3_floattable left join `fault reason mapping new` on cutter3_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 10 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_Cutter3_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def Cutter4(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from cutter4_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_cutter4_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select cutter4_floattable.DateAndTime, cutter4_floattable.val, cutter4_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from cutter4_floattable left join `fault reason mapping new` on cutter4_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_cutter4_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from cutter4_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_cutter4_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("cutter4")
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
            cursor.execute("Select cutter4_floattable.DateAndTime, cutter4_floattable.val, cutter4_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from cutter4_floattable left join `fault reason mapping new` on cutter4_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_cutter4_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from cutter4_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_cutter4_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("cutter4")
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
            cursor.execute("Select cutter4_floattable.DateAndTime, cutter4_floattable.val, cutter4_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from cutter4_floattable left join `fault reason mapping new` on cutter4_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_cutter4_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from cutter4_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_cutter4_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("cutter4")
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
            cursor.execute("Select cutter4_floattable.DateAndTime, cutter4_floattable.val, cutter4_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from cutter4_floattable left join `fault reason mapping new` on cutter4_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_cutter4_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def Stamper3(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from stamper3_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_stamper3_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select stamper3_floattable.DateAndTime, stamper3_floattable.val, stamper3_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from stamper3_floattable left join `fault reason mapping new` on stamper3_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_stamper3_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from stamper3_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_stamper3_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
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
            cursor.execute("Select stamper3_floattable.DateAndTime, stamper3_floattable.val, stamper3_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from stamper3_floattable left join `fault reason mapping new` on stamper3_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_stamper3_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from stamper3_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_stamper3_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
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
            cursor.execute("Select stamper3_floattable.DateAndTime, stamper3_floattable.val, stamper3_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from stamper3_floattable left join `fault reason mapping new` on stamper3_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_stamper3_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from stamper3_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_stamper3_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
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
            cursor.execute("Select stamper3_floattable.DateAndTime, stamper3_floattable.val, stamper3_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from stamper3_floattable left join `fault reason mapping new` on stamper3_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_stamper3_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def Stamper4(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from stamper4_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_stamper4_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select stamper4_floattable.DateAndTime, stamper4_floattable.val, stamper4_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from stamper4_floattable left join `fault reason mapping new` on stamper4_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_stamper4_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from stamper4_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_stamper4_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("stamper4")
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
            cursor.execute("Select stamper4_floattable.DateAndTime, stamper4_floattable.val, stamper4_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from stamper4_floattable left join `fault reason mapping new` on stamper4_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_stamper4_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from stamper4_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_stamper4_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("stamper4")
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
            cursor.execute("Select stamper4_floattable.DateAndTime, stamper4_floattable.val, stamper4_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from stamper4_floattable left join `fault reason mapping new` on stamper4_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_stamper4_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from stamper4_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_stamper4_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("stamper4")
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
            cursor.execute("Select stamper4_floattable.DateAndTime, stamper4_floattable.val, stamper4_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from stamper4_floattable left join `fault reason mapping new` on stamper4_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_stamper4_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def Wrapper5(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper5_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper5_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select wrapper5_floattable.DateAndTime, wrapper5_floattable.val, wrapper5_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper5_floattable left join `fault reason mapping new` on wrapper5_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_wrapper5_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper5_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper5_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper5")
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
            cursor.execute("Select wrapper5_floattable.DateAndTime, wrapper5_floattable.val, wrapper5_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper5_floattable left join `fault reason mapping new` on wrapper5_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper5_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper5_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper5_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper5")
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
            cursor.execute("Select wrapper5_floattable.DateAndTime, wrapper5_floattable.val, wrapper5_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper5_floattable left join `fault reason mapping new` on wrapper5_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper5_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper5_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper5_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper5")
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
            cursor.execute("Select wrapper5_floattable.DateAndTime, wrapper5_floattable.val, wrapper5_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper5_floattable left join `fault reason mapping new` on wrapper5_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper5_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def Wrapper6(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper6_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper6_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select wrapper6_floattable.DateAndTime, wrapper6_floattable.val, wrapper6_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper6_floattable left join `fault reason mapping new` on wrapper6_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_wrapper6_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper6_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper6_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper6")
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
            cursor.execute("Select wrapper6_floattable.DateAndTime, wrapper6_floattable.val, wrapper6_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper6_floattable left join `fault reason mapping new` on wrapper6_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper6_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper6_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper6_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper6")
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
            cursor.execute("Select wrapper6_floattable.DateAndTime, wrapper6_floattable.val, wrapper6_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper6_floattable left join `fault reason mapping new` on wrapper6_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper6_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper6_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper6_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper6")
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
            cursor.execute("Select wrapper6_floattable.DateAndTime, wrapper6_floattable.val, wrapper6_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper6_floattable left join `fault reason mapping new` on wrapper6_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper6_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def Wrapper7(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper7_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper7_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select wrapper7_floattable.DateAndTime, wrapper7_floattable.val, wrapper7_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper7_floattable left join `fault reason mapping new` on wrapper7_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_wrapper7_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper7_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper7_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper7")
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
            cursor.execute("Select wrapper7_floattable.DateAndTime, wrapper7_floattable.val, wrapper7_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper7_floattable left join `fault reason mapping new` on wrapper7_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper7_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper7_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper7_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper7")
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
            cursor.execute("Select wrapper7_floattable.DateAndTime, wrapper7_floattable.val, wrapper7_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper7_floattable left join `fault reason mapping new` on wrapper7_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper7_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper7_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper7_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper7")
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
            cursor.execute("Select wrapper7_floattable.DateAndTime, wrapper7_floattable.val, wrapper7_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper7_floattable left join `fault reason mapping new` on wrapper7_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper7_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def Wrapper8(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper8_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper8_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select wrapper8_floattable.DateAndTime, wrapper8_floattable.val, wrapper8_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper8_floattable left join `fault reason mapping new` on wrapper8_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_wrapper8_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper8_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper8_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper8")
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
            cursor.execute("Select wrapper8_floattable.DateAndTime, wrapper8_floattable.val, wrapper8_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper8_floattable left join `fault reason mapping new` on wrapper8_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper8_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper8_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper8_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper8")
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
            cursor.execute("Select wrapper8_floattable.DateAndTime, wrapper8_floattable.val, wrapper8_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper8_floattable left join `fault reason mapping new` on wrapper8_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper8_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from wrapper8_floattable where tagIndex = 4 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_wrapper8_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("wrapper8")
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
            cursor.execute("Select wrapper8_floattable.DateAndTime, wrapper8_floattable.val, wrapper8_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from wrapper8_floattable left join `fault reason mapping new` on wrapper8_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 11 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_wrapper8_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    return ""

def MPC(ReportType,Date,Shift,Week,Year,Month):
    print (Date)
    # format = '%c'
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
    # Date=datetime.datetime.strptime(Date,format)
    if (ReportType=='Shift'):
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

        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from mpc_floattable where tagIndex = 1 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        d=Date.date()
        workbook= xlsxwriter.Workbook(f"Reports/Report_mpc_{shift}_{d}.xlsx")
        worksheet=workbook.add_worksheet("stamper3")
        worksheet.write(0,0,'Date')
        worksheet.write(0,1,'Shift')
        worksheet.write(0,2,'Loss Type')
        worksheet.write(0,3,'Loss Reason')
        worksheet.write(0,4,'SKU')
        worksheet.write(0,5,'Stop Date/Time')
        worksheet.write(0,6,'Start Date/Time')
        worksheet.write(0,7,'Duration')
        idx=0
        for row in result:
            cursor.execute("Select mpc_floattable.DateAndTime, mpc_floattable.val, mpc_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from mpc_floattable left join `fault reason mapping new` on mpc_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 8 and val <> 0",(row[1],))
            Description=cursor.fetchone()
            print(Description)
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
        path=downloadDirectory+"/Report_mpc_"+shift+"_"+str(d)+".xlsx"
        return send_file(path)

    elif (ReportType=='Daily'):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from mpc_floattable where tagIndex = 1 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_mpc_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("mpc")
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
            cursor.execute("Select mpc_floattable.DateAndTime, mpc_floattable.val, mpc_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from mpc_floattable left join `fault reason mapping new` on mpc_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 8 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_mpc_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    elif (ReportType=='Weekly'):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from mpc_floattable where tagIndex = 1 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        workbook= xlsxwriter.Workbook(f"Reports/Report_mpc_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("mpc")
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
            cursor.execute("Select mpc_floattable.DateAndTime, mpc_floattable.val, mpc_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from mpc_floattable left join `fault reason mapping new` on mpc_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 8 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_mpc_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
    else:
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        # Date+= timedelta(days=6-Date.weekday())
        # Date+= timedelta(days=(Week-1)*7)
        # spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        # stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        # spt= spt.replace(tzinfo=None)
        print(spt,stt)
        shiftcount=-1
        cursor.execute("""select startTime, stopTime, timediff(startTime,stopTime) as duration 
        from (
            select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
            from mpc_floattable where tagIndex = 1 
        ) time
        where val=0 and Status <> 'U'and stopTime between (%s) and (%s);""",(spt,stt))
        result=cursor.fetchall()
        print(result)
        workbook= xlsxwriter.Workbook(f"Reports/Report_mpc_{spt.date()}_{stt.date()}.xlsx")
        worksheet=workbook.add_worksheet("mpc")
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
            cursor.execute("Select mpc_floattable.DateAndTime, mpc_floattable.val, mpc_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from mpc_floattable left join `fault reason mapping new` on mpc_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= (%s) and TagIndex = 8 and val <> 0",(row[1],))
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
        path=downloadDirectory+"/Report_mpc_"+str(spt.date())+"_"+str(stt.date())+".xlsx"
        return send_file(path)
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
    if (Machine=='Banding1'):
        return Banding1(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Banding2'):
        return Banding2(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Cutter3'):
        return Cutter3(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Cutter4'):
        return Cutter4(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Stamper3'):
        return Stamper3(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Stamper4'):
        return Stamper4(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Wrapper5'):
        return Wrapper5(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Wrapper6'):
        return Wrapper6(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Wrapper7'):
        return Wrapper7(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Wrapper8'):
        return Wrapper8(ReportType,Date,Shift,Week,Year,Month)
    elif (Machine=='Mpc'):
        return MPC(ReportType,Date,Shift,Week,Year,Month)
    # return jsonify(dictToReturn)
    
    # except FileNotFoundError:
    #     abort(404)
if __name__ == "__main__":
    app.run(debug=True)