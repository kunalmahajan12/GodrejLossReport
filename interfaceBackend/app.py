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
import os
currentDirectory=os.path.dirname(__file__)
print (currentDirectory)
downloadDirectory=currentDirectory+"/Reports"
print(downloadDirectory)

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

def ReportGenerator(machine,spt,stt,ReportName,StatusTag,AlarmTag,st,ReportType,Line):
    print(spt.strftime('%x %X'),stt)
    shiftcount=-1
    query=f"""select startTime, stopTime, timediff(startTime,stopTime) as duration ,'mac' as origin from (
        select DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, Status 
        from {machine}_floattable where tagIndex = {StatusTag}
    ) time where val=0 and Status <> 'U' and startTime is NOT NULL and stopTime between '{spt}' and '{stt}';"""
    cursor.execute(query)
    result=cursor.fetchall()
    
    mgmtQuery=f"""select stopTime,startTime,timediff(startTime,stopTime) as duration, 'mgmt' as origin,TagIndex from (
        SELECT DateAndTime as stopTime, lead(DateAndTime) over (order by DateAndTime asc) as startTime, val, lead(val) over(order by DateAndTime asc) as LeadVal, TagIndex, lead(TagIndex) over(order by DateAndTime asc) as leadTag FROM 
        Godrej.oee_floattable where TagIndex>=11 and TagIndex<=20 and  Status<>'U') ma  where tagIndex = leadTag and val=1 and leadVal=0 and stopTime between '{spt}' and '{stt}' order by StopTime asc;
    """
    cursor.execute(mgmtQuery)
    mgmtResult=cursor.fetchall()
    result=result+mgmtResult
    result.sort()
    print(result)
    workbook= xlsxwriter.Workbook(currentDirectory+f"/Reports/{ReportName}")
    worksheet=workbook.add_worksheet(f"{machine}")
    #Header
    worksheet.insert_image('B1', currentDirectory+'/Resources/logoGodrej.png')
    cell_format = workbook.add_format({
        'align': 'center',
    })
    worksheet.set_column(1, 8, None, cell_format)
    worksheet.set_row(0,44)
    worksheet.set_column(1,1,22,cell_format)
    worksheet.set_column(2,2,16,cell_format)
    worksheet.set_column(3,3,16,cell_format)
    worksheet.set_column(4,4,30,cell_format)
    worksheet.set_column(5,5,32,cell_format)
    worksheet.set_column(6,7,22,cell_format)
    worksheet.set_column(8,8,16,cell_format)
    firstLineFormat=workbook.add_format({
        'align':'center',
        'valign':'vcenter',
        'font_size':22,
        'font_color':'blue',
        'bold':True,
        'border':1,
        'border_color':'black'
    })
    worksheet.merge_range('C1:I1',"GODREJ CONSUMER PRODUCTS LIMITED",firstLineFormat)
    border=workbook.add_format({
        'align':'center',
        'border':1,
        'border_color':'black'
    })
    mgmtlosss=workbook.add_format({
        'align':'center',
        'border':1,
        'border_color':'black',
        'bg_color': '#FFCCE6'
    })
    secondLineFormat=workbook.add_format({
        'align':'center',
        'bold':True,
        'border':1,
        'border_color':'black'
    })
    worksheet.merge_range('C2:I2',"Plot No. 6, Apparel Park cum Industrial Area, Katha PO Baddi, (Himachal Pradesh)",secondLineFormat)
    thirdLineFormat=workbook.add_format({
        'align':'center',
        'font_size':16,
        'font_color':'#800000',
        'bold':True,
        'border':1,
        'border_color':'black'
    })
    worksheet.merge_range('C3:I3',"MANUAL ENTRY/UPDATE SHOULD BE COMPLETED BEFORE 10 MINUTES OF SHIFT END",thirdLineFormat)
    fourthLineFormat=workbook.add_format({
        'align':'center',
        'font_size':16,
        'font_color':'#0080FF',
        'bold':True,
        'border':1,
        'border_color':'black'
    })
    worksheet.merge_range('B4:I4',f"Breakdown Report For Machine : {machine}",fourthLineFormat)
    descFormat=workbook.add_format({
        'align':'center',
        'font_size':12,
        'font_color':'#800000',
        'bold':True,
        'border':1,
        'border_color':'black'
    })
    worksheet.write('B5','Report Type:',descFormat)
    worksheet.write('B6','Production Line:',descFormat)
    worksheet.write('C5',f"{ReportType}",descFormat)
    worksheet.write('C6',f"{Line}",descFormat)
    worksheet.merge_range('D5:F5',"",border)
    worksheet.merge_range('D6:F6',"",border)
    worksheet.write('G5','From Date:',descFormat)
    worksheet.write('H5',f"{spt.strftime('%d-%b-%Y')}",descFormat)
    worksheet.merge_range('B7:J7',"")
    if (ReportType=='Shift'):
        worksheet.write('G6','Shift:',descFormat)
        worksheet.write('H6',f"{st}",descFormat)
    else:
        worksheet.write('G6',"",descFormat)
        worksheet.write('H6',"",descFormat)
    worksheet.write('B2',"",border)
    worksheet.write('B3',"",border)
    worksheet.write('I5',"",border)
    worksheet.write('I6',"",border)
    titleCellFormat=workbook.add_format({
        'bold':True, 
        'font_size':14,
        'font_color':'white',
        'bg_color':'#0080C0',
        'align':'center',
        'valign':'vcenter',
        'border':1,
        'border_color':'black'
    })
    # worksheet.set_row(8,40,titleCellFormat)
    worksheet.write(7,1,'Date',titleCellFormat)
    worksheet.write(7,2,'Shift',titleCellFormat)
    worksheet.write(7,3,'Loss Type',titleCellFormat)
    worksheet.write(7,4,'Loss Reason',titleCellFormat)
    worksheet.write(7,5,'SKU',titleCellFormat)
    worksheet.write(7,6,'Stop Date/Time',titleCellFormat)
    worksheet.write(7,7,'Start Date/Time',titleCellFormat)
    worksheet.write(7,8,'Duration',titleCellFormat)
    worksheet.set_row(7,40)
    summaryFormat=workbook.add_format({
        'text_wrap':True,
        'align':'right',
        'bg_color':'#EFFEFF',
        'border':1,
        'border_color':'black',
        'bold':True
    })
    summaryDataFormat=workbook.add_format({
        'text_wrap':True,
        'align':'center',
        'bg_color':'#EFFEFF',
        'border':1,
        'border_color':'black',
        'bold': True
    })
    idx=7
    shift=""
    shiftsize= timedelta(hours=8, minutes=0, seconds=0)
    totalDuration=timedelta(days=0,hours=0,minutes=0,seconds=0)
    mngmtLossDuration=timedelta(days=0,hours=0,minutes=0,seconds=0)
    downReason=[]
    startDate=spt
    startDate-=timedelta(hours=spt.hour-6,minutes=spt.minute,seconds=spt.second)
    print(startDate)
    # TotalTime= stt-spt
    for row in result:
        if (st=="var"):
            tup=row[1]-startDate
            counter=int(tup/shiftsize)
            if (shiftcount==-1):
                shiftcount=counter
            elif (shiftcount==counter):
                shiftcount=counter
            elif (shiftcount!=counter):
                idx+=1
                worksheet.set_row(idx,60)
                worksheet.merge_range(idx,1,idx,7,f"Shift {shift} Total D/T (HH:MM:SS): \n Shift {shift} Mgmt / Legal / Shortages / CO / PM Losses D/T (HH:MM:SS): \n Shift {shift} D/T %: \n Shift {shift} OLE impact %:",summaryFormat)
                totalLossper=round(totalDuration/shiftsize*100,2)
                totaloleLossPer=mngmtLossDuration/shiftsize
                worksheet.write(idx,8,f"{totalDuration} \n {mngmtLossDuration} \n {totalLossper} \n {totaloleLossPer}",summaryDataFormat)
                totalDuration=timedelta(days=0,hours=0,minutes=0,seconds=0)
                mngmtLossDuration=timedelta(days=0,hours=0,minutes=0,seconds=0)
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

        if (row[3]=='mac'):
            query=f"Select {machine}_floattable.DateAndTime, {machine}_floattable.val, {machine}_floattable.TagIndex, `fault reason mapping new`.Description ,`fault reason mapping new`.`fault type` from {machine}_floattable left join `fault reason mapping new` on {machine}_floattable.val=`fault reason mapping new`.`fault code` where DateAndTime >= '{row[1]}' and TagIndex = {AlarmTag} and val <> 0"
            cursor.execute(query)
            Description=cursor.fetchone()
            cursor.execute("select oee_stringtable.val from oee_stringtable where tagIndex=9 and DateAndTime <= (%s) order by DateAndTime desc limit 1",(Description[0],))
            sku=cursor.fetchone()
            if (Description[0]-row[1]>timedelta(hours=0, minutes=10, seconds=0)):
                continue
            downReason.append((Description[3],row[2]))
            worksheet.write(idx+1,1,Description[0].strftime('%d-%b-%Y'),border)
            worksheet.write(idx+1,2,shift,border)
            worksheet.write(idx+1,3,Description[4],border)
            worksheet.write(idx+1,4,Description[3],border)
            worksheet.write(idx+1,5,sku[0],border)
            worksheet.write(idx+1,6,row[1].strftime('%d-%b-%Y %X'),border)
            worksheet.write(idx+1,7,row[0].strftime('%d-%b-%Y %X'),border)
            worksheet.write(idx+1,8,strfdelta(row[2],'%H:%M:%S'),border)
            totalDuration+=row[2]
            idx+=1
        elif (row[3]=='mgmt'):
            cursor.execute("select oee_stringtable.val from oee_stringtable where tagIndex=9 and DateAndTime <= (%s) order by DateAndTime desc limit 1",(row[0],))
            sku=cursor.fetchone()
            worksheet.write(idx+1,1,row[0].strftime('%d-%b-%Y'),border)
            worksheet.write(idx+1,2,shift,border)
            worksheet.write(idx+1,3,'mgmt',border)
            worksheet.write(idx+1,4,'mgmt',mgmtlosss)
            worksheet.write(idx+1,5,sku[0],border)
            worksheet.write(idx+1,6,row[0].strftime('%d-%b-%Y %X'),border)
            worksheet.write(idx+1,7,row[1].strftime('%d-%b-%Y %X'),border)
            worksheet.write(idx+1,8,strfdelta(row[2],'%H:%M:%S'),mgmtlosss)
            mngmtLossDuration+=row[2]
            idx+=1
    if (idx!=7):
        idx+=1
        worksheet.set_row(idx,60)
        worksheet.merge_range(idx,1,idx,7,f"Shift {shift} Total D/T (HH:MM:SS): \n Shift {shift} Mgmt / Legal / Shortages / CO / PM Losses D/T (HH:MM:SS): \n Shift {shift} D/T %: \n Shift {shift} OLE impact %:",summaryFormat)
        totalLossper=round(totalDuration/shiftsize*100,2)
        totaloleLossPer=mngmtLossDuration/shiftsize
        worksheet.write(idx,8,f"{totalDuration} \n {mngmtLossDuration} \n {totalLossper} \n {totaloleLossPer}",summaryDataFormat)
        totalDuration=timedelta(days=0,hours=0,minutes=0,seconds=0)
        mngmtLossDuration=timedelta(days=0,hours=0,minutes=0,seconds=0)
    workbook.close()
    path=downloadDirectory+f"/{ReportName}"
    return send_file(path)

def summary(machine,ReportType,Shift,Date,ToDate,Week,Year,Month,statusTag,alarmTag,Line):
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
        return ReportGenerator(machine,spt,stt,ReportName,statusTag,alarmTag,shift,ReportType,Line)
    elif(ReportType=="Daily"):
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( hours=30,minutes=0, seconds=0)
        spt= spt.replace(tzinfo=None)
        ReportName=f"Report_{machine}_{Date}.xlsx"
        return ReportGenerator(machine, spt,stt,ReportName,statusTag,alarmTag,"var",ReportType,Line)
    elif (ReportType=="Weekly"):
        Date= datetime.datetime(Year,1,1,0,0,0)
        Date+= timedelta(days=6-Date.weekday())
        Date+= timedelta(days=(Week-1)*7)
        spt=Date+timedelta(hours=6, minutes=0, seconds=0)
        stt=Date+timedelta( days=7,hours=6,minutes=0, seconds=0)
        ReportName=f"Report_{machine}_{Year}_Week{Week}.xlsx"
        return ReportGenerator(machine,spt,stt,ReportName,statusTag,alarmTag,"var",ReportType,Line)
    elif (ReportType=='Monthly'):
        spt= datetime.datetime(Year,Month,1,6,0,0)
        if (Month==12):
            stt= datetime.datetime(Year+1,1,1,6,0,0)
        else:
            stt= datetime.datetime(Year,Month+1,1,6,0,0)
        ReportName=f"Report_{Month}{Year}.xlsx"
        return ReportGenerator(machine,spt,stt,ReportName,statusTag,alarmTag,"var",ReportType,Line)
    elif(ReportType=='Custom'):
        spt=Date
        stt=ToDate
        spt= spt.replace(tzinfo=None)
        ReportName=f"Report_{machine}_{Date}_{ToDate}.xlsx"
        return ReportGenerator(machine, spt,stt,ReportName,statusTag,alarmTag,"var",ReportType,Line)
    return ""

def OleReport(Month,Year):
    workbook= xlsxwriter.Workbook("Reports/Ole.xlsx")
    worksheet=workbook.add_worksheet("shift")
    worksheet.write(9,0,'Paticulars (SKU-1 gm)')
    worksheet.write(10,0,'varient 1')
    worksheet.write(11,0,'machine design speet for machine 5 & machine 6')
    worksheet.write(12,0,'machine design speet for machine 7 & machine 8')
    worksheet.write(13,0,'Average Design Speed')
    worksheet.write(14,0,'Shift time (SKU wise)')
    worksheet.write(16,0,'Downtime analysis (excluded from OEE calcutlation) ')
    worksheet.write(17,0,'Legal --> National/gazatted holiday')
    worksheet.write(18,0,'Managment Loss --> No plan/Sanitasation/Startu up & Shut Down Loss/Power faliure')
    worksheet.write(19,0,'Man Power Shortage --> Man Power Shortage')
    worksheet.write(20,0,'RM/PM/SSFA not available --> RM/PM Shortage')
    worksheet.write(21,0,'Changeover-->SKU Changeover')
    worksheet.write(22,0,'Changeover-->Color Changeover')
    worksheet.write(23,0,'Preventive Maintance --> Planned Condition maintance, CLIT, JH')
    worksheet.write(24,0,'Total machine stoppage due to managment loss (min)')
    worksheet.write(25,0,'Total machine stoppage includeing changeover, manpower & RM/PM shortage')
    worksheet.write(26,0,'Equipment Breakdown Details (Before Bar Cutter)- To be included in OEE Calulation')
    worksheet.write(27,0,'Breakdown Time of Mixer (in min)')
    worksheet.write(29,0,'Equipment Breakdown Details (From Bar Cutter) >10 minutes- To be included in OEE')
    worksheet.write(30,0,'Breakdown Time of Bar-Cutter-III(in min)')
    worksheet.write(31,0,'Breakdown Time of Bar-Cutter-IV(in min)')
    worksheet.write(33,0,'Breakdown Time of Stamping Machine-III(in min)')
    worksheet.write(34,0,'Breakdown Time of Stamping Machine-IV(in min)')
    worksheet.write(36,0,'Breakdown Time of Wrapping Machine-V(in min)')
    worksheet.write(37,0,'Breakdown Time of Wrapping Machine-VI(in min)')
    worksheet.write(38,0,'Breakdown Time of Wrapping Machine-VII(in min)')
    worksheet.write(39,0,'Breakdown Time of Wrapping Machine-VIII(in min)')
    worksheet.write(41,0,'Breakdown Time of Banding Machine-I(in min)')
    worksheet.write(42,0,'Breakdown Time of Banding Machine-II(in min)')
    worksheet.write(44,0,'Total Breakdown (min)')
    worksheet.write(46,0,'Equipment Breakdown Details (From Bar Cutter) <=10 minutes')
    worksheet.write(47,0,'Minor Stoppage Breakdown Time of Stamping Machine-III(in min)')
    worksheet.write(48,0,'Minor Stoppage Breakdown Time of Stamping Machine-IV(in min)')
    worksheet.write(49,0,'Minor Stoppage Breakdown Time of Wrapping Machine-V(in min)')
    worksheet.write(50,0,'Minor Stoppage Breakdown Time of Wrapping Machine-VI(in min)')
    worksheet.write(51,0,'Minor Stoppage Breakdown Time of Wrapping Machine-VII(in min)')
    worksheet.write(52,0,'Minor Stoppage Breakdown Time of Wrapping Machine-VIII(in min)')
    worksheet.write(54,0,'Total Minor Stoppage (min)')
    worksheet.write(56,0,'Total Run Time of Wrapping Machine-V (in min)')
    worksheet.write(57,0,'Total Run Time of Wrapping Machine-VI (in min)')
    worksheet.write(58,0,'Total Run Time of Wrapping Machine-VII (in min)')
    worksheet.write(59,0,'Total Run Time of Wrapping Machine-VIII (in min)')
    worksheet.write(60,0,'Average run time of wrapping machine (4 nos.)')
    worksheet.write(62,0,'Run time of wrapping machines (includein mgmt losses- changeover, RM PM shortage, manpower)')
    worksheet.write(63,0,'Total Run Time of Wrapping Machine-V (in min)')
    worksheet.write(64,0,'Total Run Time of Wrapping Machine-VI (in min)')
    worksheet.write(65,0,'Total Run Time of Wrapping Machine-VII (in min)')
    worksheet.write(66,0,'Total Run Time of Wrapping Machine-VIII (in min)')
    worksheet.write(67,0,'Average run time of wrapping machine (4 nos.)')
    worksheet.write(69,0,'Production Data')
    worksheet.write(70,0,'Bar Output from Cutter-III')
    worksheet.write(71,0,'Bar Output from Cutter-IV')
    worksheet.write(72,0,'Output Cakes at Wrapping Machine-V (in Nos)')
    worksheet.write(73,0,'Output Cakes at Wrapping Machine-VI (in Nos)')
    worksheet.write(74,0,'Output Cakes at Wrapping Machine-VII (in Nos)')
    worksheet.write(75,0,'Output Cakes at Wrapping Machine-VIII (in Nos)')
    worksheet.write(76,0,'Total Output Cakes of Wrapping Machine (4 Nos)')
    worksheet.write(78,0,'Final Production Output Wrapping Machine-V in MT')
    worksheet.write(79,0,'Final Production Output Wrapping Machine-VI in MT')
    worksheet.write(80,0,'Final Production Output Wrapping Machine-VII in MT')
    worksheet.write(81,0,'Final Production Output Wrapping Machine-VIII in MT')
    worksheet.write(83,0,'TOTAL WRAPPING PRODUCTION (MT)')
    worksheet.write(84,0,'Total Quantity Recycled at Wrapping Machine-V in MT')
    worksheet.write(85,0,'Total Quantity Recycled at Wrapping Machine-VI in MT')
    worksheet.write(86,0,'Total Quantity Recycled at Wrapping Machine-VII in MT')
    worksheet.write(87,0,'Total Quantity Recycled at Wrapping Machine-VIII in MT')
    worksheet.write(88,0,'Total Quantity Recycled at Wrapping Machines')
    worksheet.write(90,0,'Total Quantity Recycled at Banding Machine-I in MT')
    worksheet.write(91,0,'Total Quantity Recycled at Banding Machine-II in MT')
    worksheet.write(92,0,'Total Quantity Recycled at Banding in MT')
    worksheet.write(94,0,'Total Line Production (MT)')
    worksheet.write(96,0,'Check Production in MT')
    MonthStart=datetime.datetime(Year,Month,1,6,0,0)
    MonthEnd=datetime.datetime(Year,Month,6,0,0)
    if (Month==12):
        MonthEnd=datetime.datetime(Year+1,Month,1,6,0,0)
    else:
        MonthEnd=datetime.datetime(Year,Month+1,1,6,0,0)
    NoOfDays=MonthEnd-MonthStart
    idx=1
    for day in range(NoOfDays.days):
        for shift in range(3):
            if (shift==0):
                Shift='A'
                StartTime=datetime.datetime(Year,Month,1+day,6,0,0)
                EndTime=datetime.datetime(Year,Month,1+day,14,0,0)
            elif(shift==1):
                Shift='B'
                StartTime=datetime.datetime(Year,Month,1+day,14,0,0)
                EndTime=datetime.datetime(Year,Month,1+day,22,0,0)
            else:
                Shift='C'
                StartTime=datetime.datetime(Year,Month,1+day,22,0,0)
                if (day!=NoOfDays.days-1):
                    EndTime=datetime.datetime(Year,Month,2+day,6,0,0)
                else:
                    EndTime=datetime.datetime(Year,Month+1,1,6,0,0)
            Date=MonthStart+timedelta(days=day)
            worksheet.write(7,idx,Date.strftime('%d-%b'))
            worksheet.write(8,idx,Shift)
            idx+=1
    workbook.close()
    path='Reports/Ole.xlsx'
    return send_file(path)
            

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
    ToDate=input_json['ToDate']
    Date=dateutil.parser.isoparse(Date)
    Date+= timedelta(hours=5,minutes=30,seconds=0)
    ToDate=dateutil.parser.isoparse(ToDate)
    ToDate+= timedelta(hours=5,minutes=30,seconds=0)
    Year=dateutil.parser.isoparse(Year)
    Year+= timedelta(hours=5,minutes=30,seconds=0)
    Year=Year.year
    Month=dateutil.parser.isoparse(Month)
    Month+= timedelta(hours=5,minutes=30,seconds=0)
    Month=Month.month
    print(Month)
    Week=int(Week)
    if (Report=='OLEReport'):
        return OleReport(Month,Year)
    if (Machine=='Banding1'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,4,11,Line)
    elif (Machine=='Banding2'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,4,11,Line)
    elif (Machine=='Cutter3'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,3,10,Line)
    elif (Machine=='Cutter4'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,4,11,Line)
    elif (Machine=='Stamper3'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,4,11,Line)
    elif (Machine=='Stamper4'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,4,11,Line)
    elif (Machine=='Wrapper5'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,4,11,Line)
    elif (Machine=='Wrapper6'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,4,11,Line)
    elif (Machine=='Wrapper7'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,4,11,Line)
    elif (Machine=='Wrapper8'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,4,11,Line)
    elif (Machine=='Mpc'):
        return summary(Machine.lower(),ReportType,Shift,Date,ToDate,Week,Year,Month,1,8,Line)
    # return jsonify(dictToReturn)
    
    # except FileNotFoundError:
    #     abort(404)
if __name__ == "__main__":
    app.run(debug=True)