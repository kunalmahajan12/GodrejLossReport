import axios from "axios";
import React, {useState} from "react";
import FileDownload from 'js-file-download'
import DatePicker from 'react-date-picker';
import Card from 'react-bootstrap/Card';


const ReportGenerator=()=>{
    const [criteria,setCriteria]=useState({
        Line:'Line1',
        Report: 'DowntimeReport',
        ReportType: 'Shift',
        ReportTypeIsHidden:false,
        Machine: 'Cutter3',
        MachineIsHidden: false,
        Date: new Date(),
        DateIsHidden:false,
        Year: new Date(),
        YearIsHidden: true,
        Month: new Date(),
        MonthIsHidden: true,
        Shift: 'ShiftA',
        ShiftIsHidden: false,
        Week: "1",
        WeekIsHidden: true,
    })
    
    
    const handleLineChange=(event)=> {
        setCriteria(prevState=>({
            ...prevState,
            Line: event.target.value}));
    }
    const handleReportChange=(event) =>{
        setCriteria(prevState=>({
            ...prevState,
            Report: event.target.value
        }));
        if (event.target.value==='OLEReport'){
            setCriteria(prevState=>({
                ...prevState,
                Month: new Date(),
                Year: new Date(),
                MachineIsHidden: true,
                DateIsHidden: true,
                WeekIsHidden:true,
                ShiftIsHidden:true,
                MonthIsHidden: false,
                YearIsHidden: false,
                ReportTypeIsHidden: true,
            }))
        }
        else{
            setCriteria(prevState=>({
                ...prevState,
                MachineIsHidden: false,
                ReportTypeIsHidden: false,
                ReportType:'Shift',
                Shift: 'ShiftA',
                DateIsHidden:false,
                ShiftIsHidden: false,
                YearIsHidden: true,
                MonthIsHidden: true,
                WeekIsHidden: true,
            }))
        }

    }
    const handleReportTypeChange=(event) =>{
        setCriteria(prevState=>({
            ...prevState,
            ReportType: event.target.value
        }));
        if(event.target.value==='Shift'){
            setCriteria(prevState=>({
                ...prevState,
                Date: new Date(),
                Shift: 'ShiftA',
                DateIsHidden:false,
                ShiftIsHidden: false,
                YearIsHidden: true,
                MonthIsHidden: true,
                WeekIsHidden: true,
            }))
        }
        else if(event.target.value==='Daily'){
            setCriteria(prevState=>({
                ...prevState,
                Date: new Date(),
                DateIsHidden: false,
                YearIsHidden: true,
                MonthIsHidden: true,
                WeekIsHidden: true,
                ShiftIsHidden: true
            }))
        }
        else if (event.target.value==='Weekly'){
            setCriteria(prevState=>({
                ...prevState,
                Year: new Date(),
                Week: 1,
                DateIsHidden: true,
                YearIsHidden: false,
                MonthIsHidden: true,
                WeekIsHidden: false,
                ShiftIsHidden: true
            }))
        }
        else if (event.target.value==='Monthly'){
            setCriteria(prevState=>({
                ...prevState,
                Month: new Date(),
                Year: new Date(),
                DateIsHidden: true,
                YearIsHidden: false,
                MonthIsHidden: false,
                WeekIsHidden: true,
                ShiftIsHidden: true
            }))
        }
    }
    const handleMachineChange=(event)=> {
        setCriteria(prevState=>({
            ...prevState,
            Machine: event.target.value
        }))
    }
    const handleDateChange=(date) =>{
        setCriteria(prevState =>({
            ...prevState,
            Date: date
        }))
    }
    const handleYearChange=(Date) =>{
        setCriteria(prevState=>({
            ...prevState,
            Year: Date
        }))
    }
    const handleMonthChange=(Date)=> {
        setCriteria(prevState=>({
            ...prevState,
            Month: Date
        }))
    }
    const handleWeekChange=(event)=> {
        setCriteria(prevState=>({
            ...prevState,
            Week: event.target.value
        }))
    }
    const handleShiftChange=(event) =>{
        setCriteria(prevState=>({
            ...prevState,
            Shift: event.target.value
        }))
    }
    const handleSubmit=(event)=> {
        console.log(criteria);
        // useEffect(() => {
        //     const url = "http://127.0.0.1:5000/";
        
        //     const fetchData = async () => {
        //       try {
        //         const response = await fetch(url);
        //         const json = await response.json();
        //         console.log(json);
        //       } catch (error) {
        //         console.log("error", error);
        //       }
        //     };
        
        //     fetchData();
        // }, []);
        axios({
            url:'http://127.0.0.1:5000/lossReport',
            method:'POST',
            data: criteria,
            responseType: 'blob'
        }).then((res)=>{
            FileDownload(res.data,"report.xlsx");
       //Perform action based on response
        })
        .catch(function(error){
            console.log(error);
       //Perform action based on error
        });
        event.preventDefault();
    }
    return(
        <form onSubmit={handleSubmit}>
            <Card>
                <Card.Header>
                    <Card.Title>Line Selection</Card.Title>
                </Card.Header>
                <Card.Body>
                    
                    <Card.Subtitle>Line:</Card.Subtitle>
                    <select value={criteria.Line} onChange={handleLineChange} className="w">
                    <option value="Line1">Line 1</option>
                    <option value="Line2">Line2</option>
                    <option value="Line3">Line3</option>
                </select>
                </Card.Body>
            </Card>
            <Card>
                <Card.Header>
                    <Card.Title>Report Selection</Card.Title>
                </Card.Header>
                <Card.Body>
                    
                    <Card.Subtitle>Report</Card.Subtitle>
                    <select value={criteria.Report} onChange={handleReportChange}className="w">
                        <option value="DowntimeReport">Downtime Report</option>
                        <option value="MinorStoppageReport">Minor Stoppage Report</option>
                        <option value="BreakdownReport">Breakdown Report</option>
                        <option value="OLEReport">OLE Report</option>
                    </select>
                    {!criteria.ReportTypeIsHidden&& 
                    <div>
                        <Card.Subtitle>Report Type</Card.Subtitle>
                        <select value={criteria.ReportType} onChange={handleReportTypeChange}className="w">
                            <option value="Shift">Shift Wise</option>
                            <option value="Daily">Daily</option>
                            <option value="Weekly">Weekly</option>
                            <option value="Monthly">Monthly</option>
                        </select>
                    </div>}
                </Card.Body>
            </Card>
            <Card>
                <Card.Header>
                    <Card.Title>Machine Selection</Card.Title>
                </Card.Header>
                <Card.Body>
                    
                    {!criteria.MachineIsHidden&&
                    <div>
                        <Card.Subtitle>Machine</Card.Subtitle>
                        <select value={criteria.Machine} onChange={handleMachineChange}className="w">
                            <option value="Cutter3">Cutter3</option>
                            <option value="Cutter4">Cutter4</option>
                            <option value="Stamper3">Stamper3</option>
                            <option value="Stamper4">Stamper4</option>
                            <option value="Wrapper5">Wrapper5</option>
                            <option value="Wrapper6">Wrapper6</option>
                            <option value="Wrapper7">Wrapper7</option>
                            <option value="Wrapper8">Wrapper8</option>
                            <option value="Mpc">MPC</option>
                            <option value="Banding1">Banding1</option>
                            <option value="Banding2">Banding2</option>
                        </select>
                    </div>}
                </Card.Body>
            </Card>
            <Card>
                <Card.Header>
                    <Card.Title>Date Selection</Card.Title>
                </Card.Header>
                <Card.Body>
                    
                    {!criteria.DateIsHidden&&
                    <div>
                        <Card.Subtitle>Date</Card.Subtitle>
                        <DatePicker value={criteria.Date} onChange={handleDateChange} className="w" clearIcon={null}/>
                    </div>}
                    {!criteria.ShiftIsHidden&&
                    <div>
                        <Card.Subtitle>Shift</Card.Subtitle>
                        <select value={criteria.Shift} onChange={handleShiftChange} className="w">
                            <option value="ShiftA">Shift A</option>
                            <option value="ShiftB">Shift B</option>
                            <option value="ShiftC">Shift C</option>
                        </select>
                    </div>}
                    {!criteria.YearIsHidden&&
                    <div>
                        <Card.Subtitle>Year</Card.Subtitle>
                        <DatePicker onChange={handleYearChange} value={criteria.Year} format="y" disabled={criteria.YearIsHidden} className="w"/>
                    </div>}
                    {!criteria.MonthIsHidden&&
                    <div>
                        <Card.Subtitle>Month</Card.Subtitle>
                        <DatePicker onChange={handleMonthChange} value={criteria.Month} format="MMM" className="w"/>
                    </div>}
                    {!criteria.WeekIsHidden&&
                    <div>
                        <Card.Subtitle>Week</Card.Subtitle>
                        <input type="text" name="week" value={criteria.Week} onChange={handleWeekChange} className="w"/>
                    </div>}
                </Card.Body>
            </Card>
            
            <input type="submit" value="Show Report" className="btn" />
        </form>
    )
        
}
export default ReportGenerator;