import React, { useState } from "react";
import Container from 'react-bootstrap/Container';
import Navbar from 'react-bootstrap/Navbar';

const Infobar=()=>{
    const [date,setDate]=useState(new Date().toLocaleDateString())
    return (
        <Navbar>
            <Container>
                <Navbar.Text className="report-navbar" >REPORT CRITERIA</Navbar.Text>
                <Navbar.Text className="date-navbar">
                    {date}
                </Navbar.Text>
            </Container>
        </Navbar>
    );
}
export default Infobar