import './App.css';
import Infobar from './components/infobar';
import ReportGenerator from './components/reportGenerator';
function App() {
  return (
    <div className="App">
      <Infobar></Infobar>
      <h1>LOSS CAPTURING SYSTEM</h1>
      <ReportGenerator></ReportGenerator>
    </div>
  );
}

export default App;
