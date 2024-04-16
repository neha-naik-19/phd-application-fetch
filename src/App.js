import logo from './logo.svg';
import './App.css';
import FileInput from "./components/FileInput";
import TableExtractor from "./components/TableExtractor";
import TableParser from "./components/TableParser";
import TestMultipleFile from "./components/TestMultipleFile";

function App() {
  return (
    <TableParser />

    // <TestMultipleFile />

    // <div className="App">
    //   <header className="App-header">
    //     <img src={logo} className="App-logo" alt="logo" />
    //     <p>
    //       Edit <code>src/App.js</code> and save to reload.
    //     </p>
    //     <a
    //       className="App-link"
    //       href="https://reactjs.org"
    //       target="_blank"
    //       rel="noopener noreferrer"
    //     >
    //       Learn React
    //     </a>
    //   </header>
    // </div>
  );
}

export default App;
