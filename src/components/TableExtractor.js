import React, { Component, useEffect, useState } from 'react'

const TableExtractor = () => {
    const [tableData, setTableData] = useState([]);
    const [selectedFile, setSelectedFile] = useState(null);

    const handleFileChange = (event) => {
        const file = event.target.files[0];
        setSelectedFile(file);
    };

    function extractDataFromHTML(html) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        const table = doc.querySelector('table');
        const tableData = [];

      console.log(doc.table);

        if (table) {
            const rows = table.querySelectorAll('tr');



            rows.forEach((row) => {
                const rowData = [];
                const cells = row.querySelectorAll('td');

                console.log(cells);
                
                cells.forEach((cell) => {
                  rowData.push(cell.innerText);
                });
                
                tableData.push(rowData);
            });
        }

        setTableData(tableData);

    }

    const handleButtonClick = () => {
        if (selectedFile) {
          const reader = new FileReader();
    
          reader.onload = (event) => {
            const fileContent = event.target.result;

            // console.log(fileContent);

            // const extractedData = extractDataFromHTML(fileContent);
            // TODO: Save extractedData or perform Excel-related operations

            extractDataFromHTML(fileContent);
          };
    
          reader.readAsText(selectedFile);
        }
      };

    return (
        <div>
            <input type="file" onChange={handleFileChange} />
            <button onClick={handleButtonClick}>Extract Data</button>
        </div>
      );
}

export default TableExtractor;