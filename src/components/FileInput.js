import React, { useState, Component } from 'react';
import parse from 'html-react-parser';

function FileInput() {
  const [selectedFile, setSelectedFile] = useState(null);

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    setSelectedFile(file);
  };

  function extractDataFromHTML(html) {
    // const dom = parse(html);
    // const paragraphs = dom.querySelectorAll('p');
    // const extractedData = paragraphs.map((p) => p.textContent.trim());
    // return extractedData;

    // const dom = parse(html);
    // const parser = new DOMParser();
    // const doc = parser.parseFromString(dom, 'text/html');
    // const rows = doc.querySelectorAll('tr');

    var dom = new DOMParser();
    dom = parse(html);
    const paragraphs = dom.querySelectorAll('p');
    const extractedData = paragraphs.map((p) => p.textContent.trim());
    // return extractedData;

    console.log(extractedData);
  }

  const handleButtonClick = () => {
    if (selectedFile) {
      const reader = new FileReader();

      console.log(selectedFile);

      reader.onload = (event) => {
        const fileContent = event.target.result;
        // const extractedData = extractDataFromHTML(fileContent);
        // TODO: Save extractedData or perform Excel-related operations
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

export default FileInput;
