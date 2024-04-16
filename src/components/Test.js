import React from 'react';
import cheerio from 'cheerio';
import XLSX from 'xlsx';

class Test extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      data: [],
    };
  }

  handleFileChange = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    
    reader.onload = (e) => {
      const html = e.target.result;
      const $ = cheerio.load(html); // Load the HTML into Cheerio

      // Manipulate the HTML using Cheerio selectors
      const extractedData = [];
      $('table tr').each((index, element) => {
        const row = [];
        $(element)
          .find('td')
          .each((i, el) => {
            row.push($(el).text().trim());
          });
        extractedData.push(row);
      });

      this.setState({ data: extractedData });
    };

    reader.readAsText(file);
  };

  exportToExcel = () => {
    const { data } = this.state;

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet 1');

    const excelBuffer = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });

    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    const filename = 'data.xlsx';

    if (typeof window.navigator.msSaveBlob !== 'undefined') {
      // IE workaround
      window.navigator.msSaveBlob(blob, filename);
    } else {
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  render() {
    return (
      <div>
        <input type="file" onChange={this.handleFileChange} />
        <button onClick={this.exportToExcel}>Export to Excel</button>
      </div>
    );
  }
}

export default Test;
