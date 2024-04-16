import React, {  useState } from 'react';
import XLSXC from "xlsx-color";
import parse from "html-react-parser";
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

// const EXCEL_TYPE ="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";
// const EXCEL_EXTENSION = ".xlsx";

const TableParser = () => {
    const [selectedFile, setSelectedFile] = useState(null);
    const [uploadedFiles, setUploadedFiles] = useState([])
    const [fileLimit, setFileLimit] = useState(false);

    var [CstwelthPerc, setCstwelthPerc] = useState('');
    var [selectCsTwelthPerc, setSelectCsTwelthPerc] = useState('0');

    var [etctwelthPerc, setEtctwelthPerc] = useState('');
    var [selectEtcTwelthPerc, setSelectEtcTwelthPerc] = useState('0');

    var [othertwelthPerc, setOthertwelthPerc] = useState('');
    var [selectOtherTwelthPerc, setSelectOtherTwelthPerc] = useState('0');

    /** CS **/
    const handleCsTwelthPercFunc = event => {
      setCstwelthPerc(event.target.value);
      // console.log('value is:', event.target.value);
    };

    const selectCsTwelthPercFunc = event => {
      setSelectCsTwelthPerc(event.target.value);
      // console.log('select value is:', selectCsTwelthPerc);
    };

    /** ETC **/
    const handleEtcTwelthPercFunc = event => {
      setEtctwelthPerc(event.target.value);
    };

    const selectEtcTwelthPercFunc = event => {
      setSelectEtcTwelthPerc(event.target.value);
    };

    /** OTHERS **/
    const handleOtherTwelthPercFunc = event => {
      setOthertwelthPerc(event.target.value);
    };

    const selectOtherTwelthPercFunc = event => {
      setSelectOtherTwelthPerc(event.target.value);
    };

    const handleFileChange = (event) => {
        const file = event.target.files[0];
        setSelectedFile(file);
    };

    const handleUploadFiles = files => {
      const uploaded = [...uploadedFiles];
      let limitExceeded = false;
      files.some((file) => {
          if (uploaded.findIndex((f) => f.name === file.name) === -1) {
              uploaded.push(file);
          }
      })
      if (!limitExceeded) setUploadedFiles(uploaded)
    }

    const handleFileEvent =  (e) => {
      const chosenFiles = Array.prototype.slice.call(e.target.files)
      handleUploadFiles(chosenFiles);
      // setSelectedFile(chosenFiles);
    }

    let slno = 0;
    let programmeType = ''
    const parseTablesFromHTML = (htmlData) => {
      const parser = new DOMParser().parseFromString(htmlData, 'text/html').querySelector('center');
      const allTables = parser.querySelectorAll('table')

      let objectItems = Object.entries(allTables).map((i) => {
        return {
          id : i[1].id,
          innerHTML: i[1].innerHTML
        };
      });

      const filteredObjectItems = Object.fromEntries(
        Object.entries(objectItems).filter(([key, value]) => value.id === 'Table9' || value.id === 'Table11' 
        || value.id === 'Table12' || value.id === 'Table14' || value.id === 'Table15' || value.id === 'Table27')
      );

      const tables = [];
      const tableData = [];

      Object.keys(filteredObjectItems).forEach((objKeys) => {
        var obj = {
          id : filteredObjectItems[objKeys]['id'],
          innerHTML : filteredObjectItems[objKeys]['innerHTML']
        }

        tables.push(obj);
      })

      let obj = [];
      let appno ='', department = '', name = '', gender = '', decodedEmail = '', mobileNo = '',
      campusPreference = '', isBitsStudent = '', currentWorkPlace = '', qualifyExam = '', nameOfDegreeFDHD = '', subjectFDHD = '',
      cgpaFDHD = '', YearOfPassingFDHD = '', twelthcCgpa = '', ugDegree1 = '', ugCgpa1 = '', ugSubject1 = '', ugYearOfPassing1 = '',
      ugDegree2 = '', ugCgpa2 = '', ugSubject2 = '', ugYearOfPassing2 = '', pgDegree1 = '', pgDegree2 = '', pgCgpa1 = '', pgCgpa2 = '', 
      pgSubject1 = '', pgSubject2 = '', pgYearOfPassing1 = '', pgYearOfPassing2 = '', fdOrHd = '', anyOtherExam = '', broadArea = '',
      universityFDHD = '', uguniversity1 = '', uguniversity2 = '', pguniversity1 = '', pguniversity2 = '';


      for (const tableElement of tables) {
        // Parsing the table element using html-react-parser
        const parsedElement = parse(tableElement.innerHTML);

        //Extract Table9 ==> Department, App No. , Programme Type
        if(tableElement.id === 'Table9'){
          // department = parsedElement[1].props.children[2].props.children[0].props.children[1].props.children.slice(0,4);
          programmeType = parsedElement[1].props.children[2].props.children[0].props.children[1].props.children.toString().split('/')[1].toUpperCase();
          appno = parsedElement[1].props.children[2].props.children[0].props.children[8].props.children;

        /* Extract Table12 ==> Name, Gender, Mobile Number, Email-Id */
        }else if(tableElement.id === 'Table12'){
          name = parsedElement[1].props.children[0].props.children[2].props.children[1].props.children;
          gender = parsedElement[1].props.children[1].props.children[2].props.children[1].props.children.trim();
          decodedEmail = decodeCfemail(parsedElement[1].props.children[6].props.children[2].props.children[1].props.children.props['data-cfemail']);
          mobileNo = parsedElement[1].props.children[11].props.children[2].props.children[1].props.children;

        /* Extract Table14 ==> Campus Preference, Higher Degree from BITS programme, Current Work Place */
        }else if(tableElement.id === 'Table14'){
          // Campus Preference
          let childNode = parsedElement[1].props.children[0].props.children[2].props.children[1].props.children.props.children.props.children;
          for (let i = 0; i < childNode.length; i++) {
            if(childNode[i].props.children[1].props.children[1].props.children.toUpperCase() === "GOA"){
              campusPreference = childNode[i].props.children[1].props.children[1].props.children.toUpperCase() + (i + 1);
            }
          }

          // Higher Degree from BITS programme
          childNode = parsedElement[1].props.children[1].props.children[2].props.children[1].props.children;

          isBitsStudent = '';
          if(childNode.toLowerCase().trim() !== 'no'){
            isBitsStudent = 'YES';
          }

          // Current Work Place
          childNode = parsedElement[1].props.children[2].props.children[2].props.children[1];
          if(childNode.props.children[0].toLowerCase() === 'yes'){
            currentWorkPlace = childNode.props.children[2].split(':')[1].trim()
          }
        
        /* Extract Table11 ==> Qualify Exam */  
        }else if(tableElement.id === 'Table11'){
          fdOrHd =  parsedElement[1].props.children[3].props.children.props.children[1].props.children;
          if(fdOrHd.toUpperCase() === 'FD'){
            qualifyExam = 'First Degree'
          }else if(fdOrHd.toUpperCase() === 'HD'){
            qualifyExam = 'Higher Degree'
          }

        /* Extract Table11 ==> HD => Name of the Degree, Subject, CGPA, Year of Passing */  
        }else if(tableElement.id === 'Table15'){
          nameOfDegreeFDHD = parsedElement[1].props.children[0].props.children[2].props.children[1].props.children.trim();
          subjectFDHD = parsedElement[1].props.children[2].props.children[2].props.children[1].props.children.trim();
          cgpaFDHD = parsedElement[1].props.children[3].props.children[2].props.children[1].props.children.trim();
          YearOfPassingFDHD = parsedElement[1].props.children[4].props.children[2].props.children[1].props.children.trim();
          universityFDHD = parsedElement[1].props.children[1].props.children[2].props.children[1].props.children.trim();
      
        /* Extract Table27 ==> Academic Record */  
        }else if(tableElement.id === 'Table27'){
          //Academic Record
          let childNode = parsedElement[1].props.children[1].props.children.props.children[1].props.children;

          let tbodyData = childNode.props.children.slice(1);
          let internalObj = {};
          let tableArray = [];
          let cnt = 1;
          let lbldip = '', lblyop = '', lblduration = '', lbluniversity = '', lblmarks = '', lblmaxmarks = '', lbldivision = '', lblsubjects = '';
          
          for (let i = 0; i < tbodyData.length; i++) {
            let element = tbodyData[i].props.children;

            lbldip = ''; lblyop = ''; lblduration = ''; lbluniversity = ''; lblmarks = ''; lblmaxmarks = ''; lbldivision = ''; lblsubjects = '';

            for (let j = 0; j < element.length; j++) {
              if(element[j].props.children[1].props.children !== null){
                if (element[j].props.children[1].props.id === 'lbldip' + cnt){lbldip = element[j].props.children[1].props.children}
                if (element[j].props.children[1].props.id === 'lblyop' + cnt){lblyop = element[j].props.children[1].props.children}
                if (element[j].props.children[1].props.id === 'lblduration' + cnt){lblduration = element[j].props.children[1].props.children}
                if (element[j].props.children[1].props.id === 'lbluniversity' + cnt){lbluniversity = element[j].props.children[1].props.children}
                if (element[j].props.children[1].props.id === 'lblmarks' + cnt){lblmarks = element[j].props.children[1].props.children}
                if (element[j].props.children[1].props.id === 'lblmaxmarks' + cnt){lblmaxmarks = element[j].props.children[1].props.children}
                if (element[j].props.children[1].props.id === 'lbldivision' + cnt){lbldivision = element[j].props.children[1].props.children}
                if (element[j].props.children[1].props.id === 'lblsubjects' + cnt){lblsubjects = element[j].props.children[1].props.children}
    
                if(j === 7){
                  internalObj = {
                    'degrees': lbldip,
                    'yearofPassing':  lblyop,
                    'duration': lblduration,
                    'university' : lbluniversity,
                    'aggregateCGPA' : lblmarks,
                    'maximumCGPA' : lblmaxmarks,
                    'division' : lbldivision,
                    'subject' : lblsubjects
                  }

                  tableArray.push(internalObj);
                  cnt++;
                }
              }
            }
          }

          tableArray.forEach(element => { //JINAL MUKESHKUMAR SHAH, SAJAN BHAGEERATHAN
            let data = element.degrees.replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s/g, '').toLowerCase().trim();

            if(data.includes('mtech') || (data.includes('me') && data.length === 2) || (data.includes('ms')  && data.length === 2) || (data.includes('msc')  && data.length === 3) || data.includes('master') 
            || (data.includes('mca')  && data.length === 3) || (data.includes('mba')  && data.length === 3)){
              if(pgDegree1 === '' && pgCgpa1 === ''){
                pgDegree1 =  element.degrees.trim();
                pgCgpa1 = element.aggregateCGPA.trim();
                pgSubject1 = element.subject.trim();
                pgYearOfPassing1 = element.yearofPassing.trim()
              }
              else{
                pgDegree2 =  element.degrees.trim();
                pgCgpa2 = element.aggregateCGPA.trim();
                pgSubject2 = element.subject.trim();
                pgYearOfPassing2 = element.yearofPassing.trim()
              }
            }
            if(data.includes('btech') || (data.includes('be') && data.length === 2) ||  data.includes('bsc') || data.includes('bca') || data.includes('bachelor')){ 
              if(ugDegree1 === '' && ugCgpa1 === ''){
                ugDegree1 =  element.degrees.trim();
                ugCgpa1 = element.aggregateCGPA.trim();
                ugSubject1 = element.subject.trim();
                ugYearOfPassing1 = element.yearofPassing.trim()
              }
              else{
                ugDegree2 =  element.degrees.trim();
                ugCgpa2 = element.aggregateCGPA.trim();
                ugSubject2 = element.subject.trim();
                ugYearOfPassing2 = element.yearofPassing.trim()
              }
            }
            if(data.includes('highersecondary') || data.includes('puc') || data.includes('12th') || data.includes('hsc')  
            || data.includes('interboard') || data.includes('diploma') || data.includes('hs') || data.includes('xii') 
            || data.includes('intermediate') || data.includes('predegree') || data.includes('12') || data.includes('aissce')
            || data.includes('10+12') || data.includes('secondaryschool')){
              twelthcCgpa = element.aggregateCGPA;
            }

            let universityData = element.university.replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s/g, '').toLowerCase().trim()

            if(universityData.includes('bits') || universityData.includes('pilani')){
              if(isBitsStudent.toLowerCase().trim() !== "yes"){
                isBitsStudent = "YES";
              }
            }
          });

          if(universityFDHD.replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s/g, '').toLowerCase().trim().includes('bits') || universityFDHD.includes('pilani') ){
            if(isBitsStudent.toLowerCase().trim() !== "yes"){
              isBitsStudent = "YES";
            }
          }

          var testDegree = '';
          var testCgpa = '';
          var testSubject = '';
          var testYearOfPassing = '';

          //For FD degree
          if(fdOrHd.toUpperCase() === 'FD'){
            if(ugDegree1 === '' && ugDegree2 === ''){
              ugDegree1 = nameOfDegreeFDHD;
              ugCgpa1 = cgpaFDHD;
              ugSubject1 = subjectFDHD;
              ugYearOfPassing1 = YearOfPassingFDHD;
            }else{
              testDegree = ugDegree1;
              testCgpa = ugCgpa1;
              testSubject = ugSubject1;
              testYearOfPassing = ugYearOfPassing1;

              ugDegree1 = nameOfDegreeFDHD;
              ugCgpa1 = cgpaFDHD;
              ugSubject1 = subjectFDHD;
              ugYearOfPassing1 = YearOfPassingFDHD;
        
              if(ugDegree1 === ugDegree2){
                ugDegree2 = testDegree;
                ugCgpa2 = testCgpa;
                ugSubject2 = testSubject;
                ugYearOfPassing2 = testYearOfPassing;
              }
            }
          }

          //for HD degree
          testDegree = '';
          if(fdOrHd.toUpperCase() === 'HD'){
            if(pgDegree1 === '' && pgDegree2 === ''){
              pgDegree1 = nameOfDegreeFDHD;
              pgCgpa1 = cgpaFDHD;
              pgSubject1 = subjectFDHD;
              pgYearOfPassing1 = YearOfPassingFDHD;
            }else{
              testDegree = pgDegree1;
              testCgpa = pgCgpa1;
              testSubject = pgSubject1;
              testYearOfPassing = pgYearOfPassing1;

              pgDegree1 = nameOfDegreeFDHD;
              pgCgpa1 = cgpaFDHD;
              pgSubject1 = subjectFDHD;
              pgYearOfPassing1 = YearOfPassingFDHD;
        
              if(pgDegree1 === pgDegree2){
                pgDegree2 = testDegree;
                pgCgpa2 = testCgpa;
                pgSubject2 = testSubject;
                pgYearOfPassing2 = testYearOfPassing;
              }
            }
          }

          //Have cleared any of the following examination
          childNode = parsedElement[1].props.children[2].props.children.props.children[1].props.children;

          tbodyData = childNode.props.children.slice(1);
          tbodyData = childNode.props.children.slice(2);

          let node = tbodyData[0].props.children.props.children[1].props.children.props.children.props.children;

          let otherExamNameArray = [];
          for (let i = 0; i < node.length; i++) {
            otherExamNameArray.push(node[i].props.children[0].replace(/\s/g, '').replace(':','').trim());

            if(node[i].props.children[1].props.children !== null){
              if(node[i].props.children[1].props.children.trim().toLowerCase() === 'yes'){
                anyOtherExam = 'YES';
              }else{
                if(anyOtherExam === ''){
                  anyOtherExam = 'NO';
                }
              }
            }
          }

          department = tbodyData[1].props.children.props.children[1].props.children.props.children[0].props.children[1].props.children[1].props.children;
          broadArea = tbodyData[1].props.children.props.children[1].props.children.props.children[1].props.children[1].props.children[3].props.children.slice(0, 40);
        }
      }

      ugDegree1 = ugDegree1.length > 0 ? `${ugDegree1} (${ugSubject1})` : '';
      ugDegree2 = ugDegree2.length > 0 ? `${ugDegree2} (${ugSubject2})` : '';
      pgDegree1 = pgDegree1.length > 0 ? `${pgDegree1} (${pgSubject1})` : '';
      pgDegree2 = pgDegree2.length > 0 ? `${pgDegree2} (${pgSubject2})` : '';

      obj = {
        'Sl No.' : ++slno,
        'Qualify Exam' : qualifyExam,
        'App No.' : appno,
        'Name' : name,
        'Degree' : nameOfDegreeFDHD,
        'Subject' : subjectFDHD,
        'Programme Type' : programmeType,
        'Gender' : gender,
        'Email_ID' : decodedEmail,
        'Mobile No.': mobileNo,
        'Department' : department,
        'Campus Preference (GOA1 GOA2 GOA3)' : campusPreference,
        // 'Broad Area wish to pursue PhD' : broadArea,
        'Percentage /CGPA in 12th' : twelthcCgpa,
        'Department for BE/BTech 1' : ugDegree1,
        'Percentage /CGPA in BE/BTech/BCA/BSc 1' : ugCgpa1,
        'Department for BE/BTech 2' : ugDegree2,
        'Percentage /CGPA in BE/BTech/BCA/BSc 2' : ugCgpa2,
        'Department for ME/MTech/MCA/MSc 1' : pgDegree1,
        'Percentage /CGPA in ME/MTech/MCA/MSc 1' : pgCgpa1,
        'Department for ME/MTech/MCA/MSc 2' : pgDegree2,
        'Percentage /CGPA in ME/MTech/MCA/MSc 2' : pgCgpa2,
        'Higher Degree from BITS programme' : isBitsStudent,
        'Cleared any of the following examination on: CSIR-NET/GATE/GPATE/UGC-NET/Any Other' : anyOtherExam,
        'Year of passing' : YearOfPassingFDHD,
        'Current Work Place' : currentWorkPlace,
        'HOD': '',
        'HOD Comment' : '',
        'DRC' : '',
        'DRC Comment' : '',
        'Eligible (Yes OR No)' : ''
      }

      tableData.push(obj)

      //Here clear all variables
      appno =''; department = ''; name = ''; gender = ''; decodedEmail = ''; mobileNo = '';
      campusPreference = ''; isBitsStudent = ''; currentWorkPlace = ''; qualifyExam = ''; nameOfDegreeFDHD = ''; subjectFDHD = '';
      cgpaFDHD = ''; YearOfPassingFDHD = ''; twelthcCgpa = ''; ugDegree1 = ''; ugCgpa1 = ''; ugSubject1 = ''; ugYearOfPassing1 = '';
      ugDegree2 = ''; ugCgpa2 = ''; ugSubject2 = ''; ugYearOfPassing2 = ''; pgDegree1 = ''; pgDegree2 = ''; pgCgpa1 = ''; pgCgpa2 = ''; 
      pgSubject1 = ''; pgSubject2 = ''; pgYearOfPassing1 = ''; pgYearOfPassing2 = ''; fdOrHd = ''; anyOtherExam = ''; broadArea = '';
      universityFDHD = ''; uguniversity1 = ''; uguniversity2 = ''; pguniversity1 = ''; pguniversity2 = '';

      // return tableData;
      return obj;
    };

    //Decode encoded email
    const decodeCfemail = (cfemail) => {
      let email = '';
      const key = parseInt(cfemail.substr(0, 2), 16);
    
      for (let i = 2; i < cfemail.length; i += 2) {
        const charCode = parseInt(cfemail.substr(i, 2), 16) ^ key;
        email += String.fromCharCode(charCode);
      }
    
      return email;
    };

    const handleButtonClick = () => {
      /**** For single file ****/
      let extractedData = [];
      if (uploadedFiles) {
        // Get the files selected by the user
        const files = uploadedFiles;

        files.sort(function(a,b) {
          if (a.text > b.text) return 1;
          if (a.text < b.text) return -1;
          return 0
        })

        extractedData = [];
      
        // Loop through the selected files
        for (let i = 0; i < files.length; i++) {
          // Create a new FileReader for each file
          const reader = new FileReader();

          // FileReader's onload event handler
          reader.onload = function (fileEvent) {
            const contents = fileEvent.target.result;
            const fileName = files[i].name; // Get the current file name

            extractedData.push(parseTablesFromHTML(contents));

            //Create excel file at the end of the loop
            if((i + 1) === files.length){
              
            //  console.log('extractedData :: ', extractedData);

              const workbook = XLSXC.utils.book_new();
              const worksheet = XLSXC.utils.json_to_sheet(extractedData);

              //                   1            2           3           4          5            6           7             8      
              const colWidths = [{ wch: 7 },{ wch: 13 },{ wch: 15 },{ wch: 30 },{ wch: 20 },{ wch: 30 },{ wch: 25 },{ wch: 15 },
              //                      9           10          11          12         13           14          15         16                                         
                                 { wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },
              //                      17          18         19          20          21          22          23          24                    
                                 { wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },
              //                      25          26         27          28          29          30        
                                 { wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 }];

              worksheet['!cols'] = colWidths;

              var range = XLSXC.utils.decode_range(worksheet['!ref']);
              for(var R = range.s.r; R <= range.e.r; ++R) {
                for(var C = range.s.c; C <= range.e.c; ++C) {
                  var cellref = XLSXC.utils.encode_cell({c:C, r:R}); // construct A1 reference for cell
                  if(!worksheet[cellref]) continue; // if cell doesn't exist, move on
      
                  if(cellref.replace(/[^0-9]/g,"") === '1'){
                    worksheet[cellref].s = {
                      fill: {
                        patternType: "solid",
                        fgColor: { rgb: "FFE5B4" }
                      },
                      font: {
                        name: "Calibri",
                        sz: 12,
                        color: { rgb: "000000" },
                        bold: true,
                        italic: false,
                        underline: false
                      },
                      border: {
                        top: {style:'thin', color: {argb:'FF00FF00'}},
                        left: {style:'thin', color: {argb:'FF00FF00'}},
                        bottom: {style:'thin', color: {argb:'FF00FF00'}},
                        right: {style:'thin', color: {argb:'FF00FF00'}}
                      },
                      alignment: {vertical: 'top', horizontal: 'left', wrapText: true }
                    };
                  }else{
                    worksheet[cellref].s = {
                      fill: {
                        patternType: "solid",
                        fgColor: { rgb: "FFFFFF" }
                      },
                      font: {
                        name: "Calibri",
                        sz: 12,
                        color: { rgb: "000000" },
                        bold: false,
                        italic: false,
                        underline: false
                      },
                      border: {
                        top: {style:'thin', color: {argb:'FF00FF00'}},
                        left: {style:'thin', color: {argb:'FF00FF00'}},
                        bottom: {style:'thin', color: {argb:'FF00FF00'}},
                        right: {style:'thin', color: {argb:'FF00FF00'}}
                      },
                      alignment: {vertical: 'top', horizontal: 'left',indent: 1, wrapText: true }
                    };
                  }
                }
              }

              XLSXC.utils.book_append_sheet(workbook, worksheet, 'Data');
              XLSXC.writeFile(workbook, `${programmeType} PhD Admission Data.xlsx`); 
            }

            // Do something with the 'contents' here (e.g., display or process it)
            // displayContents(fileName, contents);
          };

          reader.readAsText(files[i]);
        }
        slno = 0; 
      }
    };

    const processButtonClick = (event) => {
      if (selectedFile) {
        const f = selectedFile;
        var name = f.name;
        const reader = new FileReader();
        reader.onload = (evt) => {
          let data = evt.target.result;
          data = new Uint8Array(data);
          const workbook = XLSXC.read(data, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];

          var range = XLSXC.utils.decode_range(worksheet['!ref']);

          //                   1            2           3           4          5            6           7             8      
          const colWidths = [{ wch: 7 },{ wch: 13 },{ wch: 15 },{ wch: 30 },{ wch: 20 },{ wch: 30 },{ wch: 25 },{ wch: 15 },
            //                      9           10          11          12         13           14          15         16                                         
                               { wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },
            //                      17          18         19          20          21          22          23          24                    
                               { wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },
            //                      25          26         27          28          29          30        
                               { wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 },{ wch: 20 }];

          worksheet['!cols'] = colWidths;

          var cellColor = 0;
          var isTwelthTrue = false, isCsUgTrue = false, isCSPgTrue = false;
          var twelthPerc = 0;
          var csdontinclude = 0;

          if(CstwelthPerc.length === 0){
            toast.warn("Please enter CS 12th criteria")
            return;
          }

          if(etctwelthPerc.length === 0){
            toast.warn("Please enter EEE 12th criteria")
            return;
          }

          if(othertwelthPerc.length === 0){
            toast.warn("Please enter 12th criteria")
            return;
          }

          var isEmpty = 0;

          for(var R = range.s.r; R <= range.e.r; ++R) {
            for(var C = range.s.c; C <= range.e.c; ++C) {
              var cellref = XLSXC.utils.encode_cell({c:C, r:R}); // construct A1 reference for cell
              if(!worksheet[cellref]) continue; // if cell doesn't exist, move on
              var cell = worksheet[cellref];

              isEmpty = 0;

              if(cellref.replace(/[^0-9]/g,"") === '1'){
                worksheet[cellref].s = {
                  fill: {
                    patternType: "solid",
                    fgColor: { rgb: "FFE5B4" }
                  },
                  font: {
                    name: "Calibri",
                    sz: 12,
                    color: { rgb: "000000" },
                    bold: true,
                    italic: false,
                    underline: false
                  },
                  border: {
                    top: {style:'thin', color: {argb:'FF00FF00'}},
                    left: {style:'thin', color: {argb:'FF00FF00'}},
                    bottom: {style:'thin', color: {argb:'FF00FF00'}},
                    right: {style:'thin', color: {argb:'FF00FF00'}}
                  },
                  alignment: {vertical: 'top', horizontal: 'left',indent: 1, wrapText: true}
                };
              }else{
                worksheet[cellref].s = {
                  border: {
                    top: {style:'thin', color: {argb:'FF00FF00'}},
                    left: {style:'thin', color: {argb:'FF00FF00'}},
                    bottom: {style:'thin', color: {argb:'FF00FF00'}},
                    right: {style:'thin', color: {argb:'FF00FF00'}}
                  },
                  alignment: {vertical: 'top', horizontal: 'left',indent: 1, wrapText: true}
                };

                // console.log('test 0 :: ',cellref);
                // console.log(cellref.substring(0, 1), ' : ', cellref.replace(/[^0-9]/g,""));

                // Check CS 12th percentage
                if (cellref === 'M'+ (R + 1)){
                  twelthPerc = cell.v.trim();
                }

                let degreeSubject = '';

                isEmpty = 0;

                if (cellref === 'N'+ (R + 1)){
                  var degreestring = (cell.v.split('(')[0]).replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s/g, '').toLowerCase().trim();
                  degreeSubject = ((cell.v.substring(cell.v.indexOf("(") + 1)).slice(0, -1)).replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s/g, '').toLowerCase().trim();
                    
                  if(worksheet['M'+ (R + 1)].v.trim().length === 0){
                      isEmpty = 1;

                      worksheet['M'+ (R + 1)].v = 'NA';
                      worksheet['AD'+ (R + 1)].v = 'No';
                      worksheet['M'+ (R + 1)].s = {
                      fill: {
                        patternType: "solid",
                        fgColor: { rgb: "e6ffe6" }
                      },
                      font: {
                        name: "Calibri",
                        sz: 12,
                        color: { rgb: "000000" },
                        bold: true,
                        italic: false,
                        underline: false
                      },
                        border: {
                        top: {style:'thin', color: {argb:'FF00FF00'}},
                        left: {style:'thin', color: {argb:'FF00FF00'}},
                        bottom: {style:'thin', color: {argb:'FF00FF00'}},
                        right: {style:'thin', color: {argb:'FF00FF00'}}
                      },
                        alignment: {vertical: 'top', horizontal: 'left',indent: 1, wrapText: true }
                      };
                  }

                  if((degreestring.includes('be') && degreestring.length === 2) || degreestring.includes('bachelorofengineering') || degreestring.includes('computerengineering')
                  || degreestring.includes('btech')){

                    // If CS department
                    if(degreeSubject.includes('datastructure') || (degreeSubject.includes('dbms')  && degreeSubject.length === 4) || degreeSubject.includes('cryptography') || degreeSubject.includes('compilers') || degreeSubject.includes('network') || degreeSubject.includes('computerscience') || degreeSubject.includes('computerengineering')
                    || (degreeSubject.includes('cs') && degreeSubject.length === 2) || degreeSubject.includes('informationsystem')  || (degreeSubject.includes('is') && degreeSubject.length === 2)  || (degreeSubject.includes('it') && degreeSubject.length === 2) || degreeSubject.includes('informationtechnology')
                    || degreeSubject.includes('computerscience') || (degreeSubject.includes('cse') && degreeSubject.length === 3) || degreeSubject.includes('informationscience')
                    ){
                        isTwelthTrue  = false;

                        if(!isEmpty){
                          if(worksheet['M'+ (R + 1)].v.trim().length > 0 && CstwelthPerc > 0){
                            if(CstwelthPerc.length > 0){
                              if(parseInt(selectCsTwelthPerc) === 0 && twelthPerc >= CstwelthPerc){
                                isTwelthTrue = true;
                              }else if(parseInt(selectCsTwelthPerc) === 1 && twelthPerc <= CstwelthPerc){
                                isTwelthTrue = true;
                              }
                              else if(parseInt(selectCsTwelthPerc) === 2 && twelthPerc === CstwelthPerc){
                                isTwelthTrue = true;
                              }
                              else{
                                isTwelthTrue = false;
                              }
                            }

                            if(!isTwelthTrue){
                              worksheet['AD'+ (R + 1)].v = 'No';
                              worksheet['M'+ (R + 1)].s = {
                                fill: {
                                  patternType: "solid",
                                  fgColor: { rgb: "FFCCCB" }
                                },
                                font: {
                                  name: "Calibri",
                                  sz: 12,
                                  color: { rgb: "000000" },
                                  bold: false,
                                  italic: false,
                                  underline: false
                                },
                                border: {
                                  top: {style:'thin', color: {argb:'FF00FF00'}},
                                  left: {style:'thin', color: {argb:'FF00FF00'}},
                                  bottom: {style:'thin', color: {argb:'FF00FF00'}},
                                  right: {style:'thin', color: {argb:'FF00FF00'}}
                                },
                                alignment: {vertical: 'top', horizontal: 'left',indent: 1, wrapText: true }
                              };
                            }
                          }
                        }
                      }

                      // If EEE department
                      else if (degreeSubject.includes('telcome') || degreeSubject.includes('elect') || degreeSubject.includes('instrumentation')
                        || degreeSubject.includes('electrical') || (degreeSubject.includes('etc') && degreeSubject.length === 3) || (degreeSubject.includes('ece') && degreeSubject.length === 3) || (degreeSubject.includes('dsp') && degreeSubject.length === 3)
                        || (degreeSubject.includes('e&c') && degreeSubject.length === 3) || (degreeSubject.includes('eee') && degreeSubject.length === 3)){

                          isTwelthTrue  = false;
                          if(!isEmpty){
                            if(worksheet['M'+ (R + 1)].v.trim().length > 0 && etctwelthPerc > 0){
                              if(etctwelthPerc.length > 0){
                                if(parseInt(selectEtcTwelthPerc) === 0 && twelthPerc >= etctwelthPerc){
                                  isTwelthTrue = true;

                                }else if(parseInt(selectEtcTwelthPerc) === 1 && twelthPerc <= etctwelthPerc){
                                  isTwelthTrue = true;
                                }
                                else if(parseInt(selectEtcTwelthPerc) === 2 && twelthPerc === etctwelthPerc){
                                  isTwelthTrue = true;
                                }
                                else{
                                  isTwelthTrue = false;
                                }
                              }
    
                              if(!isTwelthTrue){
                                worksheet['AD'+ (R + 1)].v = 'No';
                                worksheet['M'+ (R + 1)].s = {
                                  fill: {
                                    patternType: "solid",
                                    fgColor: { rgb: "FFCCCB" }
                                  },
                                  font: {
                                    name: "Calibri",
                                    sz: 12,
                                    color: { rgb: "000000" },
                                    bold: false,
                                    italic: false,
                                    underline: false
                                  },
                                  border: {
                                    top: {style:'thin', color: {argb:'FF00FF00'}},
                                    left: {style:'thin', color: {argb:'FF00FF00'}},
                                    bottom: {style:'thin', color: {argb:'FF00FF00'}},
                                    right: {style:'thin', color: {argb:'FF00FF00'}}
                                  },
                                  alignment: {vertical: 'top', horizontal: 'left',indent: 1, wrapText: true }
                                };

                                // worksheet['AD'+ (R + 1)].v = 'NO';
                              }
                            }
                          }  
                      }
                    }
                  }
                  else{
                    // If M.Sc/MCA/M.Tech department
                    if ((degreeSubject.includes('mca') && degreeSubject.length === 3) || (degreeSubject.includes('msc') && degreeSubject.length === 3)
                    || degreeSubject.includes('masterofcomputerapplicaion') || degreeSubject.includes('masteroftechnology') || degreeSubject.includes('masterofscience')){
                      isTwelthTrue  = false;

                      if(!isEmpty){
                          if(worksheet['M'+ (R + 1)].v.trim().length > 0 && othertwelthPerc > 0){
                            if(othertwelthPerc.length > 0){
                              if(parseInt(selectOtherTwelthPerc) === 0 && twelthPerc >= othertwelthPerc){
                                isTwelthTrue = true;
                              }else if(parseInt(selectOtherTwelthPerc) === 1 && twelthPerc <= othertwelthPerc){
                                isTwelthTrue = true;
                              }
                              else if(parseInt(selectOtherTwelthPerc) === 2 && twelthPerc === othertwelthPerc){
                                isTwelthTrue = true;
                              }
                              else{
                                isTwelthTrue = false;
                              }
                            }

                            if(!isTwelthTrue){
                              worksheet['M'+ (R + 1)].s = {
                                fill: {
                                  patternType: "solid",
                                  fgColor: { rgb: "FFCCCB" }
                                },
                                font: {
                                  name: "Calibri",
                                  sz: 12,
                                  color: { rgb: "000000" },
                                  bold: false,
                                  italic: false,
                                  underline: false
                                },
                                border: {
                                  top: {style:'thin', color: {argb:'FF00FF00'}},
                                  left: {style:'thin', color: {argb:'FF00FF00'}},
                                  bottom: {style:'thin', color: {argb:'FF00FF00'}},
                                  right: {style:'thin', color: {argb:'FF00FF00'}}
                                },
                                alignment: {vertical: 'top', horizontal: 'left',indent: 1, wrapText: true }
                              };
                            }
                          }
                      }
                    }
                  }
                }
            }
          }

          XLSXC.writeFile(workbook, name);
        };

        reader.readAsArrayBuffer(selectedFile);
      }
      else{
        toast.warn("Please select file")
        return;
      }
    };

    // Helper function to display the contents
    function displayContents(fileName, contents) {
      const outputDiv = document.getElementById('output');
      const fileDiv = document.createElement('div');
      fileDiv.innerHTML = `<strong>${fileName}:</strong> <br>${contents}<br><br>`;
      outputDiv.appendChild(fileDiv);
    }

    /*************************************************************** */
    // const handlebook =async (event) => {
    //    validateExcelFile(selectedFile);
    // };

    // const validateExcelFile = async () => {
    //   return new Promise((resolve, reject) => {
    //     const reader = new FileReader();
    //     reader.readAsArrayBuffer(selectedFile);
    //     reader.onload = () => {
    //       const buffer = reader.result;
    //       const workbook = new ExcelJS.Workbook();
    //       // workbook.xlsx
    //       //   .load(buffer)
    //       //   .then(() => {
    //       //     const worksheet = workbook.getWorksheet(1);
    //       //     const rowCount = worksheet.getColumn(1).values;
    //       //     const columnCount = worksheet.getRow(1);

    //       //     const worksheet1 = workbook.addWorksheet('New Sheet', {properties:{tabColor:{argb:'FFC0000'}}});

    //       //     worksheet.getCell('A1').value = 'testing';

    //       //     // console.log(worksheet.getCell('A1').value);

    //       //     //  validate table data
    //       //     console.log('rowCount : ',rowCount);
    //       //     console.log('columnCount : ',columnCount);
    //       //     resolve('Excel file is valid.');
    //       //     // buffer writing
    //       //     workbook.xlsx.writeBuffer(buffer);
    //       //   })
    //       //   .catch((error) => {
    //       //     reject('Error occurred while loading the workbook.');
    //       //   });

    //       await workbook.xlsx.readFile(file);
    //     };
    //   });
    // };

    // const handleChange1 = (e) => {
    //   const file = selectedFile;
    //   const workbook = new Excel.Workbook();
    //   const reader = new FileReader();

    //   reader.readAsArrayBuffer(file);
    //   reader.onload = () => {
    //       const buffer = reader.result;
    //       workbook.xlsx.load(buffer).then((workbook) => {
    //         const worksheet = workbook.getWorksheet(1);
    //         const rowCount = worksheet.getColumn(1).values;
    //         const columnCount = worksheet.getRow(1);

    //         worksheet.getCell('A1').value = 'testing';

    //         // workbook.xlsx.readFile(file);

    //         // console.log(workbook, 'workbook instance');
    //         // workbook.eachSheet((sheet, id) => {
    //         //     sheet.eachRow((row, rowIndex) => {
    //         //         console.log(row.values, rowIndex);
    //         //     });
    //         // });
    //       });
    //   };
    // };
    /*************************************************************** */

    return (
        <div>
            {/* <input type="file" onChange={handleFileChange} /> */}
            <input type="file" multiple onChange={handleFileEvent} />
            <button onClick={handleButtonClick}>Extract Data</button>
            {/* <button onClick={handleChange1}>Test Button</button> */}

            <div>
              <input type="file" onChange={handleFileChange} />
              <button onClick={processButtonClick}>Process</button>
              <ToastContainer />
              <table>
                <tbody>
                  <tr>
                    <td>CS</td>
                    <td>12th</td>
                    <td>
                      <select id="cstwelthselect" onChange={selectCsTwelthPercFunc}>
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="cstwelthperc" onChange={handleCsTwelthPercFunc} value={CstwelthPerc} /><label>%</label></td>
                    <td>UG</td>
                    <td>
                      <select name="" id="csselect">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="csugperc" /><label>%</label></td>
                    <td>PG</td>
                    <td>
                      <select id="csugtwelthselect" >
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="cstwelthperc"  /><label>%</label></td>
                    <td>PG @ BITS</td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label>%</label></td>
                  </tr>
                  <tr>
                    <td>Non CS</td>
                  </tr>
                  <tr>
                    <td>Electrical Science</td>
                    <td>12th</td>
                    <td>
                      <select id="etctwelthselect" onChange={selectEtcTwelthPercFunc}>
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="etctwelthperc" onChange={handleEtcTwelthPercFunc} value={etctwelthPerc} /><label>%</label></td>
                    <td>UG</td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label>%</label></td>
                    <td>PG</td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label>%</label></td>
                    <td>PG @ BITS</td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label htmlFor="">%</label></td>
                  </tr>
                  <tr>
                    <td>BSc + MSc/MTech OR MCA + MTech/MSc</td>
                    <td>12th</td>
                    <td>
                      <select id="othertwelthselect" onChange={selectOtherTwelthPercFunc}>
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="etctwelthperc" onChange={handleOtherTwelthPercFunc} value={othertwelthPerc} /><label>%</label></td>
                    <td>UG</td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label htmlFor="">%</label></td>
                    <td>PG</td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label htmlFor="">%</label></td>
                    <td>PG @ BITS</td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label htmlFor="">%</label></td>
                  </tr>
                  {/* <tr>
                    <td>Others</td>
                    <td>12th</td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label htmlFor="">%</label></td>
                    <td><label htmlFor="">UG</label></td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label htmlFor="">%</label></td>
                    <td><label htmlFor="">PG</label></td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label htmlFor="">%</label></td>
                    <td><label htmlFor="">PG @ BITS</label></td>
                    <td>
                      <select name="" id="">
                        <option value="0">&ge;</option>
                        <option value="1">&le;</option>
                        <option value="2">&#61;</option>
                      </select>
                    </td>
                    <td><input type="number" id="" /><label htmlFor="">%</label></td>
                  </tr> */}
                </tbody>
              </table>
            </div>
           
            <div id="output"></div>
        </div>
    );
}

export default TableParser;