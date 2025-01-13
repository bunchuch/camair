const fs = require('fs')

const path = require('path');


const ExcelJS = require("exceljs");
const { time } = require('console');
// TicketNumber/ PaxName /SectorFlight.

exports.index = (req ,res) => {
    res.render('home', {title : "text to excel converter"});

};

exports.uploadForm = (req ,res) => {
     res.render('upload', {title : "upload text file"});
}

// exports.converTextToExcel = (req, res) => {
//     try {
//         const textFilePath = req.file.path;
//         const filename = req.file.originalname
//         const textData = fs.readFileSync(textFilePath, 'utf-8');
//         const rows = textData.split('\n').map(row => row.split('|'));

//         const columnName = ['TicketNumber', "PaxName", "SectorFligth"]
//         rows.unshift(columnName)

        
//         xlsx.utils.book_append_sheet(workbook , worksheet ,"Sheet1");
//         const getfilename = filename.replace(/(@|\.txt)/g, '').trim();
//         const excelFilePath = path.join(__dirname , '../public', `${getfilename}.xlsx`);
//         xlsx.writeFile(workbook , excelFilePath);

//         // res.send(`
//         //     <h1>Excel file created successfully!</h1>
//         //     <a href="/${excelFilePath}" download>Download Excel File</a>
//         // `);

//         // res.download(excelFilePath , `${getfilename}.xlsx`, (err)=> {
//         //     if(err) throw err
//         //     fs.unlinkSync(textFilePath);
//         //     fs.unlinkSync(excelFilePath)
//         // });

//         res.render("table", {data: data})
//     } catch (error) {
//         console.log(error)
//         res.status(500).send("Error processing file");
//     }
// }


function checkFileName(filename , keyword){
    const lowerFileName = filename.toLowerCase();
    const lowerKeyword = keyword.toLowerCase();

    //check if the key keywork exist in the filename
    return lowerFileName.includes(lowerKeyword)
}

const transformData = (data) => {
    const result = [];
    const sectorFlights = data.slice(0, 2);
    const ticketNumbers = data.slice(2, 6);
    const paxNames = data.slice(6);
  
    for (let i = 0; i < ticketNumbers.length; i++) {
      const ticketNumber = ticketNumbers[i][0].slice(0, 9); // Take the first 9 characters of the ID
      const paxName = paxNames[i]?.[0] || ""; // Handle missing names gracefully
      const sectorFlight = sectorFlights[i]?.[0] || ""; // Handle missing routes gracefully
      result.push([ticketNumber,paxName,sectorFlight].filter(Boolean));
    }
  
    return result;
  };

  function processRows(rows, regex) {
    return rows.map((row) => row.filter((value) => regex.test(value)));
}


exports.converTextToExcelWithMutiple = async (req, res) => {
    try {
        const files = req.files;
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("Text Data");
        var rowValues = []
        var ticketNumber = [];
        const regexPatterns = {
            resODflight: /^[A-Z]{3}$/, // Matches a sector/flight code
            resPassenger: /^[A-Z]+ [A-Z]+$/, // Matches a passenger name
            tktCouponHistory: /^\d{13}$/, // Matches a 13-digit ticket number
            tktEndorsement: /^\d{13}$/, // Matches a 13-digit endorsement
        };


        // Process each text file
        const columnName = ['TicketNumber', "PaxName", "SectorFligth"]
        sheet.addRow(columnName)
        const allData = [];
        for (const file of files) {
          const content = fs.readFileSync(file.path, "utf8");
          const rows = content
            .split("\n")
            .filter((line) => line.trim() !== "")
            .map((line) => line.split("|").map((item) => item.trim())); // Assuming comma-separated values
        
          if(checkFileName(file.originalname , 'resODflight')){
            const processedRows = processRows(rows , regexPatterns.resODflight);
            processedRows.forEach((value)=> {
                const sectorFlight = value.join('')
                rowValues[2] = sectorFlight
                sheet.insertRow(2, [...rowValues])
            })
               
            } else if (checkFileName(file.originalname , 'resPassenger')){
                const processedRows = processRows(rows , regexPatterns.resPassenger);
                processedRows.forEach((value)=> {
                    const passengerName = value.join('')
                    rowValues[1] = passengerName
                    sheet.insertRow(2, [...rowValues])

                })
            } else if (checkFileName(file.originalname, 'tktCouponHistory')){
                const processedRows = processRows(rows, regexPatterns.tktCouponHistory)
                processedRows.forEach((value)=> {
                   sheet.addRow(value)
                })
             }else if (checkFileName(file.originalname, 'tktEndorsement')){
                const processedRows = processRows(rows , regexPatterns.tktEndorsement)
                console.log(processedRows)
                processedRows.forEach((value)=> {
                    sheet.addRow(value)
                })   
             }else {
                console.log("No file name match")
            }
            fs.unlinkSync(file.path); // Clean up uploaded file
        }

   
        //ticketnumber to row
        // sheet.addRow(ticketNumber)
        console.log(allData)
        // swap data for all Data
        // sheet.addRow(allData)
        // const swapData = transformData(allData);
        // console.log(swapData);
        // Add data to the Excel sheet
        // sheet.addRows(swapData);

        
        //apply alternating row colors
        sheet.eachRow((row, rowIndex)=> {
            row.eachCell((cell)=> {
            
                if(rowIndex === 1){
                    //Header row styling
                    cell.font = {bold : false , color : {argb : "FFFFFFF"}}; //white text
                    cell.fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: "FF4F81BD" }, // Blue background
                      };
                      cell.alignment = { vertical: "middle", horizontal: "center" };
                }else{
                    cell.fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: {
                          argb: rowIndex % 2 === 0 ? "FFEFEFEF" : "FFFFFFFF", // Light gray or white
                        },
                      };
                }
            })
        })

      
    
       const headerRow = sheet.getRow(1); //Asumming the first row is the header
       headerRow.eachCell((cell)=> {
        cell.font = {bold : true , color: {argb : `FFFFFFFF`}};
        cell.fill = {
            type : "pattern",
            pattern : 'solid',
            fgColor : {argb : "FF4F81BD"}, //Blue background
        };
        cell.alignment = {vertical : "middle", horizontal: "center"};
       })
       sheet.columns.forEach((column)=> {
        column.width = 20
        // column.width = Math.max(...column.values.map((val)=> (val ? val.toString().length: 100 ))) + 2;
       });
       //if colum is found, delete it
        // Define a file path for the generated Excel file
        const outputPath = path.join(__dirname, "../public", "output.xlsx");
      
        // Write the Excel file to disk
        await  workbook.xlsx.writeFile(outputPath);
        const fileName = Date.now();
    //    res.render("view", {data: allData})
       //console.log(fileName)
       // Send the Excel file as a downloadable response
        res.download(outputPath, `${fileName}.xlsx`, (err) => {
          if (err) {
            console.error("File download error:", err);
          }
          // Clean up the generated file after download
          fs.unlinkSync(outputPath);
        });
    } catch (error) {
        console.log(error)
        res.status(500).json({
            message : "error",
            error : error
        })
    }
}

exports.downloadExcel = (req, res) => {
    const filePath = path.join(__dirname , "processed_data.xlsx");
    res.download(filePath, "processed_data.xlsx", (err) => {
        if (err) console.error(err);
      });
}