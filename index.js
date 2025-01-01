const express = require('express');
const multer = require('multer');
const excelToJson = require('convert-excel-to-json');
const fsExtra = require('fs-extra');
const path = require('path'); // Import path module for handling file paths
const puppeteer = require('puppeteer');

require('dotenv').config();

const app = express();
const PORT = 5001;
const upload = multer({ dest: "uploads/" });

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname, 'public')));

const { sendPdfEmail, formatDate, createHTML, byptmapping } = require('./utils');

app.listen(PORT, () => {
    console.log(`App running at http://localhost:${PORT}`);
});

app.get('/', (req, res) => {
    res.render('homepage');
});

// app.post('/upload/techm', upload.single('file'), async (req, res) => {

//     var filePath = "";
//     if (req.file == null || req.file.filename == 'undefined') {
//         res.status(400).json("No file uploaded");
//     }
//     else {
//         filePath = 'uploads/' + req.file.filename;
//         const excelData = excelToJson({
//             sourceFile: filePath,
//             header: {
//                 rows: 1
//             },
//             columnToKey: {
//                 A: 'empId',             // EMP ID
//                 B: 'type',              // Type
//                 C: 'name',              // Name
//                 D: 'basic',             // Basic Salary
//                 E: 'hra',               // HRA
//                 F: 'conAllo',           // Con Allo (Conveyance Allowance)
//                 G: 'bonus',             // Bonus
//                 H: 'tfp',               // TFP (Total Fixed Pay)
//                 I: 'variablePay',       // Variable Pay
//                 J: 'total',             // Total Salary
//                 K: 'protax',            // Protax (Professional Tax)
//                 L: 'tds',               // TDS (Tax Deducted at Source)
//                 M: 'leaves',            // Leaves Taken
//                 N: 'totalDeduction',    // Total Deduction
//                 O: 'netSalary',         // Net Salary
//                 P: 'presentDays',       // Present Days
//                 Q: 'payDays',           // Pay Days
//                 R: 'amtInWords',        // Amount in Words
//                 S: 'designation',       // Designation
//                 T: 'accountNo'          // A/c No (Account Number)
//             },
//         });

//         const data = excelData.Sheet1

//         for (const i of data) {
//             const browser = await puppeteer.launch();
//             const page = await browser.newPage();

//             const date = new Date();
//             const monthNames = [
//                 "January", "February", "March", "April", "May", "June",
//                 "July", "August", "September", "October", "November", "December"
//             ];
//             const currentMonth = monthNames[date.getMonth()];

//             const htmlfile = `
//                         <!DOCTYPE html>
//                         <html lang="en">
//                         <head>
//                             <meta charset="UTF-8">
//                             <meta name="viewport" content="width=device-width, initial-scale=1.0">
//                             <title>Salary Slip</title>
//                             <style>
//                                 body {
//                                     font-family: Arial, sans-serif;
//                                     margin: 0;
//                                     padding: 0;
//                                     background-color: #f9f9f9;
//                                 }

//                                 .container {
//                                     width: 800px;
//                                     margin: 20px auto;
//                                     background: #f5f5f5;
//                                     padding: 20px;
//                                     border: 2px solid black;
//                                     box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
//                                 }

//                                 .header {
//                                     text-align: center;
//                                     margin-bottom: 20px;
//                                 }

//                                 .header img {
//                                     max-width: 150px;
//                                     height: 120px;
//                                     margin-bottom: 10px;
//                                     border: 2px solid black;
//                                 }

//                                 table {
//                                     width: 100%;
//                                     border-collapse: collapse;
//                                     margin-bottom: 20px;
//                                 }

//                                 table th, table td {
//                                     border: 2px solid black;
//                                     padding: 8px;
//                                     text-align: left;
//                                 }

//                                 table th {
//                                     background-color: #f4f4f4;
//                                     font-weight: bold;
//                                 }

//                                 .footer {
//                                     text-align: center;
//                                     font-size: 14px;
//                                     color: #555;
//                                 }

//                                 .highlight {
//                                     font-weight: bold;
//                                     color: #333;
//                                 }
//                             </style>
//                         </head>
//                         <body>
//                             <div class="container">
//                                 <div class="header">
//                                     <img src="https://bypeopletechnologies.com/wp-content/uploads/2017/01/byPeople-Logo.png" alt="Company Logo">
//                                     <h1>byPeople Technologies</h1>
//                                     <h2>Regd Office: Z-208, Dev Castle, Opp. Radhe Krishna Complex, Isanpur, Ahmedabad-382443</h2>
//                                     <h2>Salary Slip for ${currentMonth}, ${date.getFullYear()}</h2>
//                                 </div>

//                                 <table>
//                                     <tr>
//                                         <td>Ref. No.: <span class="highlight">${i.empId}</span></td>
//                                         <td>Employee Name: <span class="highlight">${i.name}</span></td>
//                                     </tr>
//                                     <tr>
//                                         <td>PF No.: <span class="highlight">${i.type}</span></td>
//                                         <td>Pay Days: <span class="highlight">${i.payDays}</span></td>
//                                     </tr>
//                                     <tr>
//                                         <td>Present Days: <span class="highlight">${i.presentDays}</span></td>
//                                         <td>Designation: <span class="highlight">${i.designation}</span></td>
//                                     </tr>
//                                     <tr>
//                                         <td>Branch: <span class="highlight">Work From Home</span></td>
//                                         <td>A/c. No.: <span class="highlight">${i.accountNo}</span></td>
//                                     </tr>
//                                 </table>

//                                 <table>
//                                     <thead>
//                                         <tr>
//                                             <th>Earnings</th>
//                                             <th>Amount</th>
//                                             <th>Deductions</th>
//                                             <th>Amount</th>
//                                         </tr>
//                                     </thead>
//                                     <tbody>
//                                         <tr>
//                                             <td>BASIC</td>
//                                             <td>${i.basic}</td>
//                                             <td>PROFESSIONAL TAX</td>
//                                             <td>${i.protax}</td>
//                                         </tr>
//                                         <tr>
//                                             <td>HRA</td>
//                                             <td>${i.hra}</td>
//                                             <td>TDS</td>
//                                             <td>${i.tds}</td>
//                                         </tr>
//                                         <tr>
//                                             <td>CONVAYNCE ALLOWANCE</td>
//                                             <td>${i.conAllo}</td>
//                                             <td></td>
//                                             <td></td>
//                                         </tr>
//                                         <tr>
//                                             <td>BONUS</td>
//                                             <td>${i.bonus}</td>
//                                             <td></td>
//                                             <td></td>
//                                         </tr>
//                                         <tr>
//                                             <td>TFP</td>
//                                             <td>${i.tfp}</td>
//                                             <td></td>
//                                             <td></td>
//                                         </tr>
//                                         <tr>
//                                             <td>Variable Pay</td>
//                                             <td>${i.variablePay}</td>
//                                             <td></td>
//                                             <td></td>
//                                         </tr>
//                                         <tr>
//                                             <th>Total</th>
//                                             <th>${i.total}</th>
//                                             <th>Total</th>
//                                             <th>${i.totalDeduction}</th>
//                                         </tr>
//                                         <tr>
//                                             <th colspan="1">NET PAY</th>
//                                             <th colspan="3">${i.netSalary}</th>
//                                         </tr>
//                                     </tbody>
//                                 </table>

//                                 <table>
//                                     <tr>
//                                         <td>In Words: Rs.</td>
//                                         <td><span class="highlight">${i.amtInWords}</span></td>
//                                     </tr>
//                                     <tr>
//                                         <td colspan="2" style="text-align: center;">This is a computer-generated salary slip. Hence doesn’t require any signature.</td>
//                                     </tr>
//                                     <tr>
//                                         <td colspan="2" style="text-align: center;">Thank You for your efforts</td>
//                                     </tr>
//                                 </table>
//                             </div>
//                         </body>
//                         </html>
//                         `;

//             await page.setContent(htmlfile);
//             const fn = `${i.email}_${i.name}.pdf`;
//             const outputPath = path.join(__dirname, 'uploads', fn);
//             const fp = 'uploads/' + fn;
//             await page.pdf({ path: outputPath, format: 'A4' });

//             console.log('PDF generated for', `${i.name}`);

//             try {
//                 await sendPdfEmail(fn);
//             } catch (error) {
//                 console.log(error);
//             }

//             console.log(`Mail sent to ${i.name}'s email `);

//             fsExtra.remove(fp);
//             console.log('PDF cleared for', `${i.name}`);
//             await browser.close();

//         }

//         fsExtra.remove(filePath);
//         // res.status(200).send(excelData);
//         res.redirect('/');
//     }
//     // res.redirect("/");

// })

// api endpoint for mailing slip to every employee


app.post('/upload/bypt', upload.single('file'), async (req, res) => {

    var filePath = "";
    if (req.file == null || req.file.filename == 'undefined') {
        res.status(400).json("No file uploaded");
    }

    filePath = 'uploads/' + req.file.filename;
    const excelData = excelToJson({
        sourceFile: filePath,
        header: {
            rows: 1
        },
        // mapping to excel file cell locations
        columnToKey: byptmapping
    });

    const data = excelData.Sheet1

    for (const i of data) {
        i.paymentDate = formatDate(i.paymentDate);
        i.dateOfJoin = formatDate(i.dateOfJoin);
    
        const browser = await puppeteer.launch();
        const page = await browser.newPage();
        
        let htmlfile = "";
        htmlfile = createHTML(htmlfile, i);

        await page.setContent(htmlfile);
        const fn = `${i.email}_${i.name}.pdf`;
        const outputPath = path.join(__dirname, 'uploads', fn);
        const fp = 'uploads/' + fn;
        await page.pdf({ path: outputPath, format: 'A4' });

        console.log('PDF generated for', `${i.name}`);

        try {
            // in sendPdfEmail set second parameter to i.email, here test email used for dev
            await sendPdfEmail(fn, `5022002a@gmail.com`);
        } catch (error) {
            console.log(error);
        }

        console.log(`Mail sent to ${i.name}'s email `);

        fsExtra.remove(fp);
        console.log('PDF cleared for', `${i.name}`);
        await browser.close();

    }

    fsExtra.remove(filePath);
    // res.status(200).send(excelData);
    res.redirect('/');
})

// Serve static files from the uploads directory
app.use('/files', express.static(path.join(__dirname, 'uploads')));

// api endpoint to store files inside upload folder
app.post('/makepdf', upload.single('file'), async (req, res) => {

    console.log("Hello")

    var filePath = "";
    if (req.file == null || req.file.filename == 'undefined') {
        res.status(400).json("No file uploaded");
    }
    else {
        filePath = 'uploads/' + req.file.filename;
        const excelData = excelToJson({
            sourceFile: filePath,
            header: {
                rows: 1
            },
            columnToKey: byptmapping
        });

        const data = excelData.Sheet1

        for (const i of data) {
        
            i.paymentDate = formatDate(i.paymentDate);
            i.dateOfJoin = formatDate(i.dateOfJoin);
        
            const browser = await puppeteer.launch();
            const page = await browser.newPage();

            let htmlfile = "";
            htmlfile = createHTML(htmlfile, i);

            await page.setContent(htmlfile);
            const fn = `${i.email}_${i.name}.pdf`;
            const outputPath = path.join(__dirname, 'uploads', fn);
            const fp = 'uploads/' + fn;
            await page.pdf({ path: outputPath, format: 'A4' });

            console.log('PDF generated for', `${i.name}`);
        
        }
    }
    fsExtra.remove(filePath);
    res.send("All pdfs generated")

})

