const express = require('express');
const app = express();
const formidable = require('formidable');
const fs = require('node:fs');
const path = require('node:path');
const cluster = require('node:cluster');
const OS = require('node:os');
const https = require('https');

const excelToJson = require('convert-excel-to-json');
const jsonDiff = require('json-diff');

// const multer  = require('multer');
// const http = require('node:http');
// const bodyParser = require('body-parser');
// const upload = multer({ dest: 'Temp/' })

PORT_1 = process.env.PORT || 1000;
app.set(express.static('public'));
app.set('view engine', 'ejs');

app.use(express.json());
app.use(express.urlencoded({extended: true})); 
// app.use(bodyParser.json());
// app.use(bodyParser.urlencoded({extended:true}));

app.get('/', async (req, res) => {
    // console.log(OS.cpus.length);
    // res.sendFile(path.join(__dirname, 'index.html'));
    var myDate = new Date();
    var hrs = myDate.getHours();
    var greet;
    if (hrs < 12)
      greet = 'Good Morning';
    else if (hrs >= 12 && hrs <= 17)
      greet = 'Good Afternoon';
    else if (hrs >= 17 && hrs <= 24)
      greet = 'Good Evening';

    res.render('index', {
        pageTitle: 'Compare two excel files here.',
        WelCome: 'WelCome!'
    })
});

app.post('/excel_compare_result', async (req, res) => {
//    cluster.fork;
//    cluster.fork;
    const start = Date.now();

    const form = new formidable.IncomingForm();
    let fields;
    let files;
    try {
        [fields, files] = await form.parse(req);
        if (!files) {
            return res.status(400).send(`<pre style="font-size:16pt;color:red;">No files were uploaded.!</pre>`);
        }
    }
    catch (err) {
        // example to check for a very specific error
        // if (err.code === formidableErrors.maxFieldsExceeded) {
        //     }
        console.error(err);
        // res.writeHead(err.httpCode || 400, { 'Content-Type': 'text/plain' });
        res.end(String(err));
        return;
        }
    // res.writeHead(200, { 'Content-Type': 'application/json' });
            // res.writeHead(err.httpCode || 400, { 'Content-Type': 'text/plain' });

   var odf = files.odf[0].filepath;
   var new_odf = path.join(__dirname , 'Temp', path.basename(odf).concat("_", files.odf[0].originalFilename));
   var cdf = files.cdf[0].filepath;
   var new_cdf =path.join(__dirname , 'Temp', path.basename(cdf).concat("_", files.cdf[0].originalFilename));

   fs.copyFileSync(odf, new_odf);
   fs.copyFileSync(cdf, new_cdf);
   console.log(`Uploaded file(s) get copied to ./Temp folder!`);

   const file1 = excelToJson({sourceFile: new_odf});
   console.log(`file1: ${(Date.now() - start)/1000} sec`);

   const file2 = excelToJson({sourceFile: new_cdf});   
   console.log(`file2: ${(Date.now() - start)/1000} sec`);
   
   fs.unlink(new_odf, function (err) {
    if (err) throw err;
    console.log(`${new_odf} File deleted!`);
  });
  fs.unlink(new_cdf, function (err) {
    if (err) throw err;
    console.log(`${new_cdf} File deleted!`);
  });

   var myindex1 = 0;
   Object.keys(file1).forEach((item,index) => {
       Object.keys(file1[item]).forEach((item1,index1) => {
           Object.keys(file1[item][item1]).forEach((item2,index2) => {
               myindex1 = Number(index1)+1;
               file1[item][item1][item2]= `<<<${myindex1}>>>: ${file1[item][item1][item2]}`;
           });
       });
   });
   console.log(`file1Key: ${(Date.now() - start)/1000} sec`);
   var myindex2 = 0;
   Object.keys(file2).forEach((item,index) => {
       Object.keys(file2[item]).forEach((item1,index1) => {
           Object.keys(file2[item][item1]).forEach((item2,index2) => {
               myindex2 = Number(index1)+1;
               file2[item][item1][item2]= `<<<${myindex2}>>>: ${file2[item][item1][item2]}`;
           });
       });
   });
   console.log(`file2Key: ${(Date.now() - start)/1000} sec`);

   var diffj = jsonDiff.diffString(file1, file2);
   
   console.log(`Prepare for HTML Preview: ${(Date.now() - start)/1000} sec`);
   
   const str = JSON.stringify(diffj, (key, val) => {
       return String(val).replace(/\n\s+\.\.\./g, ``).replace(/: "<<<(\d+)>>>:/g, '$1:');
   });
   
   console.log(JSON.parse(str));
   
   const str1 = JSON.stringify(diffj, (key, val) => {
       return String(val).replace(/\n\s+\.\.\./g, '')
       .replace(/\n\[31m/g, '\n<span style="color:red">')
       .replace(/\n\[32m/g, '\n<span style="color:green">')
       .replace(/\[39m\n/g, '</span>\n')
       .replace(/: "<<<(\d+)>>>: /g, '$1: "');
   });
   
//    fs.writeFileSync('exceldiff_log.txt', JSON.parse(str1), function (err) {
//      if (err) throw err;
//      console.log('Process completed in ', (Date.now() - start)/1000 , ' sec');
//    });

    if (diffj.replace(/\r?\n|\r/g, "").trim().length === 0 ) {
        console.log(`No Text Difference Found!: ${(Date.now() - start)/1000} sec\n`);
        res.send(`<!DOCTYPE html>
        <html>
            <head>
                <title>Excel Compare Result</title>
            </head>
            <body>
                <div>
                    <pre style="display:inline-block;font-size:16pt; border:4px black double;padding:10px;color:green;">No Text Difference Found!</pre>        
                </div>
            </body>
        </html>`);
    }
    else{
        res.send(`<!DOCTYPE html>
        <html>
            <head>
                <title>Excel Compare Result</title>
            </head>
            <body>
                <div>
                    <pre style="display:inline-block;font-size:14pt; border:4px black double;padding:10px;color:blue;">${JSON.parse(str1)}</pre>
                </div>
            </body>
        </html>`);
        console.log(`Process Completed: ${(Date.now() - start)/1000} sec\n`);
    }
});

app.listen(PORT_1, () => {
    console.log(`Example app listening on port ${PORT_1}`)
});
