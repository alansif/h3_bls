const JSZip = require('jszip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const path = require('path');
const content = fs.readFileSync(path.resolve(__dirname, 'b.docx'), 'binary');
const zip = new JSZip(content);
const doc = new Docxtemplater();
doc.loadZip(zip);

const express = require("express");
const bodyParser = require("body-parser");
const app = express();
const port = 8315;

app.use(express.static('C:/Users/Administrator/dev/h3_blc/dist'));

app.use(bodyParser.urlencoded({ extended: true }));

app.use(function(req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    res.header("Access-Control-Allow-Methods", "*");
    next();
});

const sql = require('mssql');

const config = {
    user: 'sa',
    password: 'Angelwin123',
    server: '192.168.160.101',
    database: 'EISDATA',
    options: {
    }
};

const pool1 = new sql.ConnectionPool(config);
const pool1Connect = pool1.connect();

pool1.on('error', err => {
    // ... error handler
});

const sqlstr = "SELECT RITB.DJH1,XM1,XB1,NL1,BGTXT,ZDJL,YSXM FROM RITB2 left join RITB on RITB2.DJH1=RITB.DJH1 ";

async function messageHandler(datestr) {
    const ss = sqlstr + "where ReportedDate=@datestr order by RITB.DJH1";
    await pool1Connect; // ensures that the pool has been created
    try {
    	const request = new sql.Request(pool1)
        const result = await request.input("datestr", datestr).query(ss);
    	return result;
    } catch (err) {
        console.error('SQL error', err);
        throw err;
    }
}

app.get("/api/query", function(req, res){
    let datestr = req.query['date'] || new Date().toISOString().substr(0, 10);
    let f = async() => {
        try{
            let result = await messageHandler(datestr);
            res.status(200).json(result.recordset);
        } catch(err) {
            res.status(500).json(err);
        }
    };
    f();
});

app.post("/api/pathology/:djh/make", function(req, res){
    const djh = req.params['djh'] || '';
    if (djh.length === 0) {
        res.status(400).end();
        return;
    }
    const ss = sqlstr + "where RITB2.DJH1=@djh";
    let f = async() => {
        await pool1Connect; // ensures that the pool has been created
        try {
            const request = new sql.Request(pool1)
            const result = await request.input("djh", djh).query(ss);
            if (result.recordset.length > 0) {
                const d = result.recordset[0];
                doc.setData({
                    xm: d.XM1,
                    xb: d.XB1,
                    nl: d.NL1,
                    djh: d.DJH1,
                    bgtxt: d.BGTXT,
                    zdjl: d.ZDJL,
                    ysxm: d.YSXM
                });
                doc.render();
                const buf = doc.getZip().generate({type: 'nodebuffer'});
                res.writeHead(200, {
                    'Content-Type': 'application/octet-stream',
                    'Content-Disposition': 'attachment; filename=some.docx',
                    'Content-Length': buf.length
                });
                res.end(buf);
//                fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);
            }
        } catch (err) {
            console.error('error', err);
            res.status(500).json(err);
        }
    };
    f();
});

app.listen(port, () => {
    console.log("Server is running on port " + port + "...");
});
