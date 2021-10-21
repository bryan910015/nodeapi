const express = require('express');
var fs = require('fs');
var https = require('https');
const app = express(); //建立一個Express伺服器
const MorReserve = require('./common/MorReserve');
var cors = require('cors')
var bodyParser = require('body-parser');
app.use(cors())
var bodyParser = require('body-parser');
app.use(bodyParser.json({ limit: "50mb" }));
app.use(bodyParser.urlencoded({ limit: "50mb", extended: true }))

//環境變數global
//HTTPS設定
const port = process.env.PORT || 3001;
var hskey = fs.readFileSync('www.mornjoy.com.tw.key');
var hscert = fs.readFileSync('www.mornjoy.com.tw.pem');

var credentials = {
    key: hskey,
    cert: hscert
};

//切換開發/發行模式(參數 1 開發模式))
var IsProduction = process.argv[2];
if (IsProduction == '1') {

    app.listen(port, function () {
        console.log('API 開發模式執行中  ' + port);
    });
} else {

    https.createServer(credentials, app).listen(port, function () {
        console.log('API Production 模式執行中 ' + port);
    });
}
//以上是設定==================


//執行預約功能
app.post('/api/MorReserve', function (req, res) {
    let flag = req.body.flag;

    if (flag == 'read') {
        console.log('執行預約資料查詢' + new Date() + req.socket.remoteAddress);
        (async function () {
            let response = await MorReserve.GetReserveitem(req.body.sdate, req.body.edate, res);
            res.send(response);
        })()
    }

    if (flag == 'update') {
        console.log('執行預約資料更新' + new Date() + req.socket.remoteAddres);

        (async function () {
            let response = await MorReserve.UpdateReserveitem(req.body.data, res);
            res.send(response);
        })()

    }

    if (flag == 'delete') {
        //暫時不使用
        console.log('執行預約資料刪除' + new Date() + req.socket.remoteAddres);
        (async function () {
            let response = await MorReserve.DelReserveitem(req.body.datekey, res);
            res.send(response);
        })()
    }

    if (flag == 'readKs') {
        console.log('執行科室資料查詢' + new Date() + req.socket.remoteAddress);
        let response = MorReserve.GetKsReserveitem(req.body.sdate, req.body.edate, req.body.course, res);
    }

    if (flag == 'updateKs') {
        console.log('更新科室總人數' + new Date() + req.socket.remoteAddress);
        let response = MorReserve.UpdateKsReserveitem(req.body.data, res);
    }

});

app.post('/api/MorChkname', function (req, res) {
    console.log('執行姓名缺字查詢' + new Date() + req.socket.remoteAddress);
});

app.get('/api/MorReserveCommon', function (req, res) {
    console.log('執行科室下拉選單' + new Date() + req.socket.remoteAddress);

    (async function () {
        let response = await MorReserve.GetCourseData();
        res.send(response);
    })()

});

app.post('/api/MorReserveStat', function (req, res) {
    console.log('執行預約統計資料' + new Date() + req.socket.remoteAddress);
    var flag = req.body.flag;

    //預約統計
    if (flag == 'main') {
        var totalGoal = parseInt(req.body.entTotalGoal) + parseInt(req.body.singleTotalGoal);
        let response = MorReserve.GetReserveStat(req.body.monthStr, totalGoal, req.body.entTotalGoal, req.body.singleTotalGoal, res);
    }

    //預約明細
    if (flag == 'detail') {
        let response = MorReserve.GetKsReservStatMain(req.body.monthStr, res);
    }

    //腸胃竟統計
    if (flag == 'glumain') {
        let response = MorReserve.GetReserveGluStat(req.body.monthStr, req.body.dayGoal, res);
    }

    //腸胃鏡明細
    if (flag == 'gludetail') {
        let response = MorReserve.GetGluReservStat(req.body.monthStr, res);
    }

    //心超統計
    if (flag == 'heartchomain') {
        let response = MorReserve.GetReserveHeartChoStat(req.body.monthStr, req.body.dayGoal, res);
    }

    //心超明細
    if (flag == 'heartchodetail') {
        let response = MorReserve.GetHeartChoReservStat(req.body.monthStr, res);
    }

});

app.post('/api/MorTransFile', function (req, res) {
    console.log('執行檢驗轉檔資料' + new Date() + req.socket.remoteAddress);
    var flag = req.body.flag;


    if (flag == 'list') {
        let response = MorReserve.GetExamtransfileList(req.body.sdate, req.body.sdate, res);
    }

    if (flag == 'exam') {
        var seldata = req.body.seldata;
        let response = MorReserve.Examtransfile(req.body.sdate, req.body.sdate, seldata, res);
    }
});


app.post('/api/MorTransFileList', function (req, res) {
    console.log('查詢檢驗轉檔名單BY日期' + new Date());
    if (flag == 'exam') {
        let response = MorReserve.Examtransfile(req.body.sdate, req.body.sdate, res);
    }
});

//查詢報告漏項
app.post('/api/MorChkloss', function (req, res) {
    console.log('查詢報告漏項' + new Date() + req.socket.remoteAddress);
    (async function () {
        let response = await MorReserve.ChkRptloss(req.body.sdate, req.body.edate);
        res.send(response);
    })()

});

//查詢歷年檢驗數據
app.get('/api/MorQryExam/getcat', function (req, res) {
    console.log('查詢Catgory' + new Date() + req.socket.remoteAddress);
    (async function () {
        let response = await MorReserve.QryCatgory();
        res.send(response);
    })()
});

app.post('/api/MorQryExam/getclient', function (req, res) {
    console.log('查詢Client' + new Date() + req.socket.remoteAddress);
    (async function () {
        let response = await MorReserve.QryClientdata(req.body.datestr, req.body.page, req.body.size);
        res.send(response);
    })()
});

app.post('/api/MorQryExam/getclientdDetail', function (req, res) {
    console.log('查詢Client明細資料' + new Date() + req.socket.remoteAddress);
    (async function () {
        let response = await MorReserve.QryClientDetail(req.body.id, req.body.cat);
        res.send(response);
    })()
});

app.post('/api/sendmormail', function (req, res) {

    let sender = req.body.sender;
    let recipients = req.body.recipients;
    let title = req.body.title;
    let content = req.body.content;
    let resp = MorReserve.Sendemail(sender, recipients, title, content, res);
    console.log('寄出EMAIL:' + recipients + ' ' + new Date());
});



app.get('/RedirectTpr', function (req, res) { //重導向URL DEMO
    res.redirect(301, 'https://www.yahoo.com.tw')

});

app.get('/api', function (req, res) {
    res.send('晨悅WEB API 站台運行中!<br>系統時間' + new Date())
});