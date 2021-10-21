const Connection = require('tedious').Connection;
const Request = require('tedious').Request;
const rq = require('request');
const db = require('../config/db');
const TYPES = require('tedious').TYPES;
const usrfunc = require('./Userfunction');
const nodemailer = require('nodemailer');
var dateFormat = require('dateformat');
const tp = require('tedious-promises');
const ConnectionPool = require('tedious-connection-pool');


var poolConfig = {
    min: 2,
    max: 5,
    log: false
};
var pool = new ConnectionPool(poolConfig, db.Getpoolingdb("mordb"));
tp.setConnectionPool(pool); // global
module.exports = {

    //取得預約資料
    GetReserveitem: function (sdate, edate, res) {

        let sqlstr = "select TOP(100) ROW_NUMBER() OVER(ORDER BY Date) AS ROWID, * from DateTimeRange where CONVERT(Date, Date) between @p1 and @p2 ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.VarChar, sdate)
            .parameter('p2', TYPES.VarChar, edate)
            .execute()
            .then(function (results) {
                var obj = [];
                //for                
                for (var i in results) {
                    var a = {};
                    var b = results[i];
                    a.rowid = b.ROWID;
                    a.redate = b.Date;
                    a.weekstr = getWeekStr(new Date(b.Date).getDay())
                    var timesplit = b.TimeRange.split(',');
                    for (var k in timesplit) {

                        if (timesplit[k].indexOf('07:15') >= 0) {
                            a.time15 = timesplit[k].split('@')[1];
                        }


                        if (timesplit[k].indexOf('07:30') >= 0) {
                            a.time1 = timesplit[k].split('@')[1];
                        }

                        if (timesplit[k].indexOf('08:00') >= 0) {
                            a.time2 = timesplit[k].split('@')[1];
                        }

                        if (timesplit[k].indexOf('08:30') >= 0) {
                            a.time3 = timesplit[k].split('@')[1];
                        }

                        if (timesplit[k].indexOf('09:00') >= 0) {
                            a.time4 = timesplit[k].split('@')[1];
                        }

                        if (timesplit[k].indexOf('09:30') >= 0) {
                            a.time5 = timesplit[k].split('@')[1];
                        }

                        if (timesplit[k].indexOf('10:00') >= 0) {
                            a.time6 = timesplit[k].split('@')[1];
                        }

                        if (timesplit[k].indexOf('10:30') >= 0) {
                            a.time7 = timesplit[k].split('@')[1];
                        }
                    }
                    obj.push(a);
                }
                //=======                                 
                return obj

            }).fail(function (err) {
                return err;
            });

    },
    //更新預約資料
    UpdateReserveitem: function (jsdata, res) {

        (async function () {

            var b = JSON.parse(jsdata);
            var keyDate = dateFormat(b.redate, "yyyy/mm/dd")
            var IsExist = await usrfunc.GetDateTimeSingleData(keyDate, tp);

            var timerangstr = "";

            if (b.time15 >= 0) {
                timerangstr = timerangstr + "07:15@" + b.time15 + ",";
            }


            if (b.time1 >= 0) {
                timerangstr = timerangstr + "07:30@" + b.time1 + ",";
            }

            if (b.time2 >= 0) {
                timerangstr = timerangstr + "08:00@" + b.time2 + ",";
            }

            if (b.time3 >= 0) {
                timerangstr = timerangstr + "08:30@" + b.time3 + ",";
            }

            if (b.time4 >= 0) {
                timerangstr = timerangstr + "09:00@" + b.time4 + ",";
            }

            if (b.time5 >= 0) {
                timerangstr = timerangstr + "09:30@" + b.time5 + ",";
            }

            if (b.time6 >= 0) {
                timerangstr = timerangstr + "10:00@" + b.time6 + ",";
            }

            if (b.time7 >= 0) {
                timerangstr = timerangstr + "10:30@" + b.time7;
            }

            if (IsExist == 0) {

                let sqlstr = " INSERT INTO DateTimeRange (Date,TimeRange) " + " values (@p1,@p2) ";
                return tp.sql(sqlstr)
                    .parameter('p1', TYPES.VarChar, keyDate)
                    .parameter('p2', TYPES.VarChar, timerangstr)
                    .returnRowCount()
                    .execute()
                    .then(function (results) {
                        return results;

                    }).fail(function (err) {
                        return err;
                    });

            } else {

                sqlstr = " update DateTimeRange set Date=@p1,TimeRange=@p2 where Date=@p1 ";
                return tp.sql(sqlstr)
                    .parameter('p1', TYPES.VarChar, keyDate)
                    .parameter('p2', TYPES.VarChar, timerangstr)
                    .returnRowCount()
                    .execute()
                    .then(function (results) {
                        return "OK";
                    }).fail(function (err) {
                        return err;
                    });
            }

        })()

    },
    //刪除預約資料 暫時不使用
    DelReserveitem: function (redate, res) {

        let sqlstr = " delete from  DateTimeRange where CONVERT(Date, Date)=@p1 ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.VarChar, redate)
            .returnRowCount()
            .execute()
            .then(function (results) {
                return 'OK';

            }).fail(function (err) {
                return err;
            });


    },
    //取得科室資料對應   
    GetCourseData: function () {
        let sqlstr = "SELECT CourseNo,CourseName FROM [ExamineServer3].[dbo].[Course]";
        return tp.sql(sqlstr)
            .execute()
            .then(function (results) {
                return results;
            }).fail(function (err) {
                return err;
            });
    },
    //取得科室預約資料
    GetKsReserveitem: function (sdate, edate, course, res) {

        let sqlstr = " SELECT TOP(100)  ROW_NUMBER() OVER(ORDER BY Date) AS ROWID,Date,BusinessNo, CourseNo," +
            " (select CourseName from Course where Course.CourseNo = DateBusinessQuota.CourseNo ) as CourseName, " +
            " ISNULL((select SUM(Quantity) from DateCourseQuantity where DateCourseQuantity.Date = DateBusinessQuota.Date and DateCourseQuantity.CourseNo = DateBusinessQuota.CourseNo ),0)   as Quantity, " +
            " BusinessQuota FROM [dbo].[DateBusinessQuota] ";
        if (course == '') {
            sqlstr = sqlstr + 'where Date between @p1 and @p2 order by ROWID';
            return tp.sql(sqlstr)
                .parameter('p1', TYPES.VarChar, sdate)
                .parameter('p2', TYPES.VarChar, edate)
                .execute()
                .then(function (results) {

                    var obj = [];
                    for (var i in results) {
                        var a = {};
                        var b = results[i];
                        a.rowid = b.ROWID
                        a.redate = b.Date; //日期
                        a.courseno = b.CourseNo;
                        a.weekstr = getWeekStr(new Date(b.Date).getDay()) //星期
                        a.businessNo = b.BusinessNo; //業務單位
                        a.courseName = b.CourseName; //科室
                        a.quantity = b.Quantity; //約
                        a.businessQuota = b.BusinessQuota; //總        
                        obj.push(a);
                    }
                    res.json(obj); //回傳結果            

                }).fail(function (err) {
                    return err;
                });

        } else {
            sqlstr = sqlstr + ' where CourseNo= @p3 and  Date between @p1 and @p2 order by ROWID';
            return tp.sql(sqlstr)
                .parameter('p1', TYPES.VarChar, sdate)
                .parameter('p2', TYPES.VarChar, edate)
                .parameter('p3', TYPES.VarChar, course)
                .execute()
                .then(function (results) {
                    var obj = [];
                    for (var i in results) {
                        var a = {};
                        var b = results[i];
                        a.rowid = b.ROWID
                        a.redate = b.Date; //日期
                        a.courseno = b.CourseNo;
                        a.weekstr = getWeekStr(new Date(b.Date).getDay()) //星期
                        a.businessNo = b.BusinessNo; //業務單位
                        a.courseName = b.CourseName; //科室
                        a.quantity = b.Quantity; //約
                        a.businessQuota = b.BusinessQuota; //總        
                        obj.push(a);
                    }
                    res.json(obj); //回傳結果 
                }).fail(function (err) {
                    return err;
                });
        }

    },
    //更新科室資料
    UpdateKsReserveitem: function (jsdata, res) {

        var connection = new Connection(db.Getdb('mordb'));
        connection.on('connect', function (err) {
            if (!err) {

                var b = JSON.parse(jsdata);
                sqlstr = " update DateBusinessQuota set BusinessQuota=@p1 " +
                    " where Date=@p2 and BusinessNo=@p3 and CourseNo=@p4  ";
                console.log(b.businessQuota);
                UpdatecontKs(b.businessQuota, b.redate, b.businessNo, b.courseno, res, sqlstr, connection);

            }
        });
    },
    //取得預約統計資料
    GetReserveStat: function (monthStr, totalGoal, entTotalGoal, singleTotalGoal, res) {

        (async function () {

            var totalRst = {};
            var entRst = {};
            var singleRst = {};
            var obj = [];
            var opDays = await usrfunc.GetOpendays(monthStr, tp); //開診天數

            var totalSum = await usrfunc.GetSumPeople(monthStr, '01', tp);
            totalRst.genere = '總';
            totalRst.opDays = opDays; //開診天數
            totalRst.dayGoal = totalGoal; //總-日目標
            totalRst.totalSum = totalSum; //總-已預約
            totalRst.totalAvg = Math.round(totalRst.totalSum / opDays); //總 - 累計日均(已預約/開診天數)
            totalRst.totalMonGoal = parseInt(totalRst.dayGoal) * parseInt(opDays); //總 - 月目標(日目標*開診天數)
            totalRst.rateGoal = CntPercent(totalRst.totalSum, totalRst.totalMonGoal); //總 - 達成率(已預約/月目標)

            var entTotalSum = await usrfunc.GetSumPeople(monthStr, '29', tp);
            entRst.genere = '企檢';
            entRst.opDays = opDays; //開診天數
            entRst.dayGoal = entTotalGoal; //企-日目標
            entRst.totalSum = entTotalSum; //企-已預約
            entRst.totalAvg = Math.round(entRst.totalSum / opDays); //企 - 累計日均(已預約/開診天數)
            entRst.totalMonGoal = entTotalGoal * opDays; //企 - 月目標(日目標*開診天數)
            entRst.rateGoal = CntPercent(entRst.totalSum, entRst.totalMonGoal); //達成率(已預約/月目標)

            var singleTotalSum = await usrfunc.GetSumPeople(monthStr, '30', tp);
            singleRst.genere = '個檢';
            singleRst.opDays = opDays; //開診天數
            singleRst.dayGoal = singleTotalGoal; //個-日目標
            singleRst.totalSum = singleTotalSum //個-已預約
            singleRst.totalAvg = Math.round(singleRst.totalSum / opDays); //個 - 累計日均(已預約/開診天數)
            singleRst.totalMonGoal = singleTotalGoal * opDays; //個 - 月目標(日目標*開診天數)
            singleRst.rateGoal = CntPercent(singleRst.totalSum, singleRst.totalMonGoal); //達成率(已預約/月目標)

            obj.push(totalRst);
            obj.push(entRst);
            obj.push(singleRst);
            res.json(obj); //回傳結果 
        })()


    },
    //姓名預約統計資料明細
    GetKsReservStatMain: function (monthStr, res) {

        (async function () {
            var A01 = await usrfunc.GetReserveSum(monthStr, '01', tp); //到院
            var A29 = await usrfunc.GetReserveSum(monthStr, '29', tp); //企檢
            var A30 = await usrfunc.GetReserveSum(monthStr, '30', tp); //個檢
            var A39 = await usrfunc.GetReserveSum(monthStr, '39', tp); //勞檢
            var A32 = await usrfunc.GetReserveSum(monthStr, '32', tp); //補檢
            var A08 = await usrfunc.GetReserveSum(monthStr, '08', tp); //乙狀結腸
            var A35 = await usrfunc.GetReserveSum(monthStr, '35', tp); //胃肉毒
            var A09 = await usrfunc.GetReserveSum(monthStr, '09', tp); //心超
            var A11 = await usrfunc.GetReserveSum(monthStr, '11', tp); //抹片
            var A24 = await usrfunc.GetReserveSum(monthStr, '24', tp); //GABUS
            var A31 = await usrfunc.GetReserveSum(monthStr, '31', tp); //聽力
            var A22 = await usrfunc.GetReserveSum(monthStr, '22', tp); //vip
            var A23 = await usrfunc.GetReserveSum(monthStr, '23', tp); //vipsofa
            var A12 = await usrfunc.GetReserveSum(monthStr, '12', tp); //糞便潛血
            var A13 = await usrfunc.GetReserveSum(monthStr, '13', tp); //一般超音
            var A15 = await usrfunc.GetReserveSum(monthStr, '15', tp); //DXA
            var A16 = await usrfunc.GetReserveSum(monthStr, '16', tp); //未來10年
            var A18 = await usrfunc.GetReserveSum(monthStr, '18', tp); //HRV
            var A19 = await usrfunc.GetReserveSum(monthStr, '19', tp); //肺功能
            var A20 = await usrfunc.GetReserveSum(monthStr, '20', tp); //腳踝古密
            var A21 = await usrfunc.GetReserveSum(monthStr, '21', tp); //眼底攝影
            var A25 = await usrfunc.GetReserveSum(monthStr, '25', tp); //APG
            var A26 = await usrfunc.GetReserveSum(monthStr, '26', tp); //TOMO
            var A27 = await usrfunc.GetReserveSum(monthStr, '27', tp); //XRAY
            //內視鏡
            var nayshi = await usrfunc.GetReserveSum(monthStr, '33', tp);
            var obj = [];

            for (var i in A01) {
                var a = {};

                //到院
                a.datestr = A01[i].Date;
                a.weekstr = getWeekStr(new Date(A01[i].Date).getDay()) //星期
                a.arrival = A01[i].已預約;
                a.arrivalR = A01[i].可預約;
                //企檢
                a.enterprise = A29[i].已預約;
                a.entpriseR = A29[i].可預約;
                //個檢
                a.single = A30[i].已預約;
                a.singleR = A30[i].可預約;
                //勞檢
                a.labor = A39[i].已預約;
                a.laborR = A39[i].可預約;
                //補檢
                a.makeup = A32[i].已預約;
                //乙狀
                a.Btype = parseInt(A08[i].已預約);
                //腸鏡
                a.gutscnt = await usrfunc.GetReserveGS(a.datestr, '05', '07', tp); //單腸鏡
                //胃鏡
                a.stomcnt = await usrfunc.GetReserveGS(a.datestr, '04', '06', tp); //單胃鏡
                //胃肉毒
                a.stomchdoku = A35[i].已預約;
                //腸胃人數
                a.gutstomcnt = await usrfunc.GetReserveGSIntersect(a.datestr, tp);
                //腸胃鏡總人數
                a.gutstomtotalman = a.Btype + a.gutscnt + a.stomcnt + a.gutstomcnt + a.stomchdoku;
                //腸胃鏡總人數已預約(內視鏡)
                a.gutstomtotal = nayshi[i].已預約
                ///腸胃鏡總人數可預約(內視鏡)
                a.gutstomtotalR = A08[i].可預約;
                //心超
                a.heartcho = A09[i].已預約;
                a.heartchoR = A09[i].可預約;
                //抹片
                a.moppem = A11[i].已預約;
                a.moppemR = A11[i].可預約;
                //GEABUS
                a.geabus = A24[i].已預約;
                a.geabusR = A24[i].可預約
                //聽力
                a.listen = A31[i].已預約
                a.listenR = A31[i].可預約
                //VIP
                let vip = await usrfunc.GetVipData(a.datestr, '22', tp)
                let vipsofa = await usrfunc.GetVipData(a.datestr, '23', tp)
                if (vip.length != 0) {
                    a.vip = vip[0].ProjectName + vip[0].salesName;
                }

                if (vipsofa.length != 0) {
                    a.vipsofa = vipsofa[0].ProjectName + vipsofa[0].salesName;
                }
                a.vipcnt = A22[i].已預約;
                a.vipsofacnt = A23[i].已預約;
                //糞便潛血
                a.fblood = A12[i].已預約;
                //一般超
                a.norcho = A13[i].已預約;
                //DXA
                a.dxa = A15[i].已預約;
                //未來10年
                a.tenyrs = A16[i].已預約;
                //HRV
                a.hrv = A18[i].已預約;
                //肺功能
                a.lungfun = A19[i].已預約;
                //腳踝古密
                a.feetwhy = A20[i].已預約;
                //眼底攝影
                a.eyebtn = A21[i].已預約;
                //APG
                a.apg = A25[i].已預約;
                //TOMO
                a.tomo = A26[i].已預約;
                //xRAY
                a.xray = A27[i].已預約;
                obj.push(a);
            }

            res.json(obj); //回傳結果

        })()

    },
    //取得腸胃鏡預約統計資料
    GetReserveGluStat: function (monthStr, dayGoal, res) {

        (async function () {

            var totalRst = {};
            var obj = [];
            var opDays = await usrfunc.GetOpendaysByCourseNo(monthStr, '33', tp); //開診天數
            var totalSum = await usrfunc.GetSumPeople(monthStr, '33', tp);
            totalRst.opDays = opDays; //開診天數
            totalRst.dayGoal = dayGoal; //日目標
            totalRst.totalSum = totalSum; //已預約
            totalRst.totalAvg = Math.round(totalRst.totalSum / opDays); //累計日均(已預約/開診天數)
            totalRst.totalMonGoal = dayGoal * opDays; //月目標(日目標*開診天數)
            totalRst.rateGoal = CntPercent(totalRst.totalSum, totalRst.totalMonGoal); //達成率(已預約/月目標) 
            obj.push(totalRst);
            res.json(obj); //回傳結果 
        })()


    },
    //腸胃鏡時段統計
    GetGluReservStat: function (monthStr, res) {
        (async function () {
            var obj = [];
            //內視鏡
            var nayshi = await usrfunc.GetReserveSum(monthStr, '33', tp);

            for (var i in nayshi) {
                var a = {};
                var b = nayshi[i];
                a.datestr = b.Date;
                a.weekstr = getWeekStr(new Date(b.Date).getDay()) //星期
                //內視鏡
                a.nayshi = b.已預約;
                a.nayshiR = b.可預約;
                a.nayshiPercent = CntPercent(a.nayshi, a.nayshiR);
                //腸胃竟
                a.glu0715 = await usrfunc.GetReserveGSUnion(b.Date, '07:15', tp);
                a.glu0715P = CntPercent(a.glu0715, a.nayshi);
                a.glu0730 = await usrfunc.GetReserveGSUnion(b.Date, '07:30', tp);
                a.glu0730P = CntPercent(a.glu0730, a.nayshi);
                a.glu0800 = await usrfunc.GetReserveGSUnion(b.Date, '08:00', tp);
                a.glu0800P = CntPercent(a.glu0800, a.nayshi);
                a.glu0830 = await usrfunc.GetReserveGSUnion(b.Date, '08:30', tp);
                a.glu0830P = CntPercent(a.glu0830, a.nayshi);
                //為肉毒 
                a.doku0715 = await usrfunc.GetReserveDokuUnion(b.Date, '07:15', tp);
                a.doku0715P = CntPercent(a.doku0715, a.nayshi);
                a.doku0730 = await usrfunc.GetReserveDokuUnion(b.Date, '07:30', tp);
                a.doku0730P = CntPercent(a.doku0730, a.nayshi);
                a.doku0800 = await usrfunc.GetReserveDokuUnion(b.Date, '08:00', tp);
                a.doku0800P = CntPercent(a.doku0800, a.nayshi);
                a.doku0830 = await usrfunc.GetReserveDokuUnion(b.Date, '08:30', tp);
                a.doku0830P = CntPercent(a.doku0830, a.nayshi);
                a.doku0900 = await usrfunc.GetReserveDokuUnion(b.Date, '09:00', tp);
                a.doku0900P = CntPercent(a.doku0900, a.nayshi);
                a.doku0930 = await usrfunc.GetReserveDokuUnion(b.Date, '09:30', tp);
                a.doku0930P = CntPercent(a.doku0930, a.nayshi);
                //總計
                a.totalG = a.glu0715 + a.glu0730 + a.glu0800 + a.glu0830 +
                    +a.doku0715 + a.doku0730 + a.doku0800 + a.doku0830 + a.doku0900 + a.doku0930;

                obj.push(a);
            }
            res.json(obj);
        })()

    },
    //取得心超預約統計資料
    GetReserveHeartChoStat: function (monthStr, dayGoal, res) {

        (async function () {

            var totalRst = {};
            var obj = [];
            var opDays = await usrfunc.GetOpendaysByCourseNo(monthStr, '09', tp); //開診天數
            var totalSum = await usrfunc.GetSumPeople(monthStr, '09', tp);
            totalRst.opDays = opDays; //開診天數
            totalRst.dayGoal = dayGoal; //日目標
            totalRst.totalSum = totalSum; //已預約
            totalRst.totalAvg = Math.round(totalRst.totalSum / opDays); //累計日均(已預約/開診天數)
            totalRst.totalMonGoal = dayGoal * opDays; //月目標(日目標*開診天數)
            totalRst.rateGoal = CntPercent(totalRst.totalSum, totalRst.totalMonGoal); //達成率(已預約/月目標) 
            obj.push(totalRst);
            res.json(obj); //回傳結果 
        })()
    },//心超時段統計
    GetHeartChoReservStat: function (monthStr, res) {
        (async function () {
            var obj = [];
            //心超
            var heartcho = await usrfunc.GetReserveSum(monthStr, '09', tp);

            for (var i in heartcho) {
                var a = {};
                var b = heartcho[i];
                a.datestr = b.Date;
                a.weekstr = getWeekStr(new Date(b.Date).getDay()) //星期
                //內視鏡
                a.heartcho = b.已預約;
                a.heartchoR = b.可預約;
                a.heartchoPercent = CntPercent(a.heartcho, a.heartchoR);

                a.heartcho0715 = await usrfunc.GetReserveHeartCho(b.Date, '07:15', tp);
                a.heartcho0715P = CntPercent(a.heartcho0715, a.heartcho);

                a.heartcho0730 = await usrfunc.GetReserveHeartCho(b.Date, '07:30', tp);
                a.heartcho0730P = CntPercent(a.heartcho0730, a.heartcho);

                a.heartcho0800 = await usrfunc.GetReserveHeartCho(b.Date, '08:00', tp);
                a.heartcho0800P = CntPercent(a.heartcho0800, a.heartcho);

                a.heartcho0830 = await usrfunc.GetReserveHeartCho(b.Date, '08:30', tp);
                a.heartcho0830P = CntPercent(a.heartcho0830, a.heartcho);

                a.heartcho0900 = await usrfunc.GetReserveHeartCho(b.Date, '09:00', tp);
                a.heartcho0900P = CntPercent(a.heartcho0900, a.heartcho);

                a.heartcho0930 = await usrfunc.GetReserveHeartCho(b.Date, '09:30', tp);
                a.heartcho0930P = CntPercent(a.heartcho0930, a.heartcho);

                a.heartcho1000 = await usrfunc.GetReserveHeartCho(b.Date, '10:00', tp);
                a.heartcho1000P = CntPercent(a.heartcho1000, a.heartcho);

                a.heartcho1030 = await usrfunc.GetReserveHeartCho(b.Date, '10:30', tp);
                a.heartcho1030P = CntPercent(a.heartcho1030, a.heartcho);

                obj.push(a);
            }
            res.json(obj);
        })()

    },
    //檢驗資料轉檔
    Examtransfile: function (sdate, edate, seldata, res) {
        (async function () {
            var rtObj = [];
            seldata = JSON.parse(seldata);
            for (var i in seldata) {
                var b = seldata[i];
                var BarCodeStr = b.HOS_NUM;
                var chkNo = "";
                if (b.LAB1 != "" || b.LAB1 != null) {
                    var chkNo = await usrfunc.GetLabChkno(b.LAB1.trim(), tp);
                }

                //轉換檢驗項目/名稱
                var ItemNo = "";
                var ItemName = "";
                if (typeof chkNo != 'undefined') {
                    ItemNo = chkNo.CheckNo;
                    ItemName = chkNo.Name;
                }

                var refData = await usrfunc.GetRefdata(b.HOS_NUM, ItemNo, tp);
                var Value = "";
                //有包含+-號才合併值寫入

                let input1p = b.INPUT1;
                if (b.INPUT1 == '-') {
                    input1p = '(-)';
                }
                if (b.INPUT1 == '+') {
                    input1p = '(+)';
                }

                if (b.SAY_INPUT1 != null) {

                    if (b.SAY_INPUT1.indexOf('+') > 0 || b.SAY_INPUT1.indexOf('-') > 0) {
                        Value = b.SAY_INPUT1 + input1p;
                    } else {
                        Value = input1p;
                    }
                } else {
                    Value = input1p;
                }

                // 只要有 NULL 或是 - 號 ERR值設為0
                var Err = "";
                if (b.SAY_INPUT1 == null || b.SAY_INPUT1.indexOf('-') > 0) {
                    Err = 0;
                } else {
                    Err = 1;
                }

                var Source = "API";
                var HandlerID = "API";
                var Reference = ""
                var LisHigh = "";
                var LisLow = "";
                var LisFlag = "";

                //參考值判斷性別

                if (typeof refData != 'undefined') {
                    Reference = refData.Reference;

                    LisFlag = b.STS_INPUT1;
                    if (b.STS_INPUT1 == null) {
                        LisFlag = "";
                    }
                    if (b.SEX == '1') {
                        LisHigh = refData.MHigh;
                        LisLow = refData.MLow;
                    }
                    if (b.SEX == '2') {
                        LisHigh = refData.WHigh;
                        LisLow = refData.WLow;
                    }
                }
                //======================
                //檢查資料是否已存在才寫入
                // var Hasdata = await usrfunc.CheckProjHasData(BarCodeStr, ItemNo, tp);
                //檢查項目是否存在於該患者檢查項目
                var hasItem = await usrfunc.CheckHasItemNo(BarCodeStr, ItemNo, tp);
                var data = {};
                data.Transdate = sdate;
                data.BarCodeStr = BarCodeStr;
                data.ItemNo = ItemNo;
                data.ItemName = ItemName;
                data.Value = Value;
                data.Err = Err;
                data.Reference = Reference;
                data.Source = Source;
                data.HandlerID = HandlerID
                data.LisHigh = LisHigh;
                data.LisLow = LisLow;
                data.LisFlag = LisFlag;
                data.resp = "";
                data.op = "";
                data.LAB1 = b.LAB1;

                var isCorrectDate = usrfunc.CheckCorrectDate(b.OP_DATE1, BarCodeStr, tp)

                //寫入資料(如果存在於檢查項目才寫入)
                //檢查檢驗BARCODE是否過期
                /*
               

                if (isCorrectDate) {


                    //if (Hasdata.cnt == 0) {
                    //不存在資料就新增
                    
                    var insResult = await usrfunc.InsertExamData(data, tp);
                    if (insResult == 'ok') {
                        data.resp = insResult;
                    } else {
                        data.resp = insResult.message;
                    }
                    


                    // data.op = '新增';
                    // }

                    // if (Hasdata.cnt > 0) {

                    //存在資料就更新
                    
                     var upResult = await usrfunc.UpdateExamData(data, tp);
                     if (upResult == 'ok') {
                         data.resp = upResult;
                     } else {
                         data.resp = upResult.message;
                     }
                     

                    //data.op = '更新';
                    //}

                }
                */


                if (hasItem.cnt == 0) {
                    data.op = data.op + ' ' + '多餘項';
                }

                if (data.ItemNo == '') {
                    data.op = data.op + ' ' + '無項目';
                }

                if (isCorrectDate == false) {
                    data.op = data.op + ' ' + '日期錯誤';
                }

                //列出異常資料                                
                rtObj.push(data);
                // }

            }
            await usrfunc.SaveExamDataBulk(rtObj) // 儲存到站存TABLE         
            let saveRst = await usrfunc.SaveExamData();
            //console.log(saveRst.output.ErrMsg + ' ' + saveRst.rowsAffected.toString());
            //寫入LOG擋
            var writelog = await usrfunc.InsertExamLogData(sdate, edate, JSON.stringify(seldata), tp);
            console.log(writelog);
            res.json(rtObj); //回傳結果                     

        })()

    },
    //檢驗資料轉檔
    GetExamtransfileList: function (sdate, edate, res) {
        let url = "http://210.244.79.219/leeapi/?A1=B471_043WKT&A2=" + sdate + "&A3=" + edate;
        rq.get(
            url, { json: { key: 'value' } },
            function (error, response, body) {
                res.send(response.body);
            }
        );
    },
    //查詢報告漏項
    ChkRptloss: function (sdate, edate) {
        let sqlstr = "SELECT * FROM [dbo].[fn_Query_Lost_item](@p1,@p2)";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, sdate)
            .parameter('p2', TYPES.NVarChar, edate)
            .execute()
            .then(function (results) {
                return results;
            }).fail(function (err) {
                return err;
            });
    },
    //查詢科室選項
    QryCatgory: function () {
        let sqlstr = "SELECT DISTINCT CheckCategory  FROM ExamineServer3.dbo.CheckItem " +
            " WHERE CheckCategory IS NOT NULL AND CheckCategory <>'' ORDER BY CheckCategory ";
        return tp.sql(sqlstr)
            .execute()
            .then(function (results) {
                return results;
            }).fail(function (err) {
                return err;
            });
    },
    //查詢客戶資料
    QryClientdata: function (datestr, pg, pgsize) {
        let sqlstr = "select count(1) over() as TOTAL_COUNT,* from ( " +
            " SELECT A.[BarCodeStr] , A.ID " +
            " ,[AnamnesisNo] ,A.[Name]  ,[Sex] " +
            " , CAST(YEAR(@p1) as int)  - CAST(SUBSTRING([Birthday],1,4) as int) age " +
            " , C.Name CUSTNAME,[SchemeExplain] " +
            " , ( SELECT STR( COUNT(*)) FROM [ExamineServer3].[dbo].[ClinicList] " +
            " WHERE BarcodeStr=A.BarCodeStr  AND CheckDate =@p1 AND IsComplete=1 AND IsCancel=0 ) +'/'+ " +
            " (SELECT LTRIM(STR(COUNT(*)))    FROM [ExamineServer3].[dbo].[ClinicList]  " +
            " WHERE BarcodeStr=A.BarCodeStr  AND CheckDate =@p1 ) completedu " +
            " FROM [ExamineServer3].[dbo].[CheckProjectNameList] A , " +
            " [ExamineServer3].[dbo].[CheckProject] B, " +
            " [ExamineServer3].[dbo].[Customer] C " +
            " WHERE A.ProjectNo=B.ProjectNo AND B.CustomerNo=C.CustomerNo AND  " +
            " A.SpecialCheckDate=@p1 ) as table1  " +
            " ORDER BY [AnamnesisNo]  OFFSET @p2 ROWS FETCH FIRST @p3 ROWS ONLY  ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, datestr)
            .parameter('p2', TYPES.Int, parseInt(pg - 1) * parseInt(pgsize))
            .parameter('p3', TYPES.Int, parseInt(pgsize))
            .execute()
            .then(function (results) {
                var rstdata = {};
                rstdata.page = parseInt(pg);
                rstdata.records = parseInt(results[0].TOTAL_COUNT);
                rstdata.total = Math.ceil((rstdata.records / parseInt(pgsize)));
                rstdata.data = results;
                return rstdata;
            }).fail(function (err) {
                return err;
            });
    },
    //查詢客戶明細
    QryClientDetail: function (id, cat) {
        let sqlstr = "SELECT C.CheckCategory 檢驗分類 ,C.CHECKNO 項目代號, C.NAME 項目名稱, " +
            " A.SpecialCheckDate 檢驗日期, B.Value 檢驗結果 , B.Reference 參考值 " +
            " FROM  CheckProjectNameList A, CheckProjectDetail B ,CheckItem C " +
            " WHERE A.BarCodeStr=B.BarCodeStr AND B.ItemNo=C.CheckNo AND A.ID =@p1 AND C.CheckCategory=@p2 " +
            " ORDER BY  C.CheckCategory,C.CHECKNO, C.NAME , SpecialCheckDate DESC ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, id)
            .parameter('p2', TYPES.NVarChar, cat)
            .execute()
            .then(function (results) {
                return results;
            }).fail(function (err) {
                return err;
            });
    },
    Sendemail: function (sender, receiver, subject, content, res) { //新功能使用SMTP寄信

        let transporter = nodemailer.createTransport({
            host: 'mail.mornjoy.com.tw',
            port: 25,
            secure: false,
            tls: { rejectUnauthorized: false },
            auth: {
                user: 'mornjoy',
                pass: 'Mj123456'
            },
            debug: true
        })

        transporter.sendMail(
            {
                from: sender,
                to: receiver,
                subject: subject,
                html: content,
                attachments: [
                    {
                        filename: 'brcode1.png',
                        path: './img/brcode1.png',
                        cid: 'brcode1'
                    },
                    {
                        filename: 'smp1.jpg',
                        path: './img/smp1.jpg',
                        cid: 'smp1'
                    }
                ]
            },
            function (err, info) {

                if (err) {
                    res.send('Unable to send email: ' + err + new Date());

                } else {
                    res.send('OK');
                }
            },
        );
    }
}


//科室的update
function UpdatecontKs(BusinessQuota, Date, BusinessNo, CourseNo, res, sqlstr, connection) {
    request = new Request(sqlstr, function (err, rowCount, rows) {
        (async function () {
            if (err) {
                console.log(err);
            } else {
                res.send("OK"); //回傳結果
            }
            connection.close();
        })()
    });
    request.addParameter('p1', TYPES.Int, BusinessQuota);
    request.addParameter('p2', TYPES.VarChar, Date);
    request.addParameter('p3', TYPES.VarChar, BusinessNo);
    request.addParameter('p4', TYPES.VarChar, CourseNo);
    connection.execSql(request);
}

//取得日期轉星期
function getWeekStr(wkday) {
    weekStr = "";
    if (wkday == 0) {
        weekStr = "星期日";
    } else if (wkday == 1) {
        weekStr = "星期一";
    } else if (wkday == 2) {
        weekStr = "星期二";
    } else if (wkday == 3) {
        weekStr = "星期三";
    } else if (wkday == 4) {
        weekStr = "星期四";
    } else if (wkday == 5) {
        weekStr = "星期五";
    } else if (wkday == 6) {
        weekStr = "星期六";
    } else {

    }

    return weekStr;
}

//自動對應全部JSON欄位
function OptJson(data) {
    var obj = [];
    data.forEach(function (item, index) {
        var rowObject = {};
        item.forEach(function (item2, index) {
            rowObject[item2.metadata.colName] = item2.value;
        });
        obj.push(rowObject);
    });
    return obj;
}

//計算百分比
function CntPercent(n1, n2) {
    if (n1 == '0' || n2 == '0') {
        return 0;
    } else {
        return Math.round((parseInt(n1) / parseInt(n2)) * 100);
    }
}