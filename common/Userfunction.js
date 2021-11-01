var TYPES = require('tedious').TYPES;
const sql = require('mssql');
const db = require('../config/db');
const pool = new sql.ConnectionPool(db.GetMsssqldb())
module.exports = {

    GetDateTimeSingleData: function (datestr, tp) { //取得有開診日數        
        let sqlstr = "SELECT * " +
            " FROM [ExamineServer3].[dbo].[DateTimeRange] " +
            " where Date =@p1 ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, datestr)
            .returnRowCount()
            .execute()
            .then(function (results) {
                return results;

            }).fail(function (err) {
                return err;
            });

    },

    GetOpendays: function (monthStr, tp) { //取得有開診日數        
        let sqlstr = "SELECT Date " +
            " FROM [ExamineServer3].[dbo].[DateBusinessQuota] " +
            " where CHARINDEX(@p1,Date) >0 group by Date having SUM([BusinessQuota])>0 ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, monthStr)
            .returnRowCount()
            .execute()
            .then(function (results) {
                return results;

            }).fail(function (err) {
                return err;
            });

    },
    GetOpendaysByCourseNo: function (monthStr, courseno, tp) { //取腸胃竟得有開診日數        
        let sqlstr = "SELECT Date " +
            " FROM [ExamineServer3].[dbo].[DateBusinessQuota] " +
            " where CHARINDEX(@p1,Date) >0  and CourseNo=@p2 group by Date having SUM([BusinessQuota])>0 ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, monthStr)
            .parameter('p2', TYPES.NVarChar, courseno)
            .returnRowCount()
            .execute()
            .then(function (results) {
                return results;

            }).fail(function (err) {
                return err;
            });

    },
    GetSumPeople: function (monthStr, courseNo, tp) { //取得預約人數總和

        let sqlstr = " SELECT SUM([Quantity]) as cnt " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where CHARINDEX(@p1,Date) >0  and CourseNo=@p2 ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, monthStr)
            .parameter('p2', TYPES.NVarChar, courseNo)
            .execute()
            .then(function (results) {
                return results[0].cnt;
            }).fail(function (err) {
                return err;
            });

    },
    GetReserveSum: function (datestr, courseno, tp) { //取得各科可預約/已預約人數           
        let sqlstr = "SELECT　 Date, " +
            " ISNULL((select SUM(Quantity) from DateCourseQuantity " +
            " where DateCourseQuantity.Date = DateBusinessQuota.Date and DateCourseQuantity.CourseNo =@p1 ),0)   as '已預約', " +
            " BusinessQuota　as '可預約' " +
            " FROM [dbo].[DateBusinessQuota] " +
            " where CHARINDEX(@p2,Date) >0   and  CourseNo =@p1 ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, courseno)
            .parameter('p2', TYPES.NVarChar, datestr)
            .execute()
            .then(function (results) {
                return results;
            }).fail(function (err) {
                return err;
            });
    },
    GetReserveGS: function (datestr, s1, s2, tp) { //取得胃鏡/腸鏡的人數

        let sqlstr = " (SELECT distinct  [Barcodestr],[Date],[Quantity] " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where Date=@p1 and CourseNo in ('" + s1 + "','" + s2 + "')) " +
            " EXCEPT " +
            " ((SELECT distinct  [Barcodestr],[Date] ,[Quantity] " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where Date=@p1 and CourseNo in ('04','06')) " +
            " INTERSECT " +
            " (SELECT distinct  [Barcodestr],[Date],[Quantity] " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where Date=@p1 and CourseNo in ('05','07'))) ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, datestr)
            .returnRowCount()
            .execute()
            .then(function (results) {
                return results;
            }).fail(function (err) {
                return err;
            });

    },
    GetReserveGSIntersect: function (datestr, tp) { //取得腸鏡胃鏡交集的人數 

        let sqlstr =
            " (SELECT distinct  [Barcodestr],[Date] ,[Quantity] " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where Date=@p1 and CourseNo in ('04','06')) " +
            " INTERSECT " +
            " (SELECT distinct  [Barcodestr],[Date],[Quantity] " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where Date=@p1 and CourseNo in ('05','07'))";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, datestr)
            .returnRowCount()
            .execute()
            .then(function (results) {
                return results;
            }).fail(function (err) {
                return err;
            });
    },
    GetReserveGSUnion: function (datestr, orderTime, tp) { //取得腸鏡胃鏡聯集的人數 
        let sqlstr = "SELECT count(*) as cnt  FROM( " +
            " SELECT distinct [Barcodestr],[Date] ,[Quantity] , " +
            " (select OrderTimeRange from CheckProjectNameList where CheckProjectNameList.BarCodeStr = DateCourseQuantity.Barcodestr  ) as Ordertime " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where Date=@p1 and CourseNo in ('04','06') " +
            " UNION " +
            " SELECT distinct [Barcodestr],[Date],[Quantity] , " +
            " (select OrderTimeRange from CheckProjectNameList where CheckProjectNameList.BarCodeStr = DateCourseQuantity.Barcodestr ) as Ordertime " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where  Date=@p1 and CourseNo in ('05','07')) as table1 " +
            " where Ordertime=@p2";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, datestr)
            .parameter('p2', TYPES.NVarChar, orderTime)
            //.returnRowCount()
            .execute()
            .then(function (results) {
                return results[0].cnt;
            }).fail(function (err) {
                return err;
            });
    },
    GetReserveDokuUnion: function (datestr, orderTime, tp) { //取得為肉毒聯集的人數 
        let sqlstr = "SELECT count(*) as cnt  FROM( " +
            " SELECT distinct [Barcodestr],[Date] ,[Quantity] , " +
            " (select OrderTimeRange from CheckProjectNameList where CheckProjectNameList.BarCodeStr = DateCourseQuantity.Barcodestr  ) as Ordertime " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where Date=@p1 and CourseNo in ('35') " +
            " UNION " +
            " SELECT distinct [Barcodestr],[Date],[Quantity] , " +
            " (select OrderTimeRange from CheckProjectNameList where CheckProjectNameList.BarCodeStr = DateCourseQuantity.Barcodestr ) as Ordertime " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where  Date=@p1 and CourseNo in ('35')) as table1 " +
            " where Ordertime=@p2";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, datestr)
            .parameter('p2', TYPES.NVarChar, orderTime)
            //.returnRowCount()
            .execute()
            .then(function (results) {
                return results[0].cnt;
            }).fail(function (err) {
                return err;
            });
    },
    GetReserveHeartCho: function (datestr, orderTime, tp) { //取得心超的人數 
        let sqlstr = "SELECT count(*) as cnt  FROM( " +
            " SELECT distinct [Barcodestr],[Date] ,[Quantity] , " +
            " (select OrderTimeRange from CheckProjectNameList where CheckProjectNameList.BarCodeStr = DateCourseQuantity.Barcodestr  ) as Ordertime " +
            " FROM [ExamineServer3].[dbo].[DateCourseQuantity] " +
            " where Date=@p1 and CourseNo in ('09') ) as table1 " +
            " where Ordertime=@p2";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, datestr)
            .parameter('p2', TYPES.NVarChar, orderTime)
            //.returnRowCount()
            .execute()
            .then(function (results) {
                return results[0].cnt;
            }).fail(function (err) {
                return err;
            });
    },
    GetVipData: function (datestr, courseno, tp) { //取得專案資料名稱+業務  

        let sqlstr = " SELECT ProjectName, " +
            " (select UserName from Account where UserID =CheckProject.SalesmanID) as salesName " +
            " FROM [ExamineServer3].[dbo].[CheckProject] " +
            " where ProjectNo =(select ProjectNo from CheckProjectNameList " +
            " where BarCodeStr=(select top(1) Barcodestr from DateCourseQuantity where Date=@p1 and CourseNo=@p2))";
        return tp.sql(sqlstr)
            .parameter('p2', TYPES.NVarChar, courseno)
            .parameter('p1', TYPES.NVarChar, datestr)
            .execute()
            .then(function (results) {
                return results;
            }).fail(function (err) {
                return err;
            });

    },
    GetLabChkno: function (labstr, tp) { //轉換LAB代碼  

        let sqlstr = " SELECT top(1) CheckNo,Name FROM [ExamineServer3].[dbo].[CheckItem] " +
            " WHERE [CheckNo] in " +
            " (SELECT CheckNo FROM  [ExamineServer3].[dbo].[CheckItem_Lab_checkno] WHERE [LAB1]=@p1 )";

        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, labstr)
            .execute()
            .then(function (results) {
                return results[0];
            }).fail(function (err) {
                return err;
            });

    },
    GetPatientsex: function (barcodestr, tp) { //取得性別

        let sqlstr = " SELECT top(1) sex　FROM " +
            " [ExamineServer3].[dbo].[[CheckProjectNameList] " +
            " WHERE [BarCodeStr]=@p1";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, barcodestr)
            .execute()
            .then(function (results) {
                return results[0];
            }).fail(function (err) {
                return err;
            });

    },
    GetRefdata: function (barcodestr, labstr, tp) { //參考值資料  

        let sqlstr = "SELECT top(1) " +
            " MLow,MHigh,WLow,WHigh,Reference,Year " +
            " FROM [ExamineServer3].[dbo].[CompareValue] " +
            " where Year = (SELECT  B.Year " +
            " FROM [ExamineServer3].[dbo].[CheckProjectNameList] A, " +
            " [ExamineServer3].[dbo].[CheckProject]  B " +
            " WHERE A.ProjectNo=B.ProjectNo　and  a.barcodestr =@p1)" +
            " and  CheckItemNo=@p2";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, barcodestr)
            .parameter('p2', TYPES.NVarChar, labstr)
            .execute()
            .then(function (results) {
                return results[0];
            }).fail(function (err) {
                return err;
            });

    },
    CheckProjHasData: function (barcodestr, itemno, tp) { //檢查porject資料是否重複 

        let sqlstr = "SELECT count(*) as cnt from CheckProjectDetail " +
            " where BarCodeStr=@p1 and ItemNo=@p2";

        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, barcodestr)
            .parameter('p2', TYPES.NVarChar, itemno)
            .execute()
            .then(function (results) {
                return results[0];
            }).fail(function (err) {
                return err;
            });
    },
    CheckCorrectDate: function (opdate, barcodestr, tp) { //檢查檢驗BARCODE是否過期

        let sqlstr = "SELECT SpecialCheckDate from CheckProjectNameList " +
            " where BarCodeStr=@p1";

        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, barcodestr)
            .execute()
            .then(function (results) {
                var rst = results[0].SpecialCheckDate;
                if (rst != "" || rst != null) {
                    rst = rst.replaceAll("/");
                    if (rst == opdate) { return true }
                } else {
                    return false
                }

            }).fail(function (err) {
                return err;
            });
    },
    CheckHasItemNo: function (barcodestr, itemno, tp) { //檢查該患者檢查項目是否存在

        let sqlstr = " SELECT count(*) as cnt FROM fn_GetPersonalItem(@p1) where ItemNo=@p2";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, barcodestr)
            .parameter('p2', TYPES.NVarChar, itemno)
            .execute()
            .then(function (results) {
                return results[0];
            }).fail(function (err) {
                return err;
            });
    },
    InsertExamData: function (data, tp) { //新增API寫入資料(如果正常可刪)
        let sqlstr = " INSERT INTO CheckProjectDetail " +
            " ( BarCodeStr,ItemNo,Value,Err,Reference,Source,HandlerID,LisHigh,LisLow,LisFlag )" +
            " VALUES " +
            " (@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10)";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, data.BarCodeStr)
            .parameter('p2', TYPES.NVarChar, data.ItemNo)
            .parameter('p3', TYPES.NVarChar, data.Value)
            .parameter('p4', TYPES.Bit, data.Err)
            .parameter('p5', TYPES.NVarChar, data.Reference)
            .parameter('p6', TYPES.NVarChar, data.Source)
            .parameter('p7', TYPES.NVarChar, data.HandlerID)
            .parameter('p8', TYPES.NVarChar, data.LisHigh)
            .parameter('p9', TYPES.NVarChar, data.LisLow)
            .parameter('p10', TYPES.NVarChar, data.LisFlag)
            .execute()
            .then(function (results) {
                return "ok";
            }).fail(function (err) {
                return err;
            });

    },
    UpdateExamData: function (data, tp) { //檢驗資料已存在就更新(如果正常可刪))
        let sqlstr = " UPDATE CheckProjectDetail SET " +
            "  BarCodeStr=@p1,ItemNo=@p2,Value=@p3,Err=@p4, " +
            "  Reference=@p5,Source=@p6,HandlerID=@p7, " +
            "  LisHigh=@p8,LisLow=@p9,LisFlag=@p10 " +
            " WHERE " +
            " BarCodeStr=@p1 and ItemNo=@p2 ";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, data.BarCodeStr)
            .parameter('p2', TYPES.NVarChar, data.ItemNo)
            .parameter('p3', TYPES.NVarChar, data.Value)
            .parameter('p4', TYPES.Bit, data.Err)
            .parameter('p5', TYPES.NVarChar, data.Reference)
            .parameter('p6', TYPES.NVarChar, data.Source)
            .parameter('p7', TYPES.NVarChar, data.HandlerID)
            .parameter('p8', TYPES.NVarChar, data.LisHigh)
            .parameter('p9', TYPES.NVarChar, data.LisLow)
            .parameter('p10', TYPES.NVarChar, data.LisFlag)
            .execute()
            .then(function (results) {
                return "ok";
            }).fail(function (err) {
                return err;
            });
    },
    SaveExamData: function () { //轉檔儲存
        return pool.connect().then(pool => {
            return pool.request()
                .output('ErrMsg', sql.NVarChar)
                .execute('prc_sp_InsertExamdetail')
        }).then(result => {
            return result;
        }).catch(err => {
            return err;
        })
    },

    SaveExamDataBulk: function (data) { //檢驗資料存黨(批次測試)     



        return pool.connect().then(pool => {

            const table = new sql.Table('#datatmp')
            table.create = true
            table.columns.add('BarCodeStr', sql.VarChar(10), { nullable: false })
            table.columns.add('ItemNo', sql.VarChar(10), { nullable: false })
            table.columns.add('ItemName', sql.NVarChar(50), { nullable: true })
            table.columns.add('Value', sql.NVarChar(50), { nullable: true })
            table.columns.add('Err', sql.NVarChar(50), { nullable: true })
            table.columns.add('Reference', sql.NVarChar(200), { nullable: true })
            table.columns.add('Source', sql.NVarChar(50), { nullable: true })
            table.columns.add('HandlerID', sql.NVarChar(50), { nullable: true })
            table.columns.add('LisHigh', sql.NVarChar(50), { nullable: true })
            table.columns.add('LisLow', sql.NVarChar(50), { nullable: true })
            table.columns.add('LisFlag', sql.NVarChar(50), { nullable: true })
            table.columns.add('resp', sql.NVarChar(50), { nullable: true })
            table.columns.add('op', sql.NVarChar(50), { nullable: true })
            table.columns.add('LAB1', sql.NVarChar(50), { nullable: true })

            for (let item of data) {
                //更新處理 
                table.rows.add(item.BarCodeStr, item.ItemNo, item.ItemName,
                    item.Value, item.Err, item.Reference, item.Source,
                    item.HandlerID, item.LisHigh, item.LisLow,
                    item.LisFlag, item.resp, item.op, item.LAB1
                )
            }
            return pool.request().bulk(table);
        }).then(result => {
            return result;
        }).catch(err => {
            return err;
        })

    },
    InsertExamLogData: function (sdate, edate, data, tp) { //API寫入檢驗LOG

        let sqlstr = " INSERT INTO Log_GetLabData " +
            " ( Date_S,Date_N,GetData,INPUT_DATE) " +
            " VALUES " +
            " (@p1,@p2,@p3,@p4)";
        return tp.sql(sqlstr)
            .parameter('p1', TYPES.NVarChar, sdate)
            .parameter('p2', TYPES.NVarChar, edate)
            .parameter('p3', TYPES.NVarChar, data)
            .parameter('p4', TYPES.Date, new Date())
            .execute()
            .then(function (results) {
                return "寫入轉檔LOG完成!";
            }).fail(function (err) {
                return err;
            });
    },
    TurnMinGuo: function (westdate) { //西元轉民國年
        var result = westdate.split('/');
        var year = result[0] - 1911;
        var month = result[1];
        var date = result[2];
        return year + month + date;
    }

}