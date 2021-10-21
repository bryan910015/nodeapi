const sql = require('mssql');
const config = {
    user: 'gbuser',
    password: 'gbsoft2019',
    server: 'HISLIS',
    database: 'ExamineServer3',
    parseJSON: true,
    pool: {
        max: 10,
        min: 1,
        idleTimeoutMillis: 30000
    }
}

const pool = new sql.ConnectionPool(config)


module.exports.doquery = function () {

    var cpool = await pool.connect();
    let result = cpool.request()
        .input('i_sdate', sql.NVarChar, params.sdate)
        .input('i_edate', sql.NVarChar, params.edate)
        .execute('prc_sp_query_transnamelist')
    let res = await result;
    return res.recordset;
};

module.exports.doinsert = async function (res) {
    //交易transaction
    var cpool = await pool.connect();
    var transaction = new sql.Transaction(cpool)
    return transaction.begin(err => {

        const request = new sql.Request(transaction)
        var sqlstr = "insert into [dbo].[TestIns] (str1,str2,str3,str4,str5) values " +
            "('3', '李三', '高雄縣', 'ZZ路300號', '04123')";

        request.query(sqlstr, (err, result) => {
            console.log(err)
            transaction.commit(err => {
                res.send('交易完成');
            })
        })
    })
};