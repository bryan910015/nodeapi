module.exports = {

    Getdb: function (sname) {
        //初始化
        let server = "";
        let dbname = "";
        let username = "";
        let passwd = "";

        //判斷

        if (sname == "mordb") {
            server = "HISLIS";
            dbname = "ExamineServer3";
            username = "gbuser";
            passwd = "gbsoft2019";
        } else {

        }

        let config = {
            server: server, //update me        
            authentication: {
                type: 'default',
                options: {
                    userName: username, //update me
                    password: passwd //update me
                }
            },
            options: {
                requestTimeout: 0,
                database: dbname,
                rowCollectionOnRequestCompletion: true
            }
        };
        return config;
    },
    Getpoolingdb: function (sname) {
        //初始化
        let server = "";
        let dbname = "";
        let username = "";
        let passwd = "";

        if (sname == "mordb") {
            server = "HISLIS";
            dbname = "ExamineServer3";
            username = "gbuser";
            passwd = "gbsoft2019";
        }
        let config = {
            "userName": username,
            "password": passwd,
            "server": server,
            "options": {
                "database": dbname,
                "encrypt": false,
                "packetSize": 16384
            }
        }
        return config;
    },
    GetMsssqldb: function () {
        let config = {
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

        return config;
    }


};