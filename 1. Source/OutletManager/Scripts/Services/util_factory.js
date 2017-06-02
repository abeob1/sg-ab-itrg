App.service('util_SERVICE', ['$http', '$window', '$cookieStore', '$rootScope', function ($http, $window, $cookie, $rootScope) {
    var urlsd = window.location.href.split("/");
    this.url = "http://119.73.138.58:85/Master.asmx/";
    this.Host = "http://119.73.138.58:85/";
    this.DBName = "OUTLETMANAGER_DEV";
    this.config = {
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded;charset=utf-8;'
        }
    }

    this.configgoogle = {
        headers: {
            'Content-Type': 'application/json charset=utf-8;',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET,PUT,POST,DELETE,OPTIONS'

        }
    }

    this.islogin = function () {
        if ($cookie.get('Islogin') == false || $cookie.get('Islogin') === undefined) {
            window.location = "login.html";
        }
    }

    //GetUserInfo
    this.GetUserInfo = function (company) {
        var promise = $http.post(this.url + "GetUserInfo", "sCompany=" + company, this.config)
   .success(function (response) { if (response.returnStatus == 1) { return response; } else { return false; } });
        return promise;
    };


} ]);
