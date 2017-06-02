App.controller('Login_ctrl', ['$scope', '$rootScope', '$http', '$window', '$cookies', 'util_SERVICE',

function ($scope, $rootScope, $http, $window, $cookies, US) {
    $scope.userId = "";
    $scope.password = "";
    $cookies.put('Islogin', "false");
    var url = US.url;

    $scope.checklogin = function () {

        if ($scope.userId != "" && $scope.password != "") {
            var data = { "sUserName": $scope.userId, "sPassword": $scope.password }

            var config = {
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded;charset=utf-8;'
                }
            }

            var parms = JSON.stringify(data);
            $http.post(url + "CheckValidUser", "sJsonInput=" + parms, config)
   .then(
       function (response) {
           // success callback
           console.log(response.data);
           if (response.data[0].Result == "SUCCESS" && response.data[0].Result !== undefined) {
               //$cookies.put('MenuInfo', JSON.stringify(response.data.MenuInfo));
               $cookies.put('UserName', $scope.userId);
               $cookies.put('Islogin', "true");
               window.location = "Dashboard.aspx";
           }
           else
               alert(response.data[0].DisplayMessage);
       },
       function (response) {
           // failure callback

       }
    );

        }
    }
} ]);