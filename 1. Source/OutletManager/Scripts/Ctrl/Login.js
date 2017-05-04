App.controller('login', ['$scope', '$rootScope', '$http', '$window', '$cookies', 'util_SERVICE',

function ($scope, $rootScope, $http, $window, $cookies, US) {
    $scope.userId = "";
    $scope.password = "";
    $cookies.put('Islogin', "false");
    var url = US.url;

    $scope.checklogin = function () {

        var data = { "sUserName": $scope.userId, "sPassword": $scope.password }

        var config = {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded;charset=utf-8;'
            }
        }

        var parms = JSON.stringify(data);
        $http.post(url + "LoginValidation", "sJsonInput=" + parms, config)
   .then(
       function (response) {
           // success callback
           console.log(response.data);
           if (response.data[0].UserId != "" && response.data[0].UserId !== undefined) {
               //$cookies.put('MenuInfo', JSON.stringify(response.data.MenuInfo));
               $cookies.put('UserData', JSON.stringify(response.data));
               $cookies.put('Islogin', "true");
               window.location = "index.html";
           }
           else
               alert(response.data[0].Message);
       },
       function (response) {
           // failure callback

       }
    );

    }
} ]);