App.controller('User_ctrl', ['$scope', '$rootScope', '$http', '$window', '$cookies', 'util_SERVICE',

function ($scope, $rootScope, $http, $window, $cookies, US) {
    $scope.DBName = "OUTLETMANAGER_DEV";
    $cookies.put('Islogin', "true");
    var url = US.url;

    $scope.GetUserInfo = function () {
        US.GetUserInfo($scope.DBName).then(function (response) {
            console.log(response);
            $scope.UserList = response.data;
        });

    }

    $scope.CreateUserPopup= function () {
        $('#NewUser').modal('show');
        //$rootScope.getDriverNameList();
    }

    $scope.GetUserInfo();
} ]);