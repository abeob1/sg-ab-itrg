﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Login.aspx.cs" Inherits="Pages_Login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <!-- Styles -->
    <link href="../CSS/Style.css" rel="stylesheet" type="text/css" />
    <link href="../CSS/LoginPage.css" rel="stylesheet" type="text/css" />
    <link href="../CSS/font-awesome.css" rel="stylesheet" type='text/css' />
    <link href="../CSS/bootstrap.css" rel='stylesheet' type='text/css' />
    <link href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css"
        rel="stylesheet" />
    <link href='//fonts.googleapis.com/css?family=Roboto+Condensed:400,300,300italic,400italic,700,700italic'
        rel='stylesheet' type='text/css' />
    <!-- Styles -->
    <!--Angular JS -->
    <script src="../Scripts/libs/jquery-1.11.1.min.js" type="text/javascript"></script>
    <script src="../Scripts/libs/angular.js" type="text/javascript"></script>
    <script src="../Scripts/libs/angular-route.js" type="text/javascript"></script>
    <script src="../Scripts/libs/angular-ui-router.js" type="text/javascript"></script>
    <script src="../Scripts/libs/ocLazyLoad.js" type="text/javascript"></script>
    <script src="../Scripts/libs/ui-bootstrap.js" type="text/javascript"></script>
    <script src="../Scripts/libs/ui-bootstrap-tpls-0.9.0.js" type="text/javascript"></script>
    <script src="../Scripts/libs/angular-cookies.js" type="text/javascript"></script>
    <!-- angular js short cut keys-->
    <script src="../Scripts/libs/hotkeys.js" type="text/javascript"></script>
    <!--Angular JS -->
    <!-- User Defined JS-->
    <script src="../Scripts/config/app.js" type="text/javascript"></script>
    <script src="../Scripts/Services/util_factory.js" type="text/javascript"></script>
    <script src="../Scripts/js/Login.js" type="text/javascript"></script>
    <script src="../Scripts/Ctrl/Login_ctrl.js" type="text/javascript"></script>
    <!-- User Defined JS-->
    <title>Login Page</title>
</head>
<body ng-app="myApp" ng-controller="Login_ctrl">
    <div class="main">
        <div class="container">
            <center>
                <div class="middle">
                    <div id="login">
                        <form action="javascript:void(0);" method="get">
                        <fieldset class="clearfix">
                            <p>
                                <span class="fa fa-user"></span>
                                <input type="text" ng-model="userId" placeholder="Username" required/></p>
                            <!-- JS because of IE support; better: placeholder="Username"
-->
                            <p>
                                <span class="fa fa-lock"></span>
                                <input type="password" ng-model="password" placeholder="Password" required /></p>
                            <div>
                                <span style="width: 100%; text-align: right;
                                        display: inline-block;">
                                        <input type="submit" value="Sign In" ng-click="checklogin();" /></span>
                            </div>
                        </fieldset>
                        <div class="clearfix">
                        </div>
                        </form>
                        <div class="clearfix">
                        </div>
                    </div>
                    <!--
end login -->
                    <div class="logo">
                        <img id="logo" alt="harrys" runat="server" src="../Images/logo.png" />
                        <div class="clearfix">
                        </div>
                    </div>
                </div>
            </center>
        </div>
    </div>
</body>
</html>

