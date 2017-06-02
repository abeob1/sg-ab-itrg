<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Users.aspx.cs" Inherits="Pages_Users"
    MasterPageFile="~/Pages/Site.master" %>

<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <script src="../Scripts/Ctrl/User_ctrl.js" type="text/javascript"></script>
    <div ng-app="myApp" ng-controller="User_ctrl">
        <div class="row">
            <div class="col-lg-12">
                <h3 class="page-header">
                    <i class="fa fa-laptop"></i>User Management</h3>
                <ol class="breadcrumb">
                    <li><i class="fa fa-home"></i><a href="Dashboard.aspx">Home</a></li>
                    <li><i class="fa fa-laptop"></i>User Management</li>
                </ol>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-6">
                <input type="button" class="btn btn-primary" value="Create User" ng-click="CreateUserPopup();" />
            </div>
            <div class="col-lg-6" style="padding: 1% 0% 0% 41%;">
                <strong>User List : {{UserList.length}}</strong></div>
        </div>
        <br />
        <div role="tabpanel" class="tab-pane fade in" id="Addon" aria-labelledby="Addon-tab">
            <span class="top-buffer"></span>
            <table class="table table-bordered">
                <thead style="background-color: #2E3B46;">
                    <tr>
                        <th>
                            #
                        </th>
                        <th>
                            User Code
                        </th>
                        <th>
                            User Name
                        </th>
                        <th>
                            Default Entity
                        </th>
                        <th>
                            Default Branch Code
                        </th>
                        <th>
                            Default Dept Code
                        </th>
                        <th>
                            Password
                        </th>
                        <th>
                            Locked
                        </th>
                        <th>
                            Default Approval Level
                        </th>
                        <th>
                            Approval Scope
                        </th>
                        <th>
                            Language
                        </th>
                        <th colspan="2">
                            Action
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <tr ng-repeat="d in UserList">
                        <td>
                            {{$index+1}}
                        </td>
                        <td>
                            {{d.USERCODE}}
                        </td>
                        <td>
                            {{d.USERNAME}}
                        </td>
                        <td>
                            {{d.DEFAULTENTITY}}
                        </td>
                        <td>
                            {{d.DEFAULTBRANCHCODE}}
                        </td>
                        <td>
                            {{d.DEFAULTDEPTCODE}}
                        </td>
                        <td>
                            {{d.PASSWORD}}
                        </td>
                        <td>
                            {{d.LOCKED}}
                        </td>
                        <td>
                            {{d.DEFAULTAPPROVALLEVEL}}
                        </td>
                        <td>
                            {{d.APPROVALSCOPE}}
                        </td>
                        <td>
                            {{d.LANGUAGE}}
                        </td>
                        <td>
                            <button class="btn-info btn" ng-click="editUser(d);">
                                <i class="fa fa-pencil" aria-hidden="true"></i>Edit</button>
                            &nbsp;
                            <button class="btn-danger btn" ng-click="DeleteDriver(d);" ng-confirm-click="Are you SURE you want to Delete?">
                                <i class="fa fa-ban" aria-hidden="true"></i>Delete</button>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        <div id="NewUser" class="modal fade" role="dialog">
            <div class="modal-dialog">
                <!-- Modal content-->
                <div class="modal-content">
                    <div class="modal-header">
                        <button type="button" class="close" data-dismiss="modal">
                            &times;</button>
                        <h4 class="modal-title">
                            <i class="fa fa-truck" aria-hidden="true"></i>Add New User Details</h4>
                    </div>
                    <div class="modal-body">
                        <table width="100%" border="0" cellspacing="5" class="table-striped table" cellpadding="5">
                            <tr>
                                <td>
                                    User Name
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    <input type="text" ng-model="newdriver.UserName" id="UserName" class="form-control" />
                                    <%-- <select class="form-control" ng-options="item.DriverName as item.DriverName for item in PoPDriverList track by item.DriverId"
                                        ng-model="newdriver.DriverName">
                                        <option value="">-- select --</option>
                                    </select>--%>
                                </td>
                                <td>
                                    Default Entity
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    <input type="text" ng-model="newdriver.DefaultEntity" id="Entity" class="form-control" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Default Branch Code
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    <input type="Password" ng-model="newdriver.DefaultBranchCode" id="BranchCode" class="form-control" />
                                </td>
                                <td>
                                    Default Dept Code
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    <input type="text" ng-model="newdriver.DefaultDeptCode" id="DeptCode" class="form-control" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Password
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    <input type="text" ng-model="newdriver.Password" id="Password" class="form-control" />
                                </td>
                                <td>
                                    Locked
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    <input type="text" ng-model="newdriver.Locked" id="Locked" class="form-control" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Default Approval Level
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    <input type="text" ng-model="newdriver.DefaultAppLevel" id="AppLevel" class="form-control" />
                                </td>
                                <td>
                                    Approval Scope
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    <input type="text" ng-model="newdriver.ApprovalScope" id="ApprovalScope" class="form-control" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Language
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    <input type="text" ng-model="newdriver.Language" id="Lanuage" class="form-control" />
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div class="modal-footer">
                        <button class="btn-success btn" ng-click="SaveDriver();">
                            Save</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
</asp:Content>
