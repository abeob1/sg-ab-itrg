<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Dashboard.aspx.cs" Inherits="Pages_Dashboard"
    MasterPageFile="~/Pages/Site.master" %>

<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div class="row">
        <div class="col-lg-12">
            <h3 class="page-header">
                <i class="fa fa-laptop"></i>Dashboard</h3>
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i><a href="Dashboard.aspx">Home</a></li>
                <li><i class="fa fa-laptop"></i>Dashboard</li>
            </ol>
        </div>
    </div>
    <div class="row">
        <div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
            <div class="info-box blue-bg">
                <i class="fa fa-cloud-download"></i>
                <div class="count">
                    1</div>
                <div class="title">
                    <a href="http://www.imperialtreasure.com/about/story" style="color: #FFFFFF" target="_blank">
                        Our Story</a></div>
            </div>
            <!--/.info-box-->
        </div>
        <!--/.col-->
        <div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
            <div class="info-box brown-bg">
                <i class="fa fa-shopping-cart"></i>
                <div class="count">
                    2</div>
                <div class="title">
                    <a href="http://www.imperialtreasure.com/about/history" style="color: #FFFFFF" target="_blank">
                        History & Milestones</a></div>
            </div>
            <!--/.info-box-->
        </div>
        <!--/.col-->
        <div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
            <div class="info-box dark-bg">
                <i class="fa fa-thumbs-o-up"></i>
                <div class="count">
                    3</div>
                <div class="title">
                    <a href="http://www.imperialtreasure.com/about/accolades" style="color: #FFFFFF"
                        target="_blank">Accolades</a></div>
            </div>
            <!--/.info-box-->
        </div>
        <!--/.col-->
        <div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
            <div class="info-box green-bg">
                <i class="fa fa-cubes"></i>
                <div class="count">
                    4</div>
                <div class="title">
                    <a href="http://www.imperialtreasure.com/about/career" style="color: #FFFFFF" target="_blank">
                        Career</a></div>
            </div>
            <!--/.info-box-->
        </div>
        <!--/.col-->
    </div>
    <!--/.row-->
</asp:Content>
