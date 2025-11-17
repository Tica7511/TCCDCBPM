<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage/MasterPage.master" AutoEventWireup="true" CodeFile="DocFieldManage.aspx.cs" Inherits="WebPage_DocFieldManage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <link rel="stylesheet" href="<%= ResolveUrl("~/css/OchiLayout.css") %>" />
    <script src="<%= ResolveUrl("~/page_js/DocFieldManage.js") %>" type="text/javascript"></script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <!--#include file="LeftMenu.html"-->
    <input type="hidden" id="h_Guid"/>
    <input type="hidden" id="h_parentGuid"/>
    <input type="hidden" id="h_sn"/>
    <div class="flex-grow-1">

        <div class="container-ochi w-large-ochi mt-2" id="pagecontent">

            <nav class="itribreadcrumb d-none d-md-block" aria-label="breadcrumb">
                <ol class="breadcrumb bg-transparent">
                    <li class="breadcrumb-item"><a href="index.aspx">首頁</a></li>
                    <li class="breadcrumb-item active" aria-current="page">公文管理</li>
                    <li class="breadcrumb-item active" aria-current="page">公文範本欄位管理</li>
                </ol>
            </nav>

            <div class="filetitlewrapper mt-1">
                <span class="filetitle">
                    <h2>公文範本欄位管理</h2>
                </span>
                <span class="itemright"></span>
            </div>

            <div class="card mt-4 mb-2 border-nonfocus">
                <div class="card-header bg-nonfocus fs-5">公文範本列表</div>
                <div class="card-body">
                    <table id="tablist" class="table small">
                        <thead class="border-bottom border-dark-subtle">
                            <tr>
                                <th>範本名稱</th>
                                <th>目前版本</th>
                                <th>所屬類別</th>
                                <th>欄位數</th>
                                <th>修改者</th>
                                <th>修改日期</th>
                                <th class="text-center">功能</th>
                            </tr>
                        </thead>
                        <tbody>
                        </tbody>
                    </table>
                </div>
            </div>
            <!-- card -->

        </div>
        <!-- container-ochi -->

    </div>
    <!-- flex-grow-1 -->

    <!-- Magnific Popup -->
    <div id="mp_field" class="magpopup magSizeM mfp-hide">
        <div class="magpopupTitle"><span id="mp_filename"></span> 欄位編輯</div>
        <div class="padding10ALL">
            <div class="twocol">
                <div class="left font-size5 ">
                    <i class="fa fa-chevron-circle-right IconCa" aria-hidden="true"></i>
                    <select id="mgsellist" class="inputex">
                    </select>
                </div>
                <div class="right">
                </div>
            </div><br />
            <div class="stripeMeB tbover">
                <table id="tablistField" border="0" cellspacing="0" cellpadding="0" width="100%">
                    <thead>
                        <tr>
                            <th nowrap="nowrap" align="center">欄位代碼</th>
                            <th nowrap="nowrap" align="center">欄位名稱</th>
                            <th nowrap="nowrap" align="center">欄位類別</th>
                            <th nowrap="nowrap" align="center" width="10%">是否刪除</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>
        <!-- padding10ALL -->
    </div>
    <!--magpopup -->
</asp:Content>

