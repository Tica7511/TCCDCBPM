<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage/MasterPage.master" AutoEventWireup="true" CodeFile="DocPendingProcessing.aspx.cs" Inherits="WebPage_DocPendingProcessing" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script src="<%= ResolveUrl("~/page_js/DocPendingProcessing.js") %>" type="text/javascript"></script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <!--#include file="LeftMenu.html"-->
    <div class="flex-grow-1">
    
      <div class="container-ochi w-large-ochi mt-2" id="pagecontent">
    
        <nav class="itribreadcrumb d-none d-md-block" aria-label="breadcrumb">
          <ol class="breadcrumb bg-transparent">
            <li class="breadcrumb-item"><a href="index.aspx">首頁</a></li>
            <li class="breadcrumb-item active" aria-current="page">我的公文</li>
            <li class="breadcrumb-item active" aria-current="page">待處理公文</li>
          </ol>
        </nav>
    
        <div class="filetitlewrapper mt-1">
          <span class="filetitle"><h2>待處理公文</h2></span>
          <span class="itemright"></span>
        </div>


          <%--<div class="text-start mt-2">
              <a href="javascript:void(0);" class="btn btn-primary" id="assignsendOK" data-bs-toggle="modal" data-bs-target="#customtopicbox">篩選</a>
          </div>--%>
    
          <div class="card mt-4 mb-2 border-nonfocus">
              <div class="card-header bg-nonfocus fs-5">收文清單</div>
              <div class="card-body">
                <table id="tablist" class="table small">
                  <thead class="border-bottom border-dark-subtle">
                    <tr>
                      <th>表單名稱</th>
                      <th>表單編號</th>
                      <th>主旨</th>
                      <th>申請者</th>
                      <th>申請日期</th>
                      <th>表單狀態</th>
                      <th class="text-center">功能</th>
                    </tr>
                  </thead>
                  <tbody>
                  </tbody>
                </table>
            
                <%--<nav aria-label="Page navigation">
                  <ul class="pagination justify-content-center">
                    <li class="page-item disabled">
                      <a class="page-link">上一頁</a>
                    </li>
                    <li class="page-item"><a class="page-link" href="#">1</a></li>
                    <li class="page-item"><a class="page-link" href="#">2</a></li>
                    <li class="page-item"><a class="page-link" href="#">3</a></li>
                    <li class="page-item">
                      <a class="page-link" href="#">下一頁</a>
                    </li>
                  </ul>
                </nav>--%>
            
              </div>
            </div><!-- card -->
    
      </div><!-- container-ochi -->
    
    </div><!-- flex-grow-1 -->
</asp:Content>



