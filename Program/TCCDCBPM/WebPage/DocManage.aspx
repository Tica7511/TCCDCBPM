<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage/MasterPage.master" AutoEventWireup="true" CodeFile="DocManage.aspx.cs" Inherits="WebPage_DocManage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script src="<%= ResolveUrl("~/page_js/DocManage.js") %>" type="text/javascript"></script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <!--#include file="LeftMenu.html"-->
    <div class="flex-grow-1">
    
      <div class="container-ochi w-large-ochi mt-2" id="pagecontent">
    
        <nav class="itribreadcrumb d-none d-md-block" aria-label="breadcrumb">
          <ol class="breadcrumb bg-transparent">
            <li class="breadcrumb-item"><a href="index.aspx">首頁</a></li>
            <li class="breadcrumb-item active" aria-current="page">公文管理</li>
            <li class="breadcrumb-item active" aria-current="page">公文範本管理</li>
          </ol>
        </nav>
    
        <div class="filetitlewrapper mt-1">
          <span class="filetitle"><h2>公文範本管理</h2></span>
          <span class="itemright"></span>
        </div>


          <%--<div class="text-start mt-2">
              <a href="javascript:void(0);" class="btn btn-primary" id="assignsendOK" data-bs-toggle="modal" data-bs-target="#customtopicbox">篩選</a>
          </div>--%>
    
          <div class="card mt-4 mb-2 border-nonfocus">
              <div class="card-header bg-nonfocus fs-5">公文範本列表</div>
              <div class="card-body">
                <table id="tablist" class="table small">
                  <thead class="border-bottom border-dark-subtle">
                    <tr>
                      <th>範本名稱</th>
                      <th>所屬類別</th>
                      <th>上傳日期</th>
                      <th class="text-center">功能</th>
                    </tr>
                  </thead>
                  <tbody>
                     <%--<tr>
                       <td class="align-middle">收文表單</td>
                       <td class="align-middle">
                         <div class="muted">250900012</div>
                       </td>
                       <td class="align-middle">台灣雷力企業聯合會TEPA</td>
                       <td class="text-center align-middle">何金銀</td>
                       <td class="text-center align-middle">2025/10/29</td>
                       <td class="text-center align-middle">審核中</td>
                       <td class="text-center">
                         <div class="btn-group btn-group-sm" role="group">
                           <button type="button" class="btn btn-outline-dark">檢視</button>
                         </div>
                       </td>
                     </tr>--%>
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

    <div class="modal fade" id="customtopicbox" tabindex="-1">
      <div class="modal-dialog modal-sm modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title">分案對象</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
          </div>
          <div class="modal-body">
    
          <div id="html1">
            <ul>
              <li>行管部
                <ul>
                  <li>連家玟</li>
                  <li>馮展榮</li>
                  <li>李麗敏</li>
                  <li>全嘉傑</li>
                  <li>辛文暄</li>
                </ul>
              </li>
              <li>財務部
                <ul>
                  <li>黄士晉</li>
                  <li>賈宜德</li>
                  <li>蔣智安</li>
                  <li>包唯中</li>
                </ul>
              </li>
              <li>技術部
                <ul>
                  <li>王芷茵</li>
                  <li>侯韶恩</li>
                  <li>蕭昊天</li>
                </ul>
              </li>
            </ul>
          </div>
          </div><!-- modal-body -->
          <div class="modal-footer">
            <button type="button" class="btn btn-outline-dark" data-bs-dismiss="modal">取消</button>
            <button type="button" class="btn btn-outline-dark" data-bs-dismiss="modal" id="treelistOK">確定</button>
          </div>
        </div>
      </div>
    </div>
</asp:Content>

