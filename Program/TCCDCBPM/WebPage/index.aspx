<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage/MasterPage.master" AutoEventWireup="true" CodeFile="index.aspx.cs" Inherits="WebPage_index" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <!--#include file="LeftMenu.html"-->
    <div class="flex-grow-1">
    
      <div class="container-ochi w-large-ochi mt-2" id="pagecontent">
    
        <nav class="itribreadcrumb d-none d-md-block" aria-label="breadcrumb">
          <ol class="breadcrumb bg-transparent">
            <li class="breadcrumb-item"><a href="#">首頁</a></li>
            <li class="breadcrumb-item active" aria-current="page">頁面標題</li>
          </ol>
        </nav>
    
        <div class="filetitlewrapper mt-1">
          <span class="filetitle"><h2>頁面標題</h2></span>
          <span class="itemright"></span>
        </div>

          <div class="card mt-4 mb-2 border-nonfocus">
              <div class="card-header bg-nonfocus fs-5">我的公文匣</div>
              <div class="card-body">
                <table class="table small">
                  <thead class="border-bottom border-dark-subtle">
                  <tr>
                    <th>表單名稱</th>
                    <th>表單編號</th>
                    <th>主旨</th>
                    <th class="text-center">申請者</th>
                    <th class="text-center">申請日期</th>
                    <th class="text-center">表單狀態</th>
                    <th class="text-center">功能</th>
                  </tr>
                  </thead>
                  <tbody>
                  <tr>
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
                  </tr>
                  </tbody>
                </table>
            
                <nav aria-label="Page navigation">
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
                </nav>
            
              </div>
            </div><!-- card -->
    
      </div><!-- container-ochi -->
    
    </div><!-- flex-grow-1 -->
</asp:Content>

