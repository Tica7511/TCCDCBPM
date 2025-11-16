<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DocManageAdd.aspx.cs" Inherits="WebPage_DocManageAdd" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="content-language" content="zh-TW" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<meta name="keywords" content="關鍵字內容" />
<meta name="description" content="描述" /><!--告訴搜尋引擎這篇網頁的內容或摘要。-->
<meta name="generator" content="Notepad" /><!--告訴搜尋引擎這篇網頁是用什麼軟體製作的。-->
<meta name="author" content="ochison" /><!--告訴搜尋引擎這篇網頁是由誰製作的。-->
<meta name="copyright" content="本網頁著作權所有" /><!--告訴搜尋引擎這篇網頁是...... -->
<meta name="revisit-after" content="3 days" /><!--告訴搜尋引擎3天之後再來一次這篇網頁，也許要重新登錄。-->
<title>電子公文系統</title>
<link rel="stylesheet" href="<%= ResolveUrl("~/css/bootstrap.min.css") %>" /><!-- bootstrap 5.3.3 -->
<link rel="stylesheet" href="<%= ResolveUrl("~/css/wrapkitStyle.css") %>" />
<link rel="stylesheet" href="<%= ResolveUrl("~/css/fontawesome.min.css") %>" /><!-- css icon fontawesome 6.6 -->
<link rel="stylesheet" href="<%= ResolveUrl("~/css/OCsdiebarmenu.css") %>" />
<link rel="stylesheet" href="<%= ResolveUrl("~/css/OClayout.css") %>" /><!-- sidebar menu -->
<link rel="stylesheet" href="<%= ResolveUrl("~/css/style.css") %>" /><!-- 本系統專用 -->
<script src="<%= ResolveUrl("~/js/bootstrap.bundle.min.js") %>"></script><!-- bootstrap 5.3 -->
<script src="<%= ResolveUrl("~/js/OCsdiebarmenu.js") %>"></script>
<script src="<%= ResolveUrl("~/js/jquery-3.7.0.min.js") %>" type="text/javascript"></script>
<script src="<%= ResolveUrl("~/js/jquery-ui.js") %>" type="text/javascript"></script>
<script src="<%= ResolveUrl("~/js/NickCommon.js") %>" type="text/javascript"></script>
<script src="<%= ResolveUrl("~/page_js/DocManageAdd.js") %>" type="text/javascript"></script>
<script src="http://172.20.10.5:8080/web-apps/apps/api/documents/api.js"></script>

</head>
<body>
    <input type="hidden" id="fGuid" />
    <div id="div_container"></div>
</body>
</html>
