using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Data;
using System.Runtime.Remoting.Contexts;

public partial class handler_DocMange_handler : System.Web.UI.Page
{
    XmlDocument xDoc = new XmlDocument();
    string xmlstr = string.Empty;

    DocManage_DB db = new DocManage_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            string crud = (string.IsNullOrEmpty(Request["crud"])) ? "" : Server.HtmlEncode(Request["crud"].ToString().Trim());
            xmlstr = "<?xml version='1.0' encoding='utf-8'?><root><Response>完成</Response></root>";

            switch (crud)
            {
                case "r": // 查詢
                    GetList();
                    break;
                case "rd": // 查詢單筆
                    //GetDetail();
                    break;
                case "rj": // 查詢工程
                    //GetJobList();
                    break;
                case "cu": // 新增修改
                    //Add();
                    break;
                case "cuv": // 承攬商修改申訴
                    //AddVendor();
                    break;
                case "d": // 刪除
                    //Delete();
                    break;
                default: // 查詢
                    GetList();
                    break;
            }

            xDoc.LoadXml(xmlstr);
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }

    ///-----------------------------------------------------
    ///功    能: 公文管理列表
    ///說    明:
    /// * Request["userid"]:userid
    /// * Request["PageNo"]:欲顯示的頁碼, 由零開始
    /// * Request["PageSize"]:每頁顯示的資料筆數, 未指定預設10
    /// * Request["WhereColumn"]:下拉搜尋類別
    /// * Request["SearchStr"]:關鍵字
    ///-----------------------------------------------------
    public void GetList()
    {
        //string userid = (string.IsNullOrEmpty(Request["userid"])) ? "" : Server.HtmlEncode(Request["userid"].ToString().Trim());
        //string useridx = (string.IsNullOrEmpty(Request["useridx"])) ? "" : Server.HtmlEncode(Request["useridx"].ToString().Trim());
        //string PageNo = (string.IsNullOrEmpty(Request["PageNo"])) ? "0" : Server.HtmlEncode(Request["PageNo"].ToString().Trim());
        //int PageSize = (string.IsNullOrEmpty(Request["PageSize"])) ? 10 : int.Parse(Server.HtmlEncode(Request["PageSize"].ToString().Trim()));
        //string whereColumn = string.IsNullOrEmpty(Request["whereColumn"]) ? "" : Server.HtmlEncode(Request["whereColumn"].ToString().Trim());
        string SearchStr = string.IsNullOrEmpty(Request["SearchStr"]) ? "" : Server.HtmlEncode(Request["SearchStr"].ToString().Trim());

        //計算起始與結束
        //int pageEnd = (int.Parse(PageNo) + 1) * PageSize;
        //int pageStart = pageEnd - PageSize + 1;

        db._KeyWord = SearchStr;
        DataSet ds = db.GetList();

        //string totalxml = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
        xmlstr = DataTableToXml.ConvertDatatableToXML(ds.Tables[0], "dataList", "data_item");
        //xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + totalxml + xmlstr + "</root>";
        xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + xmlstr + "</root>";
    }
}