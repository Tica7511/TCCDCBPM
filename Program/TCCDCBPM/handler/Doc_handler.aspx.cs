using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Data;
using System.Runtime.Remoting.Contexts;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using Ionic.Zip;

public partial class handler_Doc_handler : System.Web.UI.Page
{
    XmlDocument xDoc = new XmlDocument();
    string xmlstr = string.Empty;

    Doc_DB ddb = new Doc_DB();
    DocManage_DB dmdb = new DocManage_DB();
    FileTable_DB fdb = new FileTable_DB();
    DocContent_DB dcdb = new DocContent_DB();
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
                case "new": // 新增公文
                    AddFile();
                    break;
                case "rv": // 查詢公文版本
                    GetVersionList();
                    break;
                case "rvd": // 查詢公文版本內容
                    GetVersionContentList();
                    break;
                case "cu": // 新增修改
                    //Add();
                    break;
                case "cuv": // 承攬商修改申訴
                    //AddVendor();
                    break;
                case "d": // 刪除
                    Delete();
                    break;
                default: // 查詢
                    //GetList();
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
    ///功    能: 公文列表
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

        ddb._KeyWord = SearchStr;
        DataTable dt = ddb.GetList();

        //string totalxml = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
        xmlstr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
        //xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + totalxml + xmlstr + "</root>";
        xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + xmlstr + "</root>";
    }

    ///-----------------------------------------------------
    ///功    能: 公文版本內容
    ///說    明:
    /// * Request["userid"]:userid
    /// * Request["PageNo"]:欲顯示的頁碼, 由零開始
    /// * Request["PageSize"]:每頁顯示的資料筆數, 未指定預設10
    /// * Request["WhereColumn"]:下拉搜尋類別
    /// * Request["SearchStr"]:關鍵字
    ///-----------------------------------------------------
    public void GetVersionList()
    {
        string guid = string.IsNullOrEmpty(Request["guid"]) ? "" : Server.HtmlEncode(Request["guid"].ToString().Trim());

        fdb._guid = guid;
        fdb._檔案類型 = "02";
        DataTable dt = fdb.GetData();

        //string totalxml = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
        xmlstr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
        //xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + totalxml + xmlstr + "</root>";
        xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + xmlstr + "</root>";
    }

    ///-----------------------------------------------------
    ///功    能: 公文版本內容
    ///說    明:
    /// * Request["userid"]:userid
    /// * Request["PageNo"]:欲顯示的頁碼, 由零開始
    /// * Request["PageSize"]:每頁顯示的資料筆數, 未指定預設10
    /// * Request["WhereColumn"]:下拉搜尋類別
    /// * Request["SearchStr"]:關鍵字
    ///-----------------------------------------------------
    public void GetVersionContentList()
    {
        string version = string.IsNullOrEmpty(Request["version"]) ? "" : Server.HtmlEncode(Request["version"].ToString().Trim());
        string guid = string.IsNullOrEmpty(Request["guid"]) ? "" : Server.HtmlEncode(Request["guid"].ToString().Trim());

        dcdb._guid = guid;
        dcdb._版本 = version;
        DataTable dt = dcdb.GetData();

        //string totalxml = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
        xmlstr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
        //xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + totalxml + xmlstr + "</root>";
        xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + xmlstr + "</root>";
    }

    ///-----------------------------------------------------
    ///功    能: 新增公文
    ///說    明:
    /// * Request["guid"]:guid
    /// * Request["mode"]: edit=編輯, view=檢視
    ///-----------------------------------------------------
    public void AddFile()
    {
        string UpLoadPath = ConfigurationManager.AppSettings["UploadFileRootDir"];
        string parentGuid = (string.IsNullOrEmpty(Request["parentGuid"])) ? "" : Server.HtmlEncode(Request["parentGuid"].ToString().Trim());
        string sn = (string.IsNullOrEmpty(Request["sn"])) ? "" : Server.HtmlEncode(Request["sn"].ToString().Trim());
        string title = (string.IsNullOrEmpty(Request["title"])) ? "" : Server.HtmlEncode(Request["title"].ToString().Trim());
        string docRandomGuid = Guid.NewGuid().ToString("N");
        string fileRandomGuid = Guid.NewGuid().ToString("N");
        string jwtToken = string.Empty;
        string fguid = string.Empty;
        string docName = string.Empty;
        string docNo = string.Empty;
        string fileName = string.Empty;
        string fileNewname = string.Empty;
        string fileextension = string.Empty;
        string version = string.Empty;

        fdb._父層guid = parentGuid;
        fdb._檔案類型 = "01";
        fdb._排序 = sn;
        DataTable dt = fdb.GetSnMaxData();

        if (dt.Rows.Count > 0)
        {
            fguid = dt.Rows[0]["guid"].ToString().Trim();
            parentGuid = dt.Rows[0]["父層guid"].ToString().Trim();
            docName = dt.Rows[0]["表單名稱"].ToString().Trim();
            docNo = dt.Rows[0]["表單編號"].ToString().Trim();
            fileName = dt.Rows[0]["原檔名"].ToString().Trim();
            fileNewname = fileName + "_v1";
            fileextension = dt.Rows[0]["附檔名"].ToString().Trim();
            version = dt.Rows[0]["版本"].ToString().Trim();
        }

        dmdb._guid = fguid;
        DataTable demodt = dmdb.GetDemoData();

        if (demodt.Rows.Count > 0)
        {
            UpLoadPath = Path.Combine(UpLoadPath,"公文文件",demodt.Rows[0]["RelativePath"].ToString().Trim());
        }

        string file_size = new FileInfo(UpLoadPath).Length.ToString();

        string targetDir = ConfigurationManager.AppSettings["UploadFileRootDir"] + "公文\\" + fileRandomGuid + "\\";
        string targetPath = Path.Combine(targetDir, fileNewname + fileextension);

        if (!Directory.Exists(targetDir))
        {
            Directory.CreateDirectory(targetDir);
        }

        File.Copy(UpLoadPath, targetPath, true);

        ProtectDocxForForms(targetPath);

        fdb._guid = fileRandomGuid;
        fdb._檔案類型 = "02";
        fdb._表單名稱 = docName;
        fdb._表單編號 = docNo;
        fdb._原檔名 = fileName;
        fdb._新檔名 = fileName + "_v1";
        fdb._附檔名 = fileextension;
        fdb._版本 = "1";
        fdb._建立者 = "333";
        fdb._修改者 = "333";
        fdb.InsertFile();

        ddb._guid = docRandomGuid;
        ddb._公文guid = fileRandomGuid;
        ddb._主旨 = title;
        ddb._狀態 = "0";
        ddb._建立者 = "333";
        ddb._修改者 = "333";
        ddb.InsertData();

        xmlstr = "<?xml version='1.0' encoding='utf-8'?><root><guid>" + fileRandomGuid + "</guid></root>";
    }

    ///-----------------------------------------------------
    ///功    能: 刪除公文
    ///說    明:
    /// * Request["guid"]: guid
    ///-----------------------------------------------------
    public void Delete()
    {
        string guid = (string.IsNullOrEmpty(Request["guid"])) ? "" : Server.HtmlEncode(Request["guid"].ToString().Trim());

        ddb._guid = guid;
        ddb._修改者 = "承辦人";
        ddb.DeleteData();
        xmlstr = "<?xml version='1.0' encoding='utf-8'?><root><Response>刪除完成</Response></root>";
    }

    public void ProtectDocxForForms(string docxPath)
    {
        using (ZipFile zip = ZipFile.Read(docxPath))
        {
            // 找 settings.xml
            ZipEntry settingsEntry = null;
            foreach (var e in zip.Entries)
            {
                if (e.FileName.Equals("word/settings.xml", StringComparison.OrdinalIgnoreCase))
                {
                    settingsEntry = e;
                    break;
                }
            }

            if (settingsEntry == null) return;

            // ★ 1. 讀取 settings.xml → 用 ExtractIntoMemory（不會關閉 stream）
            byte[] settingsBytes;
            using (var ms = new MemoryStream())
            {
                settingsEntry.Extract(ms);
                settingsBytes = ms.ToArray();
            }

            // ★ 2. 載入 XML
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.PreserveWhitespace = true;
            xmlDoc.Load(new MemoryStream(settingsBytes));

            string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", wNs);

            // 移除舊的 protection
            XmlNode oldProtect = xmlDoc.SelectSingleNode("//w:documentProtection", nsmgr);
            if (oldProtect != null)
            {
                oldProtect.ParentNode.RemoveChild(oldProtect);
            }

            // ★ 3. 新增 <w:documentProtection>
            XmlElement protect = xmlDoc.CreateElement("w", "documentProtection", wNs);
            protect.SetAttribute("edit", wNs, "forms");
            protect.SetAttribute("enforcement", wNs, "1");
            xmlDoc.DocumentElement.AppendChild(protect);

            // ★ 4. 將 XML 存入新的 memory stream（不要 reuse 舊的）
            byte[] newXmlBytes;
            using (MemoryStream outMs = new MemoryStream())
            {
                xmlDoc.Save(outMs);
                newXmlBytes = outMs.ToArray();
            }

            // ★ 5. 更新 entry（使用 byte[]，絕對不會有 stream 被關掉的問題）
            zip.UpdateEntry("word/settings.xml", newXmlBytes);

            zip.Save();
        }
    }
}