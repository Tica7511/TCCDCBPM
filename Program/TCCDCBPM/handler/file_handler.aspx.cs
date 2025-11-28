using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Data;
using System.Configuration;
using System.IdentityModel.Tokens.Jwt;
using System.Text;
using Microsoft.IdentityModel.Tokens;
using System.IO;

public partial class handler_file_handler : System.Web.UI.Page
{
    XmlDocument xDoc = new XmlDocument();
    string xmlstr = string.Empty;

    FileTable_DB fdb = new FileTable_DB();
    DocManage_DB ddb = new DocManage_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            string crud = (string.IsNullOrEmpty(Request["crud"])) ? "" : Server.HtmlEncode(Request["crud"].ToString().Trim());            
            xmlstr = "<?xml version='1.0' encoding='utf-8'?><root><Response>完成</Response></root>";

            switch (crud)
            {
                case "demo": // 範本
                    GetDemoList(crud);
                    break;
                case "new": // 新增公文
                    GetDoc(crud);
                    break;
                case "rf": // 查詢公文範本檔案
                    GetFileList();
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
                    GetDemoList(crud);
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
    ///功    能: 取得公文資料
    ///說    明:
    /// * Request["guid"]:guid
    /// * Request["mode"]: edit=編輯, view=檢視
    ///-----------------------------------------------------
    public void GetDoc(string filetype)
    {
        string guid = (string.IsNullOrEmpty(Request["guid"])) ? "" : Server.HtmlEncode(Request["guid"].ToString().Trim());
        string mode = (string.IsNullOrEmpty(Request["mode"])) ? "" : Server.HtmlEncode(Request["mode"].ToString().Trim());
        string fileRandomGuid = Guid.NewGuid().ToString("N");
        string parentGuid = string.Empty;
        string jwtToken = string.Empty;
        string fguid = string.Empty;
        string fileName = string.Empty;
        string fileNewname = string.Empty;
        string fileextension = string.Empty;
        string version = string.Empty;
        string filecategory = string.Empty;

        fdb._guid = guid;
        fdb._檔案類型 = "02";
        DataTable dt = fdb.GetSnMaxData();

        if (dt.Rows.Count > 0)
        {
            fguid = dt.Rows[0]["guid"].ToString().Trim();
            parentGuid = dt.Rows[0]["父層guid"].ToString().Trim();
            fileName = dt.Rows[0]["原檔名"].ToString().Trim();
            fileNewname = dt.Rows[0]["新檔名"].ToString().Trim();
            fileextension = dt.Rows[0]["附檔名"].ToString().Trim();
            version = dt.Rows[0]["版本"].ToString().Trim();
            filecategory = "File";
            jwtToken = GenerateJwt(filecategory, filetype, fguid, parentGuid, fileRandomGuid, version, fileName, fileNewname, fileextension, mode);
        }

        xmlstr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
        xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + xmlstr + "<filecategory>" + filecategory + "</filecategory><fileName>" + fileName + fileextension +
                "</fileName><onlyofficeguid>" + fguid + "</onlyofficeguid><token>" + jwtToken + "</token><fileRandomGuid>" + fileRandomGuid + "</fileRandomGuid><mGuid>333" +
                 "</mGuid><mName>測試人員</mName><parentguid>" + parentGuid + "</parentguid><version>" + version + "</version></root>";
    }

    ///-----------------------------------------------------
    ///功    能: 取得公文範本資料
    ///說    明:
    /// * Request["guid"]:guid
    /// * Request["mode"]: edit=編輯, view=檢視
    ///-----------------------------------------------------
    public void GetDemoList(string filetype)
    {
        string parentGuid = (string.IsNullOrEmpty(Request["parentGuid"])) ? "" : Server.HtmlEncode(Request["parentGuid"].ToString().Trim());
        string sn = (string.IsNullOrEmpty(Request["sn"])) ? "" : Server.HtmlEncode(Request["sn"].ToString().Trim());
        string mode = (string.IsNullOrEmpty(Request["mode"])) ? "" : Server.HtmlEncode(Request["mode"].ToString().Trim());
        string fileRandomGuid = Guid.NewGuid().ToString("N");
        string jwtToken = string.Empty;
        string fguid = string.Empty;
        string fileName = string.Empty;
        string fileNewname = string.Empty;
        string fileextension = string.Empty;
        string version = string.Empty;
        string filecategory = string.Empty;

        fdb._父層guid = parentGuid;
        fdb._檔案類型 = "01";
        fdb._排序 = sn;
        DataTable dt = fdb.GetSnMaxData();

        if (dt.Rows.Count > 0)
        {
            fguid = dt.Rows[0]["guid"].ToString().Trim();
            parentGuid = dt.Rows[0]["父層guid"].ToString().Trim();
            fileName = dt.Rows[0]["原檔名"].ToString().Trim();
            fileNewname = dt.Rows[0]["新檔名"].ToString().Trim();
            fileextension = dt.Rows[0]["附檔名"].ToString().Trim();
            version = dt.Rows[0]["版本"].ToString().Trim();
            filecategory = "Demo";
            jwtToken = GenerateJwt(filecategory, filetype, fguid, parentGuid, fileRandomGuid, version, fileName, fileNewname, fileextension, mode);
        }

        xmlstr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
        xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + xmlstr + "<filecategory>" + filecategory + "</filecategory><fileName>" + fileName + fileextension +
                "</fileName><onlyofficeguid>" + fguid + "</onlyofficeguid><token>" + jwtToken + "</token><fileRandomGuid>" + fileRandomGuid + "</fileRandomGuid><mGuid>333" +
                 "</mGuid><mName>測試人員</mName><parentguid>" + parentGuid + "</parentguid><version>" + version + "</version></root>";
    }

    ///-----------------------------------------------------
    ///功    能: 取得公文範本列表
    ///說    明:
    /// * Request["parentguid"]:父層guid
    ///-----------------------------------------------------
    public void GetFileList()
    {
        string parentguid = (string.IsNullOrEmpty(Request["parentguid"])) ? "" : Server.HtmlEncode(Request["parentguid"].ToString().Trim());

        fdb._父層guid = parentguid;
        fdb._檔案類型 = "01";
        DataTable dt = fdb.GetMaxData();

        xmlstr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
        xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + xmlstr + "</root>";
    }

    ///-----------------------------------------------------
    ///功    能: 取得 onlyoffice JWT token
    ///說    明:
    ///-----------------------------------------------------
    public static string GenerateJwt(string filecategory, string filetype, string fileguid, string parentguid, 
        string fileRandomGuid, string version, string fileName, string fileNewname, string fileextension, string mode)
    {
        bool status = true;

        if (mode == "view")
        {
            status = false;
        }

        //將 JWT secret 包成 HmacSha256 的加密格式
        var secret = ConfigurationManager.AppSettings["JwtSecret"];
        var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secret));
        var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);

        //將payload用JSON格式包成跟前端一樣
        var payload = new JwtPayload
        {
            { "document", new Dictionary<string, object>
                {
                    { "fileType", "docx" },
                    { "key", fileguid + "_" + parentguid + "_" + version + "_" + filetype + "_" + fileRandomGuid },
                    { "title", fileName + fileextension },
                    { "url", "http://172.20.10.5:7594/DOWNLOAD.aspx?category="+ filecategory +"&guid=" + fileguid}
                }
            },
            { "documentType", "word" },
            { "editorConfig", new Dictionary<string, object>
                {
                    { "mode", mode },
                    { "lang", "zh-TW" },
                    { "callbackUrl", "http://172.20.10.5:7594/handler/SaveCallback.aspx" },
                    { "canUseHistory", true },
                    { "customization", new Dictionary<string, object>
                        {
                            { "forcesave", true },
                            { "autosave", true },
                            { "autosaveInterval", 60 },
                            //{ "trackChanges", true },
                            { "buttons", new Dictionary<string, object>
                                {
                                    { "print", false },
                                    { "download", false }
                                }
                            },
                            { "logo", new Dictionary<string, object>
                                {
                                    { "image", "http://172.20.10.5:7594/images/tccLogo.png" },
                                    { "url", "https://www.cogen.com.tw/tw/" }
                                }
                            },
                            { "layout", new Dictionary<string, object>
                                {
                                    { "leftMenu", new Dictionary<string, object>
                                        {
                                            { "mode", false }
                                        }
                                    },
                                    { "rightMenu", new Dictionary<string, object>
                                        {
                                            { "mode", false }
                                        }
                                    },
                                    { "toolbar", new Dictionary<string, object>
                                        {
                                            { "layout", false },
                                            { "references", false },
                                            { "protect", false },
                                            { "plugins", false }
                                        }
                                    }
                                }
                            },
                            { "features", new Dictionary<string, object>
                                {
                                    { "watermark", false }
                                }
                            }
                        }
                    },
                    { "user", new Dictionary<string, object>
                        {
                            { "id", "333" },
                            { "name", "測試人員" }
                        }
                    },
                    { "history", new Dictionary<string, object>
                        {
                            { "serverVersion", true }
                        }
                    }
                }
            },
            { "permissions", new Dictionary<string, object>
                {
                    { "edit", status },
                    //{ "review", true },
                    { "comment", status },
                    { "print", false },
                    { "download", false }
                }
            },
            { "height", "800px" },
            { "width", "100%" },
            { "type", "desktop" }
        };

        //組合 header + payload + signature
        var token = new JwtSecurityToken(new JwtHeader(creds), payload);

        //回傳 token 字串
        return new JwtSecurityTokenHandler().WriteToken(token);
    }
}