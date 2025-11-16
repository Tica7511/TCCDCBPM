using System;
using System.Web.UI;
using System.Configuration;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Text;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Xml.Linq;
using System.IO.Compression;
using System.Collections.Generic;
using System.Linq;
using Ionic.Zip;
using System.Xml;
using System.Web.UI.WebControls;

public partial class handler_SaveCallback : System.Web.UI.Page
{
    DocManage_DB dmdb = new DocManage_DB();
    FileTable_DB fdb = new FileTable_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 儲存 onlyoffice
        ///說    明:
        ///-----------------------------------------------------

        /// Transaction
        SqlConnection oConn = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        oConn.Open();
        SqlCommand oCmmd = new SqlCommand();
        oCmmd.Connection = oConn;
        SqlTransaction myTrans = oConn.BeginTransaction();
        oCmmd.Transaction = myTrans;

        try
        {
            #region 全域字串
            string json;
            int status = 0;
            string nowVersion = string.Empty;
            string key = string.Empty;
            string filepath = string.Empty;
            string fileUrl = string.Empty;
            string changesurl = string.Empty;
            string editorUser = "unknown";
            string fileName = "";
            string fileGuid = string.Empty;
            string fileNewGuid = Guid.NewGuid().ToString("N");
            string fileparentGuid = string.Empty;
            string PublicFileName = string.Empty;
            string PublicFileNewName = string.Empty;
            string PublicFileExtension = string.Empty;
            string PublicFileFullName = string.Empty;
            string PublicFileSn = string.Empty;
            string PublicFileVersion = string.Empty;
            string UpLoadPath = ConfigurationManager.AppSettings["UploadFileRootDir"];
            #endregion

            #region 讀取json
            using (var reader = new StreamReader(Request.InputStream))
                json = reader.ReadToEnd();

            JObject data = JsonConvert.DeserializeObject<JObject>(json);

            status = Convert.ToInt32(data["status"]);
            if (data["key"] != null)
                key = data["key"].ToString();
            if (data["url"] != null)
                fileUrl = data["url"].ToString();
            if (data["changesurl"] != null)
                changesurl = data["changesurl"].ToString();
            if (data["users"] != null && data["users"].Type == JTokenType.Array && data["users"].HasValues)
            {
                editorUser = data["users"][0].ToString();
            }

            if (!string.IsNullOrEmpty(key))
            {
                string[] parts = key.Split('_');
                fileGuid = parts.Length > 0 ? parts[0] : "";
                fileparentGuid = parts.Length > 1 ? parts[1] : "";
                PublicFileVersion = parts.Length > 2 ? parts[2] : "";
            }

            File.AppendAllText(Server.MapPath("~/log-callback.txt"),
                DateTime.Now + "\nstatus=" + status.ToString() + "\nkey=" + key + "\nfileUrl=" + fileUrl + "\nfileName=" + fileName +
                "\neditorUser=333\nchange=" + changesurl + "\n\n");

            File.AppendAllText(Server.MapPath("~/log-callback.txt"),
                DateTime.Now + "\nfileGuid=" + fileGuid + "\nfileparentGuid=" + fileparentGuid + "\nPublicFileVersion=" + PublicFileVersion + "\n\n");
            #endregion

            #region 取得檔案

            //取得當前路徑以及檔案名稱
            dmdb._guid = fileGuid;
            DataTable dmdt = dmdb.GetDemoDataNoFile();
            if (dmdt.Rows.Count > 0)
            {
                fileName = dmdt.Rows[0]["新檔名"].ToString().Trim() + dmdt.Rows[0]["附檔名"].ToString().Trim();
                PublicFileName = dmdt.Rows[0]["原檔名"].ToString().Trim();
                PublicFileNewName = dmdt.Rows[0]["新檔名"].ToString().Trim();
                PublicFileExtension = dmdt.Rows[0]["附檔名"].ToString().Trim();
                PublicFileSn = dmdt.Rows[0]["排序"].ToString().Trim();
                filepath = UpLoadPath + "公文文件\\" + dmdt.Rows[0]["RelativePath"].ToString().Trim();
            }

            //dmdb._guid = fileGuid;
            //DataTable dmtdt = dmdb.GetDemoDataNoFile();

            //if (dmtdt.Rows.Count > 0)
            //{
            //    string filepathTest = UpLoadPath + "公文文件\\" + dmtdt.Rows[0]["RelativePath"].ToString().Trim();
            //    filepathTest += PublicFileNewName + PublicFileExtension;

            //    ExtractControlsViaBuilder(filepathTest);
            //}                       

            File.AppendAllText(Server.MapPath("~/log-callback.txt"),
                DateTime.Now + "\nfileName=" + fileName + "\n\n");
            #endregion

            #region 手動儲存
            if (status == 6 && !string.IsNullOrEmpty(fileUrl))
            {
                #region 儲存檔案進手動儲存資料夾

                //取得目前最新版版本
                fdb._父層guid = fileparentGuid;
                fdb._檔案類型 = "01";
                fdb._排序 = PublicFileSn;
                DataTable mddt = fdb.GetSnMaxData();

                if (mddt.Rows.Count > 0)
                {
                    nowVersion = mddt.Rows[0]["版本"].ToString().Trim();
                    nowVersion = (Convert.ToInt32(nowVersion) + 1).ToString();
                }
                else
                {
                    nowVersion = "1";
                }

                using (var client = new WebClient())
                {
                    #region 新增手動儲存資料夾檔案
                    byte[] fileBytes = client.DownloadData(fileUrl + PublicFileNewName + PublicFileExtension);
                    string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                    PublicFileFullName = PublicFileName + "_v" + nowVersion + "_" + timeStamp;
                    string manualFolder = Path.Combine(filepath, "手動儲存");

                    if (!Directory.Exists(manualFolder))
                    {
                        Directory.CreateDirectory(manualFolder);
                    }
                    string saveFullPath = Path.Combine(manualFolder, PublicFileFullName + PublicFileExtension);
                    File.WriteAllBytes(saveFullPath, fileBytes);
                    #endregion

                    #region 新增正式區檔案
                    PublicFileFullName = PublicFileName + "_v" + nowVersion;
                    string realfilepath = Path.Combine(filepath, PublicFileFullName + PublicFileExtension);
                    File.WriteAllBytes(realfilepath, fileBytes);
                    #endregion
                }
                #endregion

                #region 儲存進附件檔資料表
                fdb._guid = fileNewGuid;
                fdb._年度 = "";
                fdb._原檔名 = PublicFileName;
                fdb._新檔名 = PublicFileFullName;
                fdb._附檔名 = PublicFileExtension;
                fdb._版本 = nowVersion;
                fdb._檔案大小 = "";
                fdb._修改者 = editorUser;
                fdb._修改日期 = DateTime.Now;
                fdb._建立者 = editorUser;
                fdb._建立日期 = DateTime.Now;
                fdb.UpdateFile_Trans(oConn, myTrans);
                #endregion

                #region 儲存進指定表單資料表
                // 1. 解析內容控制項
                var controls = ParseContentControls(filepath + PublicFileFullName + PublicFileExtension);

                // 2. 把 Word 裡的欄位狀況同步到 公文欄位定義表
                foreach (var c in controls)
                {
                    string tag = c.Tag;
                    string type = c.Type;
                    string value = c.Text;

                    // 找是否已存在
                    DataRow exist = GetFieldDef(oConn, myTrans, fileGuid, Convert.ToInt32(PublicFileVersion), tag);

                    if (exist == null)
                    {
                        // ★ 新增欄位定義
                        InsertFieldDef(oConn, myTrans, fileGuid, Convert.ToInt32(PublicFileVersion),
                                       tag, type, editorUser);

                        // ★ 寫 log
                        InsertFieldChangeLog(oConn, myTrans, fileGuid,
                                             Convert.ToInt32(PublicFileVersion),
                                             tag, "Add",
                                             null, null,
                                             null, type,
                                             editorUser);
                    }
                    else
                    {
                        string oldType = exist["項目類型"].ToString();

                        if (oldType != type)
                        {
                            // ★ 更新欄位類型
                            UpdateFieldType(oConn, myTrans, fileGuid,
                                            Convert.ToInt32(PublicFileVersion),
                                            tag, type, editorUser);

                            // ★ 寫 log
                            InsertFieldChangeLog(oConn, myTrans, fileGuid,
                                                 Convert.ToInt32(PublicFileVersion),
                                                 tag, "Modify",
                                                 oldType, type,
                                                 null, null,
                                                 editorUser);
                        }
                    }
                }



                foreach (var c in controls)
                {
                    File.AppendAllText(Server.MapPath("~/log-callback.txt"),
                        "[" + DateTime.Now + "]\r\n" +
                        "Tag = " + c.Tag + "\r\n" +
                        "Type = " + c.Type + "\r\n" +
                        "Value = " + c.Text + "\r\n" +
                        (c.Items.Count > 0 ? "Items = " + string.Join(",", c.Items) + "\r\n" : "") +
                        "\r\n"
                    );
                }
                #endregion
            }
            #endregion

            #region 自動儲存
            if (status == 2 && !string.IsNullOrEmpty(fileUrl))
            {
                fdb._父層guid = fileparentGuid;
                fdb._檔案類型 = "01";
                fdb._排序 = PublicFileSn;
                DataTable mddt = fdb.GetSnMaxData();

                if (mddt.Rows.Count > 0)
                {
                    nowVersion = mddt.Rows[0]["版本"].ToString().Trim();
                }
                else
                {
                    nowVersion = "1";
                }


                #region 儲存檔案進資料夾
                using (var client = new WebClient())
                {
                    byte[] fileBytes = client.DownloadData(fileUrl + PublicFileNewName + PublicFileExtension);
                    PublicFileFullName = PublicFileName + "_v" + nowVersion;
                    string realfilepath = Path.Combine(filepath, PublicFileFullName + PublicFileExtension);

                    File.WriteAllBytes(realfilepath, fileBytes);
                }
                #endregion
            }
            #endregion

            myTrans.Commit();

            // 回應 OnlyOffice
            Response.Clear();
            Response.ContentType = "application/json";
            Response.Write("{\"error\":0}");
            Context.ApplicationInstance.CompleteRequest(); // 安全結束請求
        }
        catch (Exception ex)
        {
            File.AppendAllText(Server.MapPath("~/log-callback.txt"),
                DateTime.Now + "\n" + ex.Message + "\r\n" + ex.StackTrace);

            Common.InsertLogsTran(oConn, myTrans, Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath),
                    System.Reflection.MethodBase.GetCurrentMethod().Name, "錯誤：" + ex.Message + "\r\n" + ex.StackTrace);

            myTrans.Rollback();
            Response.Clear();
            Response.ContentType = "application/json";
            Response.Write("{\"error\":1}");
            Context.ApplicationInstance.CompleteRequest(); // 不讓 OnlyOffice 掛掉
        }
        finally
        {
            oCmmd.Connection.Close();
            oConn.Close();
        }
    }

    public void ExtractControlsViaBuilder(string docFilePath)
    {
        var script = System.IO.File.ReadAllText(Server.MapPath("~/js/extract_controls.docbuilder"));

        var payload = new
        {
            script = script,
            filetype = "docx",
            outputtype = "txt"
        };

        using (var client = new WebClient())
        {
            client.Headers[HttpRequestHeader.ContentType] = "application/json";

            byte[] docBytes = File.ReadAllBytes(docFilePath);

            string docBase64 = Convert.ToBase64String(docBytes);

            var body = new
            {
                script = payload.script,
                filetype = "docx",
                outputtype = "txt",
                async = false,
                document = new
                {
                    fileType = "docx",
                    content = docBase64
                }
            };

            var token = CreateDocBuilderToken(ConfigurationManager.AppSettings["JwtSecret"]);

            var request = (HttpWebRequest)WebRequest.Create("http://172.20.10.5:8080/docbuilder");
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Headers.Add("Authorization", "Bearer " + token);

            var jsonBody = JsonConvert.SerializeObject(body);
            var bytes = Encoding.UTF8.GetBytes(jsonBody);
            using (var stream = request.GetRequestStream())
            {
                stream.Write(bytes, 0, bytes.Length);
            }

            string result;
            var response = (HttpWebResponse)request.GetResponse();
            using (var reader = new StreamReader(response.GetResponseStream()))
            {
                result = reader.ReadToEnd();
            }

            JObject json = JObject.Parse(result);
            string logPath = Server.MapPath("~/log-callback.txt");
            string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // 把 json 內容寫進 log 檔案
            File.AppendAllText(Server.MapPath("~/log-callback.txt"),
                DateTime.Now + "\njson=" + json + "\r\n");
        }
    }

    public string CreateDocBuilderToken(string secret)
    {
        var payload = new JwtPayload
        {
            { "exp", DateTimeOffset.UtcNow.ToUnixTimeSeconds() + 300 }
        };

        var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secret));
        var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);
        var token = new JwtSecurityToken(new JwtHeader(creds), payload);

        return new JwtSecurityTokenHandler().WriteToken(token);
    }

    public class ContentControlInfo
    {
        public string Tag { get; set; }
        public string Type { get; set; }
        public string Text { get; set; }
        public List<string> Items { get; set; }
    }

    public List<ContentControlInfo> ParseContentControls(string docxPath)
    {
        List<ContentControlInfo> result = new List<ContentControlInfo>();

        //======================
        // 1. 開啟 DOCX (ZIP)
        //======================
        ZipFile zip = ZipFile.Read(docxPath);

        ZipEntry docEntry = null;

        foreach (ZipEntry e in zip.Entries)
        {
            if (e.FileName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase))
            {
                docEntry = e;
                break;
            }
        }

        if (docEntry == null)
            return result;

        //======================
        // 2. Extract document.xml
        //======================
        XmlDocument xmlDoc = new XmlDocument();

        using (MemoryStream ms = new MemoryStream())
        {
            docEntry.Extract(ms);
            ms.Position = 0;
            xmlDoc.Load(ms);
        }

        XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
        nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        nsmgr.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

        XmlNodeList sdtList = xmlDoc.SelectNodes("//w:sdt", nsmgr);

        //=======================================================
        // ⭐ Step 0：先掃描所有已存在 tag → 找最大 cc_xxxx
        //=======================================================
        int ccIndex = 1;

        foreach (XmlNode sdt in sdtList)
        {
            XmlNode tagNode = sdt.SelectSingleNode(".//w:sdtPr/w:tag", nsmgr);

            if (tagNode != null)
            {
                XmlAttribute valAttr = tagNode.Attributes["w:val"];
                if (valAttr != null)
                {
                    string tag = valAttr.Value;

                    if (tag != null && tag.StartsWith("cc_"))
                    {
                        int number;
                        string numStr = tag.Substring(3);

                        if (int.TryParse(numStr, out number))
                        {
                            if (number >= ccIndex)
                                ccIndex = number + 1;  // 下一個應用的編號
                        }
                    }
                }
            }
        }

        //=======================================================
        // ⭐ Step 1：開始逐一處理每個內容控制項
        //=======================================================
        foreach (XmlNode sdt in sdtList)
        {
            ContentControlInfo info = new ContentControlInfo();
            info.Items = new List<string>();

            XmlNode sdtPr = sdt.SelectSingleNode(".//w:sdtPr", nsmgr);

            //======================
            // 3. 讀取 tag
            //======================
            info.Tag = "";
            XmlNode tagNode = sdtPr.SelectSingleNode("./w:tag", nsmgr);

            if (tagNode != null)
            {
                XmlAttribute valAttr = tagNode.Attributes["w:val"];
                if (valAttr != null)
                    info.Tag = valAttr.Value;
            }

            //======================
            // 4. 無 tag → 自動補 cc_xxxx（使用更新後的 ccIndex）
            //======================
            if (info.Tag == null || info.Tag.Trim() == "")
            {
                string newTag = "cc_" + ccIndex.ToString("0000");
                ccIndex++;

                info.Tag = newTag;

                if (tagNode == null)
                {
                    XmlElement newTagElement = xmlDoc.CreateElement("w", "tag",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                    XmlAttribute valAttr = xmlDoc.CreateAttribute("w", "val",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                    valAttr.Value = newTag;
                    newTagElement.Attributes.Append(valAttr);

                    sdtPr.AppendChild(newTagElement);
                }
                else
                {
                    XmlAttribute valAttr = tagNode.Attributes["w:val"];
                    if (valAttr == null)
                    {
                        valAttr = xmlDoc.CreateAttribute("w", "val",
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        tagNode.Attributes.Append(valAttr);
                    }

                    valAttr.Value = newTag;
                }
            }

            //======================
            // 5. 判斷控制項類型
            //======================
            info.Type = "";
            XmlNode cbNode = sdtPr.SelectSingleNode("./w:checkbox", nsmgr);
            if (cbNode == null)
                cbNode = sdtPr.SelectSingleNode("./w14:checkbox", nsmgr);

            if (cbNode != null) info.Type = "checkbox";
            else if (sdtPr.SelectSingleNode("./w:dropDownList", nsmgr) != null) info.Type = "dropdown";
            else if (sdtPr.SelectSingleNode("./w:comboBox", nsmgr) != null) info.Type = "combobox";
            else if (sdtPr.SelectSingleNode("./w:date", nsmgr) != null) info.Type = "date";
            else info.Type = "text";

            //======================
            // 6. 讀取值
            //======================
            StringBuilder sb = new StringBuilder();
            XmlNodeList textNodes = sdt.SelectNodes(".//w:t", nsmgr);

            foreach (XmlNode t in textNodes)
                sb.Append(t.InnerText);

            info.Text = sb.ToString();

            // checkbox → 轉 0/1
            if (info.Type == "checkbox")
            {
                string raw = "0";
                XmlNode checkedNode = sdt.SelectSingleNode(".//w14:checkbox/w14:checked", nsmgr);

                if (checkedNode == null)
                    checkedNode = sdt.SelectSingleNode(".//w:checkbox/w:checked", nsmgr);

                if (checkedNode != null)
                {
                    XmlAttribute v1 = checkedNode.Attributes["w14:val"];
                    XmlAttribute v2 = checkedNode.Attributes["w:val"];

                    string val = "0";
                    if (v1 != null) val = v1.Value;
                    if (v2 != null) val = v2.Value;

                    raw = (val == "1") ? "1" : "0";
                }

                info.Text = raw;
            }

            result.Add(info);
        }

        //======================
        // 7. ⭐ 寫回 docx
        //======================
        string newXml = xmlDoc.OuterXml;
        zip.UpdateEntry("word/document.xml", newXml, Encoding.UTF8);
        zip.Save();

        return result;
    }

    private string GetSqlTypeByControlType(string controlType)
    {
        string sqlType = "nvarchar(500)";

        if (controlType == "checkbox")
            sqlType = "nvarchar(2)";
        else if (controlType == "date")
            sqlType = "nvarchar(50)";
        else if (controlType == "dropdown" || controlType == "combobox")
            sqlType = "nvarchar(50)";

        return sqlType;
    }

    private DataRow GetFieldDef(SqlConnection conn, SqlTransaction trans,
                            string guid, int version, string tag)
    {
        string sql = @"
        SELECT * FROM 公文欄位定義表
        WHERE guid = @g AND 版本 = @v AND 項目代碼 = @t
    ";

        SqlDataAdapter da = new SqlDataAdapter(sql, conn);
        da.SelectCommand.Transaction = trans;
        da.SelectCommand.Parameters.AddWithValue("@g", guid);
        da.SelectCommand.Parameters.AddWithValue("@v", version);
        da.SelectCommand.Parameters.AddWithValue("@t", tag);

        DataTable dt = new DataTable();
        da.Fill(dt);
        return dt.Rows.Count > 0 ? dt.Rows[0] : null;
    }

    private void InsertFieldDef(SqlConnection conn, SqlTransaction trans,
                            string guid, int version,
                            string tag, string type, string editorUser)
    {
        string sql = @"
        INSERT INTO 公文欄位定義表
        (guid, 版本, 項目代碼, 項目類型, 項目名稱, 是否已刪除, 建立者, 修改者)
        VALUES (@g, @v, @t, @ty, NULL, 0, @u, @u)
    ";

        SqlCommand cmd = new SqlCommand(sql, conn, trans);
        cmd.Parameters.AddWithValue("@g", guid);
        cmd.Parameters.AddWithValue("@v", version);
        cmd.Parameters.AddWithValue("@t", tag);
        cmd.Parameters.AddWithValue("@ty", type);
        cmd.Parameters.AddWithValue("@u", editorUser);
        cmd.ExecuteNonQuery();
    }

    private void UpdateFieldType(SqlConnection conn, SqlTransaction trans,
                             string guid, int version, string tag,
                             string newType, string editorUser)
    {
        string sql = @"
        UPDATE 公文欄位定義表
        SET 項目類型 = @ty,
            修改時間 = GETDATE(),
            修改者 = @u
        WHERE guid = @g AND 版本 = @v AND 項目代碼 = @t
    ";

        SqlCommand cmd = new SqlCommand(sql, conn, trans);
        cmd.Parameters.AddWithValue("@g", guid);
        cmd.Parameters.AddWithValue("@v", version);
        cmd.Parameters.AddWithValue("@t", tag);
        cmd.Parameters.AddWithValue("@ty", newType);
        cmd.Parameters.AddWithValue("@u", editorUser);
        cmd.ExecuteNonQuery();
    }

    private void InsertFieldChangeLog(SqlConnection conn, SqlTransaction trans,
                                  string guid, int version, string tag,
                                  string changeType,
                                  string oldType, string newType,
                                  string oldName, string newName,
                                  string editorUser)
    {
        string sql = @"
        INSERT INTO 公文欄位異動log
        (guid, 版本, 異動類型, 項目代碼,
         原項目名稱, 新項目名稱,
         原項目類型, 新項目類型,
         建立者, 修改者)
        VALUES
        (@g, @v, @ct, @t,
         @oldName, @newName,
         @oldType, @newType,
         @u, @u)
    ";

        SqlCommand cmd = new SqlCommand(sql, conn, trans);
        cmd.Parameters.AddWithValue("@g", guid);
        cmd.Parameters.AddWithValue("@v", version);
        cmd.Parameters.AddWithValue("@ct", changeType);
        cmd.Parameters.AddWithValue("@t", tag);
        cmd.Parameters.AddWithValue("@oldName", (object)oldName ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@newName", (object)newName ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@oldType", (object)oldType ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@newType", (object)newType ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@u", editorUser);
        cmd.ExecuteNonQuery();
    }
}