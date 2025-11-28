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
using System.Drawing;
using Aspose.Words;

public partial class handler_SaveCallback : System.Web.UI.Page
{
    DocManage_DB dmdb = new DocManage_DB();
    FileTable_DB fdb = new FileTable_DB();
    DocSaveLog_DB dsldb = new DocSaveLog_DB();
    DocContent_DB dcdb = new DocContent_DB();
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
            string latestVersion = string.Empty;
            string key = string.Empty;
            string filepath = string.Empty;
            string fileUrl = string.Empty;
            string changesurl = string.Empty;
            string editorUser = "unknown";
            string filetype = "";
            string filecode = "";
            string fileName = "";
            string fileGuid = string.Empty;
            string fileNewGuid = Guid.NewGuid().ToString("N");
            string fileparentGuid = string.Empty;
            string PublicDocName = string.Empty;
            string PublicDocNo = string.Empty;
            string PublicDoemGuid = string.Empty;
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
                filetype = parts.Length > 3 ? parts[3] : "";
            }

            if(filetype == "demo")
            {
                filecode = "01";
            }
            else
            {
                filecode = "02";
            }

            File.AppendAllText(Server.MapPath("~/log-callback.txt"),
                DateTime.Now + "\nstatus=" + status.ToString() + "\nkey=" + key + "\nfileUrl=" + fileUrl + "\nfileName=" + fileName +
                "\neditorUser=333\nchange=" + changesurl + "\n\n");

            File.AppendAllText(Server.MapPath("~/log-callback.txt"),
                DateTime.Now + "\nfileGuid=" + fileGuid + "\nfileparentGuid=" + fileparentGuid + "\nPublicFileVersion=" + PublicFileVersion + "\n\n");
            #endregion

            if (filecode == "01")
            {
                #region 取得檔案

                //取得當前路徑以及檔案名稱
                dmdb._guid = fileGuid;
                DataTable dmdt = dmdb.GetDemoDataNoFile(filecode);
                if (dmdt.Rows.Count > 0)
                {
                    fileName = dmdt.Rows[0]["新檔名"].ToString().Trim() + dmdt.Rows[0]["附檔名"].ToString().Trim();
                    PublicDocName = dmdt.Rows[0]["表單名稱"].ToString().Trim();
                    PublicDocNo = dmdt.Rows[0]["表單編號"].ToString().Trim();
                    PublicFileName = dmdt.Rows[0]["原檔名"].ToString().Trim();
                    PublicFileNewName = dmdt.Rows[0]["新檔名"].ToString().Trim();
                    PublicFileExtension = dmdt.Rows[0]["附檔名"].ToString().Trim();
                    PublicFileSn = dmdt.Rows[0]["排序"].ToString().Trim();
                    filepath = UpLoadPath + "公文文件\\" + dmdt.Rows[0]["RelativePath"].ToString().Trim();
                }                  

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
                    fdb._表單名稱 = PublicDocName;
                    fdb._表單編號 = PublicDocNo;
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
                    //var controls = ParseContentControls(Path.Combine(filepath, PublicFileName + "_v" + latestVersion + PublicFileExtension));
                    var controls = ParseContentControls(oConn, myTrans, filepath + PublicFileFullName + PublicFileExtension, fileNewGuid);

                    foreach (var c in controls)
                    {
                        File.AppendAllText(Server.MapPath("~/log-callback.txt"),
                            "[" + DateTime.Now + "]\r\n" +
                            "Tag = " + c.Tag + "\r\n" +
                            "TypeCode = " + c.TypeCode + "\r\n" +
                            "Type = " + c.Type + "\r\n" +
                            "Value = " + c.Text + "\r\n" +
                            (c.Items.Count > 0 ? "Items = " + string.Join(",", c.Items) + "\r\n" : "") +
                            "\r\n"
                        );
                    }

                    // 2. 把 Word 裡的欄位狀況同步到 公文欄位定義表
                    foreach (var c in controls)
                    {
                        string tag = c.Tag;
                        string typecode = c.TypeCode;
                        string type = c.Type;

                        // ★ 直接新增欄位定義
                        InsertFieldDef(oConn, myTrans, fileNewGuid, nowVersion,
                                       tag, typecode, editorUser);

                        // ★ 寫 Add Log
                        InsertFieldChangeLog(oConn, myTrans, fileNewGuid, nowVersion,
                                             tag, "Add", typecode, typecode, null, null, editorUser);
                    }
                    #endregion

                    #region 儲存log
                    dsldb._guid = fileNewGuid;
                    dsldb._狀態 = status.ToString();
                    dsldb._類別 = "手動儲存";
                    dsldb._版本 = nowVersion;
                    dsldb._修改者 = editorUser;
                    dsldb._修改日期 = DateTime.Now;
                    dsldb._建立者 = editorUser;
                    dsldb._建立日期 = DateTime.Now;
                    dsldb.Insert_Trans(oConn, myTrans);
                    #endregion
                }
                #endregion

                #region 自動儲存
                if (status == 2 && !string.IsNullOrEmpty(fileUrl))
                {
                    #region 儲存檔案進資料夾

                    //取得目前最新版版本
                    string latestFileGuid = string.Empty;
                    fdb._父層guid = fileparentGuid;
                    fdb._檔案類型 = "01";
                    fdb._排序 = PublicFileSn;
                    DataTable mddt = fdb.GetSnMaxData();

                    if (mddt.Rows.Count > 0)
                    {
                        latestVersion = mddt.Rows[0]["版本"].ToString().Trim();
                        latestFileGuid = mddt.Rows[0]["guid"].ToString().Trim();
                    }
                    else
                    {
                        latestVersion = "1";
                    }

                    using (var client = new WebClient())
                    {
                        byte[] fileBytes = client.DownloadData(fileUrl + PublicFileNewName + PublicFileExtension);
                        PublicFileFullName = PublicFileName + "_v" + latestVersion;
                        string realfilepath = Path.Combine(filepath, PublicFileFullName + PublicFileExtension);

                        File.WriteAllBytes(realfilepath, fileBytes);
                    }
                    #endregion

                    #region 儲存log
                    dsldb._guid = latestFileGuid;
                    dsldb._狀態 = status.ToString();
                    dsldb._類別 = "自動儲存";
                    dsldb._版本 = latestVersion;
                    dsldb._修改者 = editorUser;
                    dsldb._修改日期 = DateTime.Now;
                    dsldb._建立者 = editorUser;
                    dsldb._建立日期 = DateTime.Now;
                    dsldb.Insert_Trans(oConn, myTrans);
                    #endregion

                    #region 儲存進指定表單資料表
                    //var controls = ParseContentControls(Path.Combine(filepath, PublicFileName + "_v" + latestVersion + PublicFileExtension));
                    var controls = ParseContentControls(oConn, myTrans, Path.Combine(filepath, PublicFileName + "_v" + latestVersion + PublicFileExtension), latestFileGuid);

                    // 把 Word 的 Tag 全部放入一個 Set
                    HashSet<string> wordTags = new HashSet<string>(controls.Select(c => c.Tag));

                    // 取得目前資料庫該版本所有欄位
                    string sqlGetDbFields =
                        @"SELECT * FROM 公文欄位定義表
      WHERE guid=@g AND 版本=@v";

                    SqlCommand cmdGet = new SqlCommand(sqlGetDbFields, oConn, myTrans);
                    cmdGet.Parameters.AddWithValue("@g", latestFileGuid);
                    cmdGet.Parameters.AddWithValue("@v", latestVersion);

                    DataTable dtDb = new DataTable();
                    new SqlDataAdapter(cmdGet).Fill(dtDb);

                    // 方便比對：DB 的 tag → row
                    Dictionary<string, DataRow> dbDic = dtDb.AsEnumerable()
                        .ToDictionary(r => r["項目代碼"].ToString(), r => r);


                    //---------------------------------------------------
                    // A. 新增 / 修改：Word 有 → DB 沒有 或 類型不同
                    //---------------------------------------------------
                    foreach (var c in controls)
                    {
                        string tag = c.Tag;
                        string type = c.Type;
                        string typeCode = c.TypeCode;

                        if (!dbDic.ContainsKey(tag))
                        {
                            // ★ 新增
                            InsertFieldDef(oConn, myTrans, latestFileGuid, latestVersion,
                                           tag, typeCode, editorUser);

                            InsertFieldChangeLog(oConn, myTrans, latestFileGuid, latestVersion,
                                                 tag,
                                                 "Add",
                                                 typeCode, typeCode,
                                                 null, null,
                                                 editorUser);
                        }
                        else
                        {
                            // 比對類型是否變更
                            var row = dbDic[tag];
                            string oldType = row["項目類型"].ToString();

                            if (oldType != type)
                            {
                                // ★ 修改
                                UpdateFieldType(oConn, myTrans, latestFileGuid, latestVersion,
                                                tag, typeCode, editorUser);

                                InsertFieldChangeLog(oConn, myTrans, latestFileGuid, latestVersion,
                                                     tag,
                                                     "Modify",
                                                     oldType, typeCode,
                                                     null, null,
                                                     editorUser);
                            }
                        }
                    }


                    //---------------------------------------------------
                    // B. 刪除：DB 有 → Word 沒有（標註是否已刪除=1）
                    //---------------------------------------------------
                    foreach (var dbTag in dbDic.Keys)
                    {
                        if (!wordTags.Contains(dbTag))
                        {
                            // ★ 標為刪除
                            string sqlDel =
                                @"UPDATE 公文欄位定義表
              SET 是否已刪除 = 1,
                  修改時間 = GETDATE(),
                  修改者 = @u
              WHERE guid=@g AND 版本=@v AND 項目代碼=@t";

                            SqlCommand cmdDel = new SqlCommand(sqlDel, oConn, myTrans);
                            cmdDel.Parameters.AddWithValue("@u", editorUser);
                            cmdDel.Parameters.AddWithValue("@g", latestFileGuid);
                            cmdDel.Parameters.AddWithValue("@v", latestVersion);
                            cmdDel.Parameters.AddWithValue("@t", dbTag);
                            cmdDel.ExecuteNonQuery();

                            // ★ 寫 Delete Log
                            string oldType = dbDic[dbTag]["項目類型"].ToString();

                            InsertFieldChangeLog(oConn, myTrans, latestFileGuid, latestVersion,
                                                 dbTag,
                                                 "Delete",
                                                 oldType, null,
                                                 dbDic[dbTag]["項目名稱"].ToString(),
                                                 null,
                                                 editorUser);
                        }
                    }
                    #endregion
                }
                #endregion
            }
            else
            {
                #region 取得檔案

                fdb._guid = fileGuid;
                fdb._檔案類型 = filecode;
                DataTable fdt = fdb.GetSnMaxData();

                if (fdt.Rows.Count > 0)
                {
                    fdb._guid = "";
                    fdb._父層guid = fdt.Rows[0]["父層guid"].ToString().Trim();
                    fdb._檔案類型 = "01";
                    fdb._排序 = fdt.Rows[0]["排序"].ToString().Trim();
                    DataTable fddb = fdb.GetSnMaxData();

                    if (fddb.Rows.Count > 0)
                    {
                        PublicDoemGuid = fddb.Rows[0]["guid"].ToString().Trim();
                    }

                    fileName = fdt.Rows[0]["新檔名"].ToString().Trim() + fdt.Rows[0]["附檔名"].ToString().Trim();
                    PublicDocName = fdt.Rows[0]["表單名稱"].ToString().Trim();
                    PublicDocNo = fdt.Rows[0]["表單編號"].ToString().Trim();
                    PublicFileName = fdt.Rows[0]["原檔名"].ToString().Trim();
                    PublicFileNewName = fdt.Rows[0]["新檔名"].ToString().Trim();
                    PublicFileExtension = fdt.Rows[0]["附檔名"].ToString().Trim();
                    PublicFileSn = fdt.Rows[0]["排序"].ToString().Trim();
                    PublicFileVersion = fdt.Rows[0]["版本"].ToString().Trim();
                    filepath = UpLoadPath + "公文\\" + fileGuid + "\\";
                }

                #endregion

                #region 手動儲存

                if (status == 6 && !string.IsNullOrEmpty(fileUrl))
                {
                    #region 儲存檔案進手動儲存資料夾

                    //取得目前最新版版本+1
                    nowVersion = (Convert.ToInt32(PublicFileVersion) + 1).ToString();

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
                    fdb._guid = fileGuid;
                    fdb._檔案類型 = filecode;
                    fdb._父層guid = fileparentGuid;
                    fdb._年度 = "";
                    fdb._表單名稱 = PublicDocName;
                    fdb._表單編號 = PublicDocNo;
                    fdb._原檔名 = PublicFileName;
                    fdb._新檔名 = PublicFileFullName;
                    fdb._附檔名 = PublicFileExtension;
                    fdb._排序 = PublicFileSn;
                    fdb._版本 = nowVersion;
                    fdb._檔案大小 = "";
                    fdb._修改者 = editorUser;
                    fdb._修改日期 = DateTime.Now;
                    fdb._建立者 = editorUser;
                    fdb._建立日期 = DateTime.Now;
                    fdb.UpdateFile_Trans(oConn, myTrans);
                    #endregion

                    #region 儲存進公文內容

                    var controls = ReadContentControls(filepath + PublicFileFullName + PublicFileExtension);

                    // 2. 把 Word 的控制項內容丟到 公文內容
                    foreach (var c in controls)
                    {
                        string tag = c.Tag;
                        string typecode = c.TypeCode;
                        string type = c.Type;
                        string content = c.Text;

                        // ★ 直接新增欄位定義
                        InsertFieldContent(oConn, myTrans, fileGuid, PublicDoemGuid, nowVersion,
                                       tag, typecode, content, editorUser);

                        // ★ 寫 Add Log
                        //InsertFieldChangeLog(oConn, myTrans, fileNewGuid, nowVersion,
                        //                     tag, "Add", typecode, typecode, null, null, editorUser);
                    }
                    #endregion

                    #region 儲存log
                    dsldb._guid = fileGuid;
                    dsldb._狀態 = status.ToString();
                    dsldb._類別 = "手動儲存";
                    dsldb._版本 = nowVersion;
                    dsldb._修改者 = editorUser;
                    dsldb._修改日期 = DateTime.Now;
                    dsldb._建立者 = editorUser;
                    dsldb._建立日期 = DateTime.Now;
                    dsldb.Insert_Trans(oConn, myTrans);
                    #endregion
                }

                #endregion

                #region 自動儲存

                if (status == 2 && !string.IsNullOrEmpty(fileUrl))
                {
                    #region 儲存檔案進資料夾

                    using (var client = new WebClient())
                    {
                        byte[] fileBytes = client.DownloadData(fileUrl + PublicFileNewName + PublicFileExtension);
                        PublicFileFullName = PublicFileName + "_v" + PublicFileVersion;
                        string realfilepath = Path.Combine(filepath, PublicFileFullName + PublicFileExtension);

                        File.WriteAllBytes(realfilepath, fileBytes);
                    }
                    #endregion

                    #region 儲存進公文內容

                    var controls = ReadContentControls(filepath + PublicFileFullName + PublicFileExtension);

                    // 2. 把 Word 的控制項內容丟到 公文內容
                    foreach (var c in controls)
                    {
                        string tag = c.Tag;
                        string typecode = c.TypeCode;
                        string type = c.Type;
                        string content = c.Text;

                        dcdb._guid = fileGuid;
                        dcdb._版本 = PublicFileVersion;
                        dcdb._項目代碼 = tag;

                        DataTable dcdt = dcdb.GetData_Trans(oConn, myTrans);

                        if (dcdt.Rows.Count > 0)
                        {
                            UpdateFieldContent(oConn, myTrans, fileGuid, PublicFileVersion,
                                           tag, content, editorUser);
                        }
                        else
                        {
                            // ★ 直接新增欄位定義
                            InsertFieldContent(oConn, myTrans, fileGuid, PublicDoemGuid, PublicFileVersion,
                                           tag, typecode, content, editorUser);
                        }                        

                        // ★ 寫 Add Log
                        //InsertFieldChangeLog(oConn, myTrans, fileNewGuid, nowVersion,
                        //                     tag, "Add", typecode, typecode, null, null, editorUser);
                    }
                    #endregion

                    #region 儲存log
                    dsldb._guid = fileGuid;
                    dsldb._狀態 = status.ToString();
                    dsldb._類別 = "自動儲存";
                    dsldb._版本 = PublicFileVersion;
                    dsldb._修改者 = editorUser;
                    dsldb._修改日期 = DateTime.Now;
                    dsldb._建立者 = editorUser;
                    dsldb._建立日期 = DateTime.Now;
                    dsldb.Insert_Trans(oConn, myTrans);
                    #endregion
                }

                #endregion
            }

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

            Common.InsertLogs(Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath),
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
        public string TypeCode { get; set; }
        public string Text { get; set; }
        public List<string> Items { get; set; }
    }

    private static bool IsPlaceholderText(string val)
    {
        if (string.IsNullOrEmpty(val)) return true;

        val = val.Trim();

        // 這裡可以依你實際 Word 範本再補
        string[] placeholders = new string[]
        {
        "在此輸入文字",
        "請在此輸入文字",
        "請輸入文字",
        "請輸入日期",
        "請輸入內容",
        "請選擇",
        "選擇項目",
        "Click here to enter text",
        "Click here",
        "Enter text"
        };

        foreach (string p in placeholders)
        {
            if (val == p)
                return true;
        }

        return false;
    }

    public List<ContentControlInfo> ParseContentControls(SqlConnection conn, SqlTransaction trans, string docxPath, string fileGuid)
    {
        List<ContentControlInfo> result = new List<ContentControlInfo>();

        //======================
        // 1. 開啟 DOCX
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
        // 2. Extract XML
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
        // ⭐ Step 0：從資料庫取得最大 ccIndex
        //=======================================================
        int dbMaxIndex = GetDbMaxCcIndex(conn, trans, fileGuid);  // 你要自己定義
        if (dbMaxIndex < 0) dbMaxIndex = 0;
        int ccIndex = dbMaxIndex + 1;

        HashSet<string> usedTags = new HashSet<string>();

        //=======================================================
        // ⭐ Step 1：逐一處理控制項
        //=======================================================
        foreach (XmlNode sdt in sdtList)
        {
            ContentControlInfo info = new ContentControlInfo();
            info.Items = new List<string>();

            XmlNode sdtPr = sdt.SelectSingleNode(".//w:sdtPr", nsmgr);

            //======================
            // 3. 讀取 tag
            //======================
            XmlNode tagNode = sdtPr.SelectSingleNode("./w:tag", nsmgr);
            string tagValue = "";

            if (tagNode != null)
            {
                XmlAttribute valAttr = tagNode.Attributes["w:val"];
                if (valAttr != null)
                    tagValue = valAttr.Value;
            }

            info.Tag = tagValue;

            bool needNewTag = false;

            //======================
            // 判斷是否需要新 tag
            //======================

            // ① 無 tag
            if (string.IsNullOrWhiteSpace(info.Tag))
            {
                needNewTag = true;
            }
            // ② XML 内重複（複製貼上）
            else if (usedTags.Contains(info.Tag))
            {
                needNewTag = true;
            }
            // ③ DB 裡是已刪除（不准復活）
            else if (DbTagIsDeleted(conn, trans, fileGuid, info.Tag))
            {
                needNewTag = true;
            }

            //======================
            // 4. 需要新 tag → 編新號碼
            //======================
            if (needNewTag)
            {
                string newTag = "cc_" + ccIndex.ToString("0000");
                ccIndex++;

                info.Tag = newTag;
                usedTags.Add(newTag);

                // 寫入 w:tag
                if (tagNode == null)
                {
                    XmlElement newTagElement = xmlDoc.CreateElement("w", "tag",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                    XmlAttribute newVal = xmlDoc.CreateAttribute("w", "val",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    newVal.Value = newTag;

                    newTagElement.Attributes.Append(newVal);
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
            else
            {
                usedTags.Add(info.Tag);
            }

            //======================
            // Step 4-2：更新 alias = tag
            //======================
            XmlNode aliasNode = sdtPr.SelectSingleNode("./w:alias", nsmgr);
            if (aliasNode == null)
            {
                XmlElement newAlias = xmlDoc.CreateElement("w", "alias",
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                XmlAttribute aliasVal = xmlDoc.CreateAttribute("w", "val",
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                aliasVal.Value = info.Tag;

                newAlias.Attributes.Append(aliasVal);
                sdtPr.AppendChild(newAlias);
            }
            else
            {
                XmlAttribute aliasValAttr = aliasNode.Attributes["w:val"];
                if (aliasValAttr == null)
                {
                    aliasValAttr = xmlDoc.CreateAttribute("w", "val",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    aliasNode.Attributes.Append(aliasValAttr);
                }

                aliasValAttr.Value = info.Tag;
            }

            //======================
            // 5. 判斷控制項類型
            //======================
            info.Type = "text";

            XmlNode checkboxNode = sdtPr.SelectSingleNode("./w:checkbox", nsmgr);
            if (checkboxNode == null)
                checkboxNode = sdtPr.SelectSingleNode("./w14:checkbox", nsmgr);

            if (checkboxNode != null)
                info.Type = "checkbox";
            else if (sdtPr.SelectSingleNode("./w:dropDownList", nsmgr) != null)
                info.Type = "dropdown";
            else if (sdtPr.SelectSingleNode("./w:comboBox", nsmgr) != null)
                info.Type = "combobox";
            else if (sdtPr.SelectSingleNode("./w:date", nsmgr) != null)
                info.Type = "date";

            string typeCode = "01";
            if (ControlTypeMap.ContainsKey(info.Type))
                typeCode = ControlTypeMap[info.Type];

            info.TypeCode = typeCode;

            //======================
            // 6. 讀取 Text
            //======================
            StringBuilder sb = new StringBuilder();
            XmlNodeList textNodes = sdt.SelectNodes(".//w:t", nsmgr);

            foreach (XmlNode t in textNodes)
            {
                if (t != null && t.InnerText != null)
                    sb.Append(t.InnerText);
            }

            info.Text = sb.ToString();

            // checkbox → 0/1
            if (info.Type == "checkbox")
            {
                string raw = "0";

                XmlNode checkedNode =
                    sdt.SelectSingleNode(".//w14:checkbox/w14:checked", nsmgr);

                if (checkedNode == null)
                    checkedNode = sdt.SelectSingleNode(".//w:checkbox/w:checked", nsmgr);

                string val = null;

                if (checkedNode != null)
                {
                    XmlAttribute v1 = checkedNode.Attributes["w14:val"];
                    if (v1 != null && v1.Value != null)
                        val = v1.Value;

                    if (val == null)
                    {
                        XmlAttribute v2 = checkedNode.Attributes["w:val"];
                        if (v2 != null && v2.Value != null)
                            val = v2.Value;
                    }
                }

                if (val == "1")
                    raw = "1";

                info.Text = raw;
            }

            result.Add(info);
        }

        //======================
        // 7. 寫回 docx
        //======================
        zip.UpdateEntry("word/document.xml", xmlDoc.OuterXml, Encoding.UTF8);
        zip.Save();

        return result;
    }

    public List<ContentControlInfo> ReadContentControls(string docxPath)
    {
        List<ContentControlInfo> result = new List<ContentControlInfo>();

        ZipFile zip = ZipFile.Read(docxPath);

        ZipEntry docEntry = zip["word/document.xml"];
        if (docEntry == null)
            return result;

        XmlDocument xmlDoc = new XmlDocument();
        using (MemoryStream ms = new MemoryStream())
        {
            docEntry.Extract(ms);
            ms.Position = 0;
            xmlDoc.Load(ms);
        }

        XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
        nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        // ⚠ 這裡很重要：w14 namespace 要用這個
        nsmgr.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        nsmgr.AddNamespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml");

        XmlNodeList sdtList = xmlDoc.SelectNodes("//w:sdt", nsmgr);

        foreach (XmlNode sdt in sdtList)
        {
            ContentControlInfo info = new ContentControlInfo();
            info.Items = new List<string>();

            XmlNode sdtPr = sdt.SelectSingleNode("./w:sdtPr", nsmgr);
            if (sdtPr == null)
                continue;

            //======================
            // Tag
            //======================
            string tag = "";
            XmlNode tagNode = sdtPr.SelectSingleNode("./w:tag", nsmgr);
            if (tagNode != null)
            {
                XmlAttribute v = tagNode.Attributes["w:val"];
                if (v != null) tag = v.Value;
            }
            info.Tag = tag;

            //======================
            // Type 判斷
            //======================
            string type = "text";

            XmlNode checkBoxNode = sdtPr.SelectSingleNode("./w14:checkbox", nsmgr);
            if (checkBoxNode == null)
                checkBoxNode = sdtPr.SelectSingleNode("./w:checkbox", nsmgr);

            if (checkBoxNode != null)
            {
                type = "checkbox";
            }
            else if (sdtPr.SelectSingleNode("./w:comboBox", nsmgr) != null)
            {
                type = "combobox";
            }
            else if (sdtPr.SelectSingleNode("./w:dropDownList", nsmgr) != null)
            {
                type = "dropdown";
            }
            else if (sdtPr.SelectSingleNode("./w:date", nsmgr) != null)
            {
                type = "date";
            }

            info.Type = type;

            //======================
            // TypeCode
            //======================
            string typeCode;
            if (!ControlTypeMap.TryGetValue(type, out typeCode))
                typeCode = "01";
            info.TypeCode = typeCode;

            //======================
            // dropdown / combobox → 取得 Items
            //======================
            if (type == "dropdown" || type == "combobox")
            {
                XmlNode listNode = sdtPr.SelectSingleNode("./w:dropDownList", nsmgr);
                if (listNode == null)
                    listNode = sdtPr.SelectSingleNode("./w:comboBox", nsmgr);

                if (listNode != null)
                {
                    XmlNodeList itemNodes = listNode.SelectNodes("./w:listItem", nsmgr);
                    foreach (XmlNode item in itemNodes)
                    {
                        string text = "";
                        string value = "";

                        XmlAttribute attrDisplay = item.Attributes["w:displayText"];
                        if (attrDisplay != null && !string.IsNullOrEmpty(attrDisplay.Value))
                            text = attrDisplay.Value;

                        XmlAttribute attrValue = item.Attributes["w:value"];
                        if (attrValue != null && !string.IsNullOrEmpty(attrValue.Value))
                            value = attrValue.Value;

                        string final = "";
                        if (!string.IsNullOrEmpty(text))
                            final = text;
                        else
                            final = value;

                        if (!IsPlaceholderText(final))
                            info.Items.Add(final);
                    }
                }
            }

            //======================
            // Text 內容處理
            //======================

            if (type == "checkbox")
            {
                // ⭐ 完全用 checked 狀態決定 0 / 1，不看 "☐" / "☑"
                string raw = "0";

                XmlNode checkedNode =
                    sdt.SelectSingleNode(".//w14:checkbox/w14:checked", nsmgr);

                if (checkedNode == null)
                    checkedNode = sdt.SelectSingleNode(".//w:checkbox/w:checked", nsmgr);

                string val = null;

                if (checkedNode != null)
                {
                    XmlAttribute v1 = checkedNode.Attributes["w14:val"];
                    if (v1 != null && v1.Value != null)
                        val = v1.Value;

                    if (val == null)
                    {
                        XmlAttribute v2 = checkedNode.Attributes["w:val"];
                        if (v2 != null && v2.Value != null)
                            val = v2.Value;
                    }
                }

                if (val == "1")
                    raw = "1";

                info.Text = raw;
            }
            else
            {
                // text / dropdown / combobox / date → 把 placeholder 過濾掉
                StringBuilder sb = new StringBuilder();
                XmlNodeList textNodes = sdt.SelectNodes(".//w:t", nsmgr);

                foreach (XmlNode n in textNodes)
                {
                    string val = n.InnerText;

                    if (IsPlaceholderText(val))
                        continue;

                    sb.Append(val);
                }

                info.Text = sb.ToString().Trim();
            }

            result.Add(info);
        }

        return result;
    }

    private static Dictionary<string, string> ControlTypeMap = new Dictionary<string, string>
    {
        { "text", "01" },
        { "checkbox", "02" },
        { "dropdown", "03" },
        { "combobox", "04" },
        { "date", "05" }
    };

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

    private int GetDbMaxCcIndex(SqlConnection conn, SqlTransaction trans, string fileGuid)
    {
        string sql = @"
        SELECT MAX(TRY_CONVERT(INT, SUBSTRING(項目代碼, 4, 4))) AS maxSn
        FROM 公文欄位定義表
        WHERE guid = @fileGuid
    ";

        SqlDataAdapter da = new SqlDataAdapter(sql, conn);
        da.SelectCommand.Transaction = trans;
        da.SelectCommand.Parameters.AddWithValue("@fileGuid", fileGuid);

        DataTable dt = new DataTable();
        da.Fill(dt);

        if (dt.Rows.Count == 0 || dt.Rows[0]["maxSn"] == DBNull.Value)
            return 0;

        int maxSn;
        if (!int.TryParse(dt.Rows[0]["maxSn"].ToString(), out maxSn))
            return 0;

        return maxSn;


    }

    private bool DbTagIsDeleted(SqlConnection conn, SqlTransaction trans, string fileGuid, string tag)
    {
        string sql = @"
        SELECT 是否已刪除 
        FROM 公文欄位定義表 
        WHERE guid=@fileGuid AND 項目代碼=@tag
    ";

        SqlDataAdapter da = new SqlDataAdapter(sql, conn);
        da.SelectCommand.Transaction = trans;
        da.SelectCommand.Parameters.AddWithValue("@fileGuid", fileGuid);
        da.SelectCommand.Parameters.AddWithValue("@tag", tag);

        DataTable dt = new DataTable();
        da.Fill(dt);

        if (dt.Rows.Count == 0)
            return false; // DB 裡沒有 → 視為未刪除（新控制項）

        return dt.Rows[0]["是否已刪除"].ToString() == "1";
    }

    private void InsertFieldContent(SqlConnection conn, SqlTransaction trans,
                            string guid, string parentguid, string version,
                            string tag, string type, string content, string editorUser)
    {
        string sql = @"
        INSERT INTO 公文內容
        (guid, 範本guid, 版本, 項目代碼, 項目類型, 項目內容, 建立者, 修改者)
        VALUES (@g, @pg, @v, @t, @ty, @ct, @u, @u)
    ";

        SqlCommand cmd = new SqlCommand(sql, conn, trans);
        cmd.Parameters.AddWithValue("@g", guid);
        cmd.Parameters.AddWithValue("@pg", parentguid);
        cmd.Parameters.AddWithValue("@v", version);
        cmd.Parameters.AddWithValue("@t", tag);
        cmd.Parameters.AddWithValue("@ty", type);
        cmd.Parameters.AddWithValue("@ct", content);
        cmd.Parameters.AddWithValue("@u", editorUser);
        cmd.ExecuteNonQuery();
    }

    private void UpdateFieldContent(SqlConnection conn, SqlTransaction trans,
                            string guid, string version,string tag, string content, string editorUser)
    {
        string sql = @"
        UPDATE 公文內容 SET
            項目內容 = @ct,
            修改時間 = GETDATE(),
            修改者 = @u
        WHERE guid = @g AND 版本 = @v AND 項目代碼 = @t
    ";

        SqlCommand cmd = new SqlCommand(sql, conn, trans);
        cmd.Parameters.AddWithValue("@g", guid);
        cmd.Parameters.AddWithValue("@v", version);
        cmd.Parameters.AddWithValue("@t", tag);
        cmd.Parameters.AddWithValue("@ct", content);
        cmd.Parameters.AddWithValue("@u", editorUser);
        cmd.ExecuteNonQuery();
    }

    private void InsertFieldDef(SqlConnection conn, SqlTransaction trans,
                            string guid, string version,
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
                             string guid, string version, string tag,
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
                                  string guid, string version, string tag,
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