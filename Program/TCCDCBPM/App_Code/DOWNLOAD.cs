using System;
using System.Web;
using System.Configuration;
using System.Net;
using System.Data;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Web.UI.WebControls;
using System.Drawing;
using System.IdentityModel.Protocols.WSTrust;

namespace ED.HR.DOWNLOAD.WebForm
{
    public partial class DownloadImage : System.Web.UI.Page
    {
        string OrgName = string.Empty;
        string NewName = string.Empty;
        bool isWord = false;
        string UpLoadPath = ConfigurationManager.AppSettings["UploadFileRootDir"];
        FileTable_DB Fdb = new FileTable_DB();
        DocManage_DB DMdb = new DocManage_DB();
        //File_DB fdb = new File_DB();
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                string guid = Common.FilterCheckMarxString(Request.QueryString["guid"]);
                string category = Common.FilterCheckMarxString(Request.QueryString["category"]);
                string sn = Common.FilterCheckMarxString(Request.QueryString["sn"]);
                string dirPath = string.Empty;
                DataTable dt = new DataTable();

                switch (category)
                {
                    case "Demo":
                        dirPath = "公文文件\\";
                        DMdb._guid = guid;
                        dt = DMdb.GetDemoData();

                        if (dt.Rows.Count > 0)
                        {
                            dirPath += (dt.Rows[0]["RelativePath"].ToString().Trim());
                        }

                        Fdb._guid = guid;
                        DataTable fdt = Fdb.GetSnMaxData();

                        if(fdt.Rows.Count > 0)
                        {
                            NewName = fdt.Rows[0]["新檔名"].ToString().Trim();
                        }

                        isWord = true;
                        break;
                    case "File":
                        dirPath = "公文\\";
                        Fdb._guid = guid;
                        dt = Fdb.GetSnMaxData();

                        if (dt.Rows.Count > 0)
                        {
                            dirPath += dt.Rows[0]["guid"].ToString().Trim() + "\\" + dt.Rows[0]["新檔名"].ToString().Trim() + dt.Rows[0]["附檔名"].ToString().Trim();
                        }

                        isWord = true;
                        break;
                }

                //原檔名
                //OrgName = Common.FilterCheckMarxString(Request.QueryString["v"]);
                string finalPath = Path.Combine(UpLoadPath, dirPath);

                File.AppendAllText(Server.MapPath("~/log-callback.txt"), DateTime.Now + "\nfinalPath=" + finalPath + "\n\n");

                // 附件目錄
                //if (isWord == true)
                //{
                //    finalPath = dirPath;
                //}
                //else
                //{
                //    finalPath = UpLoadPath + dirPath;
                //}

                //DirectoryInfo dir = new DirectoryInfo(finalPath);

                //列舉全部檔案再比對檔名
                //string FileName = NewName;

                FileInfo file = new FileInfo(finalPath);

                // 判斷檔案是否存在
                if (file != null && file.Exists)
                {
                    if (isWord == true)
                    {
                        StreamDocxForOnlyOffice(file);
                    }
                    else
                    {
                        Download(file);
                    }
                }
                else
                {
                    throw new Exception("File not exist");
                }
            }
            catch (Exception ex)
            {
                string logPath = @"D:\log.txt";
                File.AppendAllText(logPath, string.Format("[{0}] 錯誤：{1}\n", DateTime.Now, ex.Message));

                Common.InsertLogs(Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath),
                    System.Reflection.MethodBase.GetCurrentMethod().Name, "錯誤：" + ex.Message + "\r\n" + ex.StackTrace);

                Response.StatusCode = 500;
                Response.Write("Download error");
                Response.End();
            }
        }


        private void Download(FileInfo DownloadFile)
        {
            Response.Clear();
            Response.ClearHeaders();
            Response.Buffer = false;
            Response.ContentType = getMineType(DownloadFile.Extension);
            string DownloadName = (OrgName == "") ? DownloadFile.Name : OrgName;
            Response.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode(DownloadName, System.Text.Encoding.UTF8));
            Response.AppendHeader("Content-Length", DownloadFile.Length.ToString());
            Response.HeaderEncoding = System.Text.Encoding.GetEncoding("Big5");
            Response.WriteFile(DownloadFile.FullName);
            Response.Flush();
            Response.End();
        }

        private void StreamDocxForOnlyOffice(FileInfo docFile)
        {
            File.AppendAllText(@"D:\log.txt", string.Format("[{0}] 觸發 OnlyOffice 檔案串流：{1}\n", DateTime.Now, docFile.FullName));
            string encodedFileName = Uri.EscapeDataString(docFile.Name);

            HttpResponse response = HttpContext.Current.Response;
            response.Clear();
            response.ClearHeaders();
            response.ClearContent();
            response.Buffer = false;

            response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            response.AddHeader("Content-Disposition", "inline; filename=\"test1.docx\""); // 避免 filename*= 問題
            response.AddHeader("Content-Length", docFile.Length.ToString());
            response.AddHeader("Access-Control-Allow-Origin", "*");
            response.AppendHeader("Content-Encoding", "identity");

            // 建議使用 TransmitFile 比 WriteFile 穩定
            response.TransmitFile(docFile.FullName);
            response.Flush();

            // 正確結束流程，不要用 End()
            HttpContext.Current.ApplicationInstance.CompleteRequest();
        }

        #region 傳回 ContentType
        /// <summary>
        /// 傳回 ContentType
        /// </summary>
        public static string getMineType(string FileExtension)
        {
            Microsoft.Win32.RegistryKey rk = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(FileExtension);
            if (rk != null && rk.GetValue("Content Type") != null)
                return rk.GetValue("Content Type").ToString();
            else
                return "application/octet-stream";
        }
        #endregion
    }
}