using System;
using System.Collections.Generic;
using System.Web;
using System.Configuration;
using System.Text;
using System.Text.RegularExpressions;
using System.Data;
using System.Security.Cryptography;
using System.IO;
using System.Data.SqlClient;
using System.Net;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Security.Principal;
using System.Runtime.Remoting.Contexts;
//using FlexCel.Core;
//using FlexCel.XlsAdapter;

/// <summary>
/// Common 的摘要描述
/// </summary>
public class Common
{

    #region Get IPv4 Adress
    public static string GetIP4Address()
    {
        System.Web.HttpContext context = System.Web.HttpContext.Current;
        string sIPAddress = context.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
        if (string.IsNullOrEmpty(sIPAddress))
        {
            string[] ipstr = context.Request.ServerVariables["REMOTE_ADDR"].Split(':');
            if (ipstr[0].Trim() != "")
                return context.Request.ServerVariables["REMOTE_ADDR"];
            else
                return "LOCAL-Name：" + Environment.MachineName;
        }
        else
        {
            string[] ipArray = sIPAddress.Split(new Char[] { ',' });
            return ipArray[0];
        }
    }
    #endregion

    #region 加解密
    /// <summary>
    /// 加密
    /// </summary>
    public static string Encrypt(string strSource)
    {
        //把字符串放到byte數組中  
        byte[] bytIn = Encoding.Default.GetBytes(strSource);
        //建立加密對象的密鑰和偏移量          
        byte[] iv = { 102, 16, 93, 156, 78, 4, 218, 32 };//定義偏移量  
        byte[] key = { 55, 103, 246, 79, 36, 99, 167, 3 };//定義密鑰
        //實例DES加密類  
        DESCryptoServiceProvider mobjCryptoService = new DESCryptoServiceProvider();
        mobjCryptoService.Key = iv;
        mobjCryptoService.IV = key;
        ICryptoTransform encrypto = mobjCryptoService.CreateEncryptor();
        //實例MemoryStream流加密密文件  
        MemoryStream ms = new MemoryStream();
        CryptoStream cs = new CryptoStream(ms, encrypto, CryptoStreamMode.Write);
        cs.Write(bytIn, 0, bytIn.Length);
        cs.FlushFinalBlock();
        return Convert.ToBase64String(ms.ToArray());
    }

    /// <summary>
    /// 解密
    /// </summary>
    public static string Decrypt(string Source)
    {
        string str = "";
        try
        {
            //將解密字符串轉換成字節數組 
            byte[] bytIn = System.Convert.FromBase64String(Source);
            //給出解密的密鑰和偏移量，密鑰和偏移量必須與加密時的密鑰和偏移量相同 
            byte[] iv = { 102, 16, 93, 156, 78, 4, 218, 32 };//定義偏移量  
            byte[] key = { 55, 103, 246, 79, 36, 99, 167, 3 };//定義密鑰 
            DESCryptoServiceProvider mobjCryptoService = new DESCryptoServiceProvider();
            mobjCryptoService.Key = iv;
            mobjCryptoService.IV = key;
            //實例流進行解密  
            System.IO.MemoryStream ms = new System.IO.MemoryStream(bytIn, 0, bytIn.Length);
            ICryptoTransform encrypto = mobjCryptoService.CreateDecryptor();
            CryptoStream cs = new CryptoStream(ms, encrypto, CryptoStreamMode.Read);
            StreamReader strd = new StreamReader(cs, Encoding.Default);
            str = strd.ReadToEnd();
        }
        catch
        {

        }
        return str;
    }

    /// <summary>
    /// 目前可通過掃描的項目_轉回半形
    /// </summary>
    /// <param name="strOri">參數</param>
    /// <returns></returns>
    public static string DoReplaceBack(string strOri)
    {
        if (strOri.Length > 0)
        {
            strOri = strOri.Replace("︱", "|");
            strOri = strOri.Replace("＆", "&");
            //strOri = strOri.Replace(";", "；");
            strOri = strOri.Replace("＄", "$");
            strOri = strOri.Replace("％", "%");
            strOri = strOri.Replace("＠", "@");
            strOri = strOri.Replace("’", "'");
            strOri = strOri.Replace("＜", "<");
            //strOri = strOri.Replace("(", "（");
            //strOri = strOri.Replace("\"", "＂");
            strOri = strOri.Replace("＞", ">");
            //strOri = strOri.Replace(")", "）");
            strOri = strOri.Replace("＋", "+");
            strOri = strOri.Replace("＃", "#");
            //strOri = strOri.Replace(" CR ", "ＣＲ");
            // strOri = strOri.Replace(" LF ", "ＬＦ");
            //strOri = strOri.Replace("\\", "＼");
            //strOri = strOri.Replace("&lt", "＆lt");
            // strOri = strOri.Replace("&gt", "＆gt");

            //如果有連續兩個以上的"-"，則將所有的"-"變成全型"－"
            if (strOri.IndexOf("－－") > -1)
                strOri = strOri.Replace("－", "-");

        }

        return strOri;
    }

    /// <summary>
    /// 目前可通過掃描的項目
    /// </summary>
    /// <param name="strOri">參數</param>
    /// <returns></returns>
    public static string DoReplace(string strOri)
    {
        if (strOri.Length > 0)
        {
            strOri = strOri.Replace("|", "︱");
            strOri = strOri.Replace("&", "＆");
            //strOri = strOri.Replace(";", "；");
            strOri = strOri.Replace("$", "＄");
            strOri = strOri.Replace("%", "％");
            strOri = strOri.Replace("@", "＠");
            strOri = strOri.Replace("'", "’");
            strOri = strOri.Replace("<", "＜");
            //strOri = strOri.Replace("(", "（");
            //strOri = strOri.Replace("\"", "＂");
            strOri = strOri.Replace(">", "＞");
            //strOri = strOri.Replace(")", "）");
            strOri = strOri.Replace("+", "＋");
            strOri = strOri.Replace("#", "＃");
            //strOri = strOri.Replace(" CR ", "ＣＲ");
            // strOri = strOri.Replace(" LF ", "ＬＦ");
            //strOri = strOri.Replace("\\", "＼");
            //strOri = strOri.Replace("&lt", "＆lt");
            // strOri = strOri.Replace("&gt", "＆gt");

            //如果有連續兩個以上的"-"，則將所有的"-"變成全型"－"
            if (strOri.IndexOf("--") > -1)
                strOri = strOri.Replace("-", "－");

        }

        return strOri;
    }

    /// <summary>
    /// SHA1加密
    /// </summary>
    public static string sha1en(string str)
    {
        string enCodeString;
        SHA1CryptoServiceProvider sha1en = new SHA1CryptoServiceProvider();
        enCodeString = BitConverter.ToString(sha1en.ComputeHash(UTF8Encoding.Default.GetBytes(str)), 4, 8);
        enCodeString = enCodeString.Replace("-", "");
        return enCodeString;
    }

    /// <summary>
    /// Base64 加密
    /// </summary>
    public static string ToBase64String(string str)
    {
        return Convert.ToBase64String(Encoding.UTF8.GetBytes(str));
    }

    /// <summary>
    /// Base64 解密
    /// </summary>
    public static string FromBase64String(string str)
    {
        return Encoding.UTF8.GetString(Convert.FromBase64String(str));
    }
    #endregion

    #region 匯出EXCEL
    //public static void ExcuteExcel(ExcelFile CelFile, string ShowfileName)
    //{
    //	HttpContext.Current.Response.ContentType = "application/ms-excel";

    //	string fileName = HttpContext.Current.Server.UrlPathEncode(ShowfileName);
    //	string strContentDisposition = String.Format("{0}; filename=\"{1}\"", "attachment", fileName);
    //	HttpContext.Current.Response.AddHeader("Content-Disposition", strContentDisposition);
    //	MemoryStream ms = new MemoryStream();
    //	CelFile.Save(ms, TFileFormats.Xlsx);
    //	ms.Position = 0;

    //	BinaryReader br = new BinaryReader(ms);
    //	BinaryWriter bw = new BinaryWriter(HttpContext.Current.Response.OutputStream);

    //	for (int i = 0; i < ms.Length; i++)
    //	{
    //		bw.Write(br.ReadByte());
    //	}
    //	bw.Close();
    //	ms.Close();

    //	bw = null;
    //	ms = null;

    //	HttpContext.Current.Response.OutputStream.Flush();
    //	HttpContext.Current.Response.OutputStream.Close();
    //	return;
    //}
    #endregion

    #region sqlInjection
    /// <summary>
    /// 檢查特殊字元
    /// </summary>
    /// <param name="checkValue">欲檢查的值</param>
    /// <returns></returns>
    public static bool CheckSQLInjection(string checkValue)
    {
        //「%27」:「'」(單引號)
        //「%2B」:「+」(加號)
        //「%3D」:「=」(等號)
        //「%7C」:「|」(｜)
        //「ALERT(」
        //「--」
        //「%2F*」:「/*」
        //「*%2F」:「*/」
        //「%26」:「&」
        //「%40」:「@」
        //「%25」:「%」
        //「%3B」:「;」
        //「%24」:「$」
        //「%26」:「*」
        //「%22」:「"」
        //「%2C」:「,」
        //「%2f」:「/」
        //「%5c」:「\」
        //「%22」:「"」
        //「%3C」:「<」
        //「%3E」:「>」

        if (checkValue.Length > 0 && (checkValue.ToUpper().IndexOf("%27") >= 0 || checkValue.ToUpper().IndexOf("%2B") >= 0
          || checkValue.ToUpper().IndexOf("'") >= 0) || checkValue.ToUpper().IndexOf("ALERT(") >= 0
          || checkValue.ToUpper().IndexOf("%3C") >= 0 || checkValue.ToUpper().IndexOf("%3E") >= 0
          || checkValue.ToUpper().IndexOf("%3D") >= 0 || checkValue.ToUpper().IndexOf("=") >= 0
          || checkValue.ToUpper().IndexOf("--") >= 0 || checkValue.ToUpper().IndexOf("%7C") >= 0
          || checkValue.ToUpper().IndexOf("%2F*") >= 0 || checkValue.ToUpper().IndexOf("*%2F") >= 0
          || checkValue.ToUpper().IndexOf("%26") >= 0
          || checkValue.ToUpper().IndexOf("%25") >= 0 || checkValue.ToUpper().IndexOf("%3B") >= 0
          || checkValue.ToUpper().IndexOf("%24") >= 0 || checkValue.ToUpper().IndexOf("*") >= 0
          || checkValue.ToUpper().IndexOf("%22") >= 0 || checkValue.ToUpper().IndexOf("%2C") >= 0
          || checkValue.ToUpper().IndexOf("%2F") >= 0 || checkValue.ToUpper().IndexOf("%5C") >= 0
          || checkValue.ToUpper().IndexOf("%40") >= 0
          || checkValue.ToUpper().IndexOf("../") >= 0 || checkValue.ToUpper().IndexOf("%") >= 0 || checkValue.ToUpper().IndexOf("@") >= 0
          || checkValue.ToUpper().IndexOf("&") >= 0 || checkValue.ToUpper().IndexOf("..\\") >= 0 || checkValue.ToUpper().IndexOf("$") >= 0
          || checkValue.ToUpper().IndexOf("?") >= 0
          )
        {
            return false;
        }
        else
        {
            return true;
        }
    }
    #endregion

    #region 檢查參數
    public void CheckParameters(System.Data.SqlClient.SqlCommand oCmd)
    {
        //檢查危險字元
        for (int i = 0; i < oCmd.Parameters.Count; i++)
        {
            if (!CheckSQLInjection(oCmd.Parameters[i].Value.ToString()))
            {
                //throw new Exception("危險字元");
                //導引至錯誤網頁
                System.Web.HttpContext.Current.Response.Redirect("Error.aspx?err=par");

            }
        }
    }
    #endregion

    #region 清除html...
    /// <summary>
    /// 輸入html後刪除html標簽...
    /// </summary>
    public static string NoHTML(string Htmlstring)
    {
        //删除脚本
        Htmlstring = Htmlstring.Replace("\r\n", "");
        Htmlstring = Regex.Replace(Htmlstring, @"<script.*?</script>", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"<style.*?</style>", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"<.*?>", "", RegexOptions.IgnoreCase);
        //删除HTML
        Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"-->", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"<!--.*", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(quot|#34);", "\"", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(nbsp|#160);", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&#(\d+);", "", RegexOptions.IgnoreCase);
        Htmlstring = Htmlstring.Replace("<", "");
        Htmlstring = Htmlstring.Replace(">", "");
        Htmlstring = Htmlstring.Replace("\r\n", "");
        Htmlstring = HttpContext.Current.Server.HtmlEncode(Htmlstring).Trim();
        return Htmlstring;
    }
    #endregion

    #region 過濾 CheckMarx 用
    /// <summary>
    /// CheckMarx 過濾
    /// </summary>
    public static string FilterCheckMarxString(string str)
    {
        string rVal = string.Empty;
        rVal = HttpContext.Current.Server.HtmlEncode(str);
        rVal = HttpContext.Current.Server.HtmlDecode(rVal);
        return rVal;
    }
    #endregion

    #region 新增Logs
    public static void InsertLogsTran(SqlConnection oConn, SqlTransaction oTran, string pageName, string tableName, string content)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(@"insert into Logs(  
PageName,
TableName,
LogContent,
InsertTime ) values ( 
@PageName,
@TableName,
@LogContent,
@InsertTime  
) ");
        SqlCommand oCmd = oConn.CreateCommand();
        oCmd.CommandText = sb.ToString();

        oCmd.Parameters.AddWithValue("@PageName", pageName);
        oCmd.Parameters.AddWithValue("@TableName", tableName);
        oCmd.Parameters.AddWithValue("@LogContent", content);
        oCmd.Parameters.AddWithValue("@InsertTime", DateTime.Now);

        oCmd.Transaction = oTran;
        oCmd.ExecuteNonQuery();
    }

    public static void InsertLogs(string pageName, string tableName, string content)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
insert into Logs(  
PageName,
TableName,
LogContent,
InsertTime ) values ( 
@PageName,
@TableName,
@LogContent,
@InsertTime  
) ";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@PageName", pageName);
        oCmd.Parameters.AddWithValue("@TableName", tableName);
        oCmd.Parameters.AddWithValue("@LogContent", content);
        oCmd.Parameters.AddWithValue("@InsertTime", DateTime.Now);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    #endregion
}


#region 會員登入使用
/// <summary>
/// MbrAccount 的摘要描述。
/// </summary>
public class Account
{
    public static void ExecSignIn(string account, string pw)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@" select * from 會員檔 where 使用者帳號=@M_Account and 使用者密碼=@M_Pwd and 資料狀態='A' ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oCmd.Parameters.AddWithValue("@M_Account", account);
        oCmd.Parameters.AddWithValue("@M_Pwd", pw);
        oda.Fill(dt);

        if (dt.Rows.Count > 0)
        {
            //LogInfo.id = dt.Rows[0]["id"].ToString();
            //LogInfo.mGuid = dt.Rows[0]["guid"].ToString();
            //LogInfo.companyGuid = dt.Rows[0]["業者guid"].ToString();
            //LogInfo.account = dt.Rows[0]["使用者帳號"].ToString();
            //LogInfo.name = dt.Rows[0]["姓名"].ToString();
            //LogInfo.tel = dt.Rows[0]["電話"].ToString();
            //LogInfo.email = dt.Rows[0]["mail"].ToString();
            //LogInfo.orgcd = dt.Rows[0]["單位"].ToString();
            //LogInfo.orgname = dt.Rows[0]["部門代號"].ToString();
            //LogInfo.competence = dt.Rows[0]["帳號類別"].ToString();
            //LogInfo.user = dt.Rows[0]["網站類別"].ToString();
        }
        else
        {
            throw new ApplicationException("無效的帳號或密碼");
        }
    }

    public static DataTable IfAccountPasswordEqual(string mGuid)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@" 
declare @M_Account nvarchar(200)
declare @M_Pwd nvarchar(200)

select @M_Account=使用者帳號 from 會員檔 where guid=@mGuid
select @M_Pwd=使用者密碼 from 會員檔 where guid=@mGuid


if(@M_Account = @M_Pwd)
    begin
        select 'Y' as ifEqual
    end
else
    begin
        select 'N' as ifEqual
    end
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();

        oCmd.Parameters.AddWithValue("@mGuid", mGuid);
        oda.Fill(dt);

        return dt;
    }
}
#endregion


#region JavaScript Alert
public class JavaScript
{
    /// <summary>
    /// AlertMessage
    /// </summary>
    public static void AlertMessage(System.Web.UI.Page objPage, string strMessage)
    {
        strMessage = strMessage.Replace("\r\n", "\\r");
        StringBuilder sb = new StringBuilder();
        sb.AppendFormat(@"<Script language=""javascript"" type=""text/javascript"">");
        sb.AppendFormat(@"alert(""{0}"");", strMessage);
        sb.AppendFormat(@"</Script>");

        //objPage.ClientScript.RegisterClientScriptBlock(objPage.GetType(), "", strJS, false);
        objPage.ClientScript.RegisterStartupScript(objPage.GetType(), "", sb.ToString(), false);
    }

    /// <summary>
    /// AlertMessageClose
    /// </summary>
    public static void AlertMessageClose(System.Web.UI.Page objPage, string strMessage)
    {
        string strJS = "";
        strMessage = strMessage.Replace("\r\n", "\\r");
        strJS = @"<Script language='javascript' type='text/javascript' >";
        strJS += "alert('" + strMessage + "');";
        strJS += "window.close();";
        strJS += "</Script>";
        //objPage.ClientScript.RegisterClientScriptBlock(objPage.GetType(), "", strJS, false);
        objPage.ClientScript.RegisterStartupScript(objPage.GetType(), "", strJS, false);
    }

    /// <summary>
    /// AlertMessageRedirect
    /// </summary>
    public static void AlertMessageRedirect(System.Web.UI.Page objPage, string strMessage, string strRedirectPage)
    {
        AlertMessageRedirect(objPage, strMessage, strRedirectPage, false);
    }

    public static void AlertMessageRedirect(System.Web.UI.Page objPage, string strMessage, string strRedirectPage, bool IsDisplayData)
    {
        string strJS = "";
        strMessage = strMessage.Replace("\r\n", "\\r");
        strJS = @"<Script language='javascript' type='text/javascript'>";
        strJS += "alert('" + strMessage + "');";
        strJS += "window.location ='" + strRedirectPage + "'; ";
        strJS += "</Script>";

        if (IsDisplayData)
            objPage.ClientScript.RegisterStartupScript(objPage.GetType(), "", strJS, false);
        else
            objPage.ClientScript.RegisterClientScriptBlock(objPage.GetType(), "", strJS, false);
    }
}
#endregion

#region 爬網頁 Html (類似全文檢索)
public class CaptureURL
{
    public string Capture(string url)
    {
        try
        {
            string strHTML = string.Empty;

            if (url.IndexOf("https://") < 0)
            {
                string E = System.IO.Path.GetExtension(url);

                if (!E.Trim().ToLower().Equals(".html") && !string.IsNullOrEmpty(E.Trim()))
                {
                    string param = "hl=zh-CN&newwindow=1";

                    byte[] bs = System.Text.Encoding.ASCII.GetBytes(param);

                    HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);

                    req.Method = "POST";

                    req.ContentType = "application/x-www-form-urlencoded";

                    req.ContentLength = bs.Length;


                    using (Stream reqStream = req.GetRequestStream())
                    {
                        reqStream.Write(bs, 0, bs.Length);
                    }

                    using (WebResponse wr = req.GetResponse())
                    {
                        //在這裡對接收到的頁面內容進行處理

                        using (Stream myStream = wr.GetResponseStream())
                        {
                            using (StreamReader myStreamReader = new StreamReader(myStream, System.Text.Encoding.UTF8))
                            {
                                strHTML = myStreamReader.ReadToEnd();
                            }
                        }
                    }
                }
                else
                {
                    Uri myUri = new Uri(url);

                    // Create a 'HttpWebRequest' object for the specified url. 

                    HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(myUri);

                    // Set the user agent as if we were a web browser

                    myHttpWebRequest.UserAgent = @"Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.0.4) Gecko/20060508 Firefox/1.5.0.4";

                    HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();

                    var stream = myHttpWebResponse.GetResponseStream();

                    var reader = new StreamReader(stream);

                    var html = reader.ReadToEnd();

                    // Release resources of response object.

                    myHttpWebResponse.Close();

                    return html;
                }
            }
            else
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                //request.Method = "HEAD";

                //request.AllowAutoRedirect = false;

                request.Credentials = CredentialCache.DefaultCredentials;

                // Ignore Certificate validation failures (aka untrusted certificate + certificate chains)

                ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                Stream resStream = response.GetResponseStream();

                StreamReader reader = new StreamReader(resStream, System.Text.Encoding.UTF8);

                strHTML = reader.ReadToEnd();
            }
            return strHTML;
        }
        catch (Exception ex) { throw ex; }
    }
}
#endregion
