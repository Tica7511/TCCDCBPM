using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;

/// <summary>
/// DocContent_DB 的摘要描述
/// </summary>
public class DocContent_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord { set { KeyWord = value; } }
    #region private
    string id = string.Empty;
    string guid = string.Empty;
    string 範本guid = string.Empty;
    string 版本 = string.Empty;
    string 項目代碼 = string.Empty;
    string 項目類型 = string.Empty;
    string 項目內容 = string.Empty;
    string 建立者 = string.Empty;
    DateTime 建立日期;
    string 修改者 = string.Empty;
    DateTime 修改日期;
    string 資料狀態 = string.Empty;
    #endregion
    #region public
    public string _id { set { id = value; } }
    public string _guid { set { guid = value; } }
    public string _範本guid { set { 範本guid = value; } }
    public string _版本 { set { 版本 = value; } }
    public string _項目代碼 { set { 項目代碼 = value; } }
    public string _項目類型 { set { 項目類型 = value; } }
    public string _項目內容 { set { 項目內容 = value; } }
    public string _建立者 { set { 建立者 = value; } }
    public DateTime _建立日期 { set { 建立日期 = value; } }
    public string _修改者 { set { 修改者 = value; } }
    public DateTime _修改日期 { set { 修改日期 = value; } }
    public string _資料狀態 { set { 資料狀態 = value; } }
    #endregion

    public DataTable GetData()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select * ,
項目名稱=(select 項目名稱 from 公文欄位定義表 
where 公文欄位定義表.項目代碼=公文內容.項目代碼 and 公文欄位定義表.guid=公文內容.範本guid)
from 公文內容 
where (@guid='' or guid=@guid) and (@版本='' or 版本=@版本)
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@版本", 版本);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oda.Fill(ds);
        return ds;
    }

    public DataTable GetContent()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select * from 公文內容 
where (@guid='' or guid=@guid) and (@版本='' or 版本=@版本)
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@版本", 版本);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oda.Fill(ds);
        return ds;
    }

    public DataTable GetData_Trans(SqlConnection oConn, SqlTransaction oTran)
    {
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select *
from 公文內容
where (@guid='' or guid=@guid)
  and (@版本='' or 版本=@版本)
  and (@項目代碼='' or 項目代碼=@項目代碼)
");

        SqlCommand oCmd = oConn.CreateCommand();
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        // 使用交易
        oCmd.Transaction = oTran;

        // 傳入參數
        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@版本", 版本);
        oCmd.Parameters.AddWithValue("@項目代碼", 項目代碼);

        // 用 DataAdapter 回填 DataTable（這是關鍵）
        SqlDataAdapter da = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        da.Fill(dt);

        return dt;
    }
}