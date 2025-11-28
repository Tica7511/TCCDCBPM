using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;

/// <summary>
/// DocFieldManage_DB 的摘要描述
/// </summary>
public class DocFieldManage_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord { set { KeyWord = value; } }
    #region Private
    string id = string.Empty;
    string guid = string.Empty;
    string 版本 = string.Empty;
    string 項目代碼 = string.Empty;
    string 項目類型 = string.Empty;
    string 項目名稱 = string.Empty;
    string 是否已刪除 = string.Empty;
    DateTime 建立日期;
    string 建立者 = string.Empty;
    DateTime 修改日期;
    string 修改者 = string.Empty;
    string 資料狀態 = string.Empty;

    #endregion
    #region Public
    public string _id { set { id = value; } }
    public string _guid { set { guid = value; } }
    public string _版本 { set { 版本 = value; } }
    public string _項目代碼 { set { 項目代碼 = value; } }
    public string _項目類型 { set { 項目類型 = value; } }
    public string _項目名稱 { set { 項目名稱 = value; } }
    public string _是否已刪除 { set { 是否已刪除 = value; } }
    public DateTime _建立日期 { set { 建立日期 = value; } }
    public string _建立者 { set { 建立者 = value; } }
    public DateTime _修改日期 { set { 修改日期 = value; } }
    public string _修改者 { set { 修改者 = value; } }
    public string _資料狀態 { set { 資料狀態 = value; } }

    #endregion

    public DataSet GetList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
SELECT * ,
項目類型_V=(select 項目名稱 from 代碼檔 where 代碼檔.群組代碼='002' and 
代碼檔.項目代碼=公文欄位定義表.項目類型),
CASE WHEN 是否已刪除 = 1 THEN '已刪除'
ELSE '' END AS 是否已刪除_V 
FROM 公文欄位定義表 
WHERE guid=@guid 
ORDER BY 項目代碼 
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        oCmd.Parameters.AddWithValue("@guid", guid);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oda.Fill(ds);
        return ds;
    }

    public void UpdateFieldName()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        StringBuilder sb = new StringBuilder();
        sb.Append(@"update 公文欄位定義表 set
項目名稱=@項目名稱,
修改者=@修改者,
修改時間=getdate() 
where guid=@guid and 項目代碼=@項目代碼 ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@項目名稱", 項目名稱);
        oCmd.Parameters.AddWithValue("@項目代碼", 項目代碼);
        oCmd.Parameters.AddWithValue("@修改者", 修改者);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}