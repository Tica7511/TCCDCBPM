using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;

/// <summary>
/// Doc_DB 的摘要描述
/// </summary>
public class Doc_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord { set { KeyWord = value; } }
    #region private
    string id = string.Empty;
    string guid = string.Empty;
    string 公文guid = string.Empty;
    string 主旨 = string.Empty;
    string 狀態 = string.Empty;
    string 建立者 = string.Empty;
    DateTime 建立日期;
    string 修改者 = string.Empty;
    DateTime 修改日期;
    string 資料狀態 = string.Empty;
    #endregion
    #region public
    public string _id { set { id = value; } }
    public string _guid { set { guid = value; } }
    public string _公文guid { set { 公文guid = value; } }
    public string _主旨 { set { 主旨 = value; } }
    public string _狀態 { set { 狀態 = value; } }
    public string _建立者 { set { 建立者 = value; } }
    public DateTime _建立日期 { set { 建立日期 = value; } }
    public string _修改者 { set { 修改者 = value; } }
    public DateTime _修改日期 { set { 修改日期 = value; } }
    public string _資料狀態 { set { 資料狀態 = value; } }
    #endregion

    public DataTable GetList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
SELECT 
    L.*,
    A.父層guid,
	A.排序,
    A.表單名稱,
    A.表單編號,
	A.版本,
    CONVERT(nvarchar(100), L.建立時間, 20) AS 申請日期,
    CONVERT(nvarchar(100), L.修改時間, 20) AS 上傳日期,
    表單狀態_V = (
        SELECT 項目名稱 
        FROM 代碼檔 
        WHERE 群組代碼 = '003' 
          AND 項目代碼 = L.狀態
    )
FROM 公文列表 L
OUTER APPLY (
    SELECT TOP 1 *
    FROM 附件檔 A
    WHERE A.guid = L.公文guid
    ORDER BY A.版本 DESC       
) A
ORDER BY L.修改時間 DESC;
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oda.Fill(ds);
        return ds;
    }

    public DataTable InsertData()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
insert into 公文列表 (
guid,
公文guid,
主旨,
狀態,
建立者,
修改者,
資料狀態
) values (
@guid,
@公文guid,
@主旨,
@狀態,
@建立者,
@修改者,
@資料狀態 
) ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@公文guid", 公文guid);
        oCmd.Parameters.AddWithValue("@主旨", 主旨);
        oCmd.Parameters.AddWithValue("@狀態", 狀態);
        oCmd.Parameters.AddWithValue("@建立者", 建立者);
        oCmd.Parameters.AddWithValue("@修改者", 修改者);
        oCmd.Parameters.AddWithValue("@資料狀態", "A");

        oda.Fill(ds);
        return ds;
    }

    public void Insert_Trans(SqlConnection oConn, SqlTransaction oTran)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(@"
insert into 公文列表 (
guid,
公文guid,
主旨,
狀態,
建立者,
修改者,
資料狀態
) values (
@guid,
@公文guid,
@主旨,
@狀態,
@建立者,
@修改者,
@資料狀態 
) 
 ");
        SqlCommand oCmd = oConn.CreateCommand();
        oCmd.CommandText = sb.ToString();

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@公文guid", 公文guid);
        oCmd.Parameters.AddWithValue("@主旨", 主旨);
        oCmd.Parameters.AddWithValue("@狀態", 狀態);
        oCmd.Parameters.AddWithValue("@建立者", 建立者);
        oCmd.Parameters.AddWithValue("@修改者", 修改者);
        oCmd.Parameters.AddWithValue("@資料狀態", "A");

        oCmd.Transaction = oTran;
        oCmd.ExecuteNonQuery();
    }

    public DataTable DeleteData()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
update 公文列表 set
狀態='2',
資料狀態=@資料狀態,
修改者=@修改者,
修改時間=GETDATE()
where guid=@guid 
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@修改者", 修改者);
        oCmd.Parameters.AddWithValue("@資料狀態", "D");

        oda.Fill(ds);
        return ds;
    }
}