using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;

/// <summary>
/// FileTable_DB 的摘要描述
/// </summary>
public class FileTable_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord { set { KeyWord = value; } }
    #region private
    string id = string.Empty;
    string guid = string.Empty;
    string 父層guid = string.Empty;
    string 年度 = string.Empty;
    string 檔案類型 = string.Empty;
    string 原檔名 = string.Empty;
    string 新檔名 = string.Empty;
    string 附檔名 = string.Empty;
    string 排序 = string.Empty;
    string 版本 = string.Empty;
    string 檔案大小 = string.Empty;
    string 建立者 = string.Empty;
    DateTime 建立日期;
    string 修改者 = string.Empty;
    DateTime 修改日期;
    string 資料狀態 = string.Empty;
    #endregion
    #region public
    public string _id { set { id = value; } }
    public string _guid { set { guid = value; } }
    public string _父層guid { set { 父層guid = value; } }
    public string _年度 { set { 年度 = value; } }
    public string _檔案類型 { set { 檔案類型 = value; } }
    public string _原檔名 { set { 原檔名 = value; } }
    public string _新檔名 { set { 新檔名 = value; } }
    public string _附檔名 { set { 附檔名 = value; } }
    public string _排序 { set { 排序 = value; } }
    public string _版本 { set { 版本 = value; } }
    public string _檔案大小 { set { 檔案大小 = value; } }
    public string _建立者 { set { 建立者 = value; } }
    public DateTime _建立日期 { set { 建立日期 = value; } }
    public string _修改者 { set { 修改者 = value; } }
    public DateTime _修改日期 { set { 修改日期 = value; } }
    public string _資料狀態 { set { 資料狀態 = value; } }
    #endregion

    public DataTable GetData()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select *, CONVERT(nvarchar(100),建立日期, 20) as 上傳日期 from 附件檔 where 資料狀態='A' and (@父層guid='' or 父層guid=@父層guid) and (@年度='' or 年度=@年度) ");
        if (!string.IsNullOrEmpty(檔案類型))
            sb.Append(@"and 檔案類型=@檔案類型 ");
        if (!string.IsNullOrEmpty(排序))
            sb.Append(@"and 排序=@排序 ");
        if (!string.IsNullOrEmpty(版本))
            sb.Append(@"and 版本=@版本 ");
        if (!string.IsNullOrEmpty(guid))
            sb.Append(@"and guid=@guid ");
        else
            sb.Append(@"and guid='' ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@父層guid", 父層guid);
        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@年度", 年度);
        oCmd.Parameters.AddWithValue("@排序", 排序);
        oCmd.Parameters.AddWithValue("@版本", 版本);
        oCmd.Parameters.AddWithValue("@檔案類型", 檔案類型);

        oda.Fill(ds);
        return ds;
    }

    public DataTable GetSnMaxData()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select TOP 1 * 
from 附件檔 
where 資料狀態='A' and (@guid='' or guid=@guid) and (@檔案類型='' or 檔案類型=@檔案類型) 
and (@父層guid='' or 父層guid=@父層guid) and  (@排序='' or 排序=@排序) 
ORDER BY CONVERT(int,版本) DESC ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@檔案類型", 檔案類型);
        oCmd.Parameters.AddWithValue("@父層guid", 父層guid);
        oCmd.Parameters.AddWithValue("@排序", 排序);

        oda.Fill(ds);
        return ds;
    }

    public void UpdateFile_Trans(SqlConnection oConn, SqlTransaction oTran)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(@"
insert into 附件檔 (
guid,
父層guid,
年度,
檔案類型,
原檔名,
新檔名, 
附檔名, 
排序, 
版本, 
檔案大小,
建立者,
修改者,
資料狀態
) values (
@guid,
@父層guid,
@年度,
@檔案類型,
@原檔名,
@新檔名, 
@附檔名, 
@排序, 
@版本, 
@檔案大小,
@建立者,
@修改者,
@資料狀態 
) 
 ");
        SqlCommand oCmd = oConn.CreateCommand();
        oCmd.CommandText = sb.ToString();

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@年度", 年度);
        oCmd.Parameters.AddWithValue("@父層guid", 父層guid);
        oCmd.Parameters.AddWithValue("@檔案類型", 檔案類型);
        oCmd.Parameters.AddWithValue("@原檔名", 原檔名);
        oCmd.Parameters.AddWithValue("@新檔名", 新檔名);
        oCmd.Parameters.AddWithValue("@附檔名", 附檔名);
        oCmd.Parameters.AddWithValue("@排序", 排序);
        oCmd.Parameters.AddWithValue("@版本", 版本);
        oCmd.Parameters.AddWithValue("@檔案大小", 檔案大小);
        oCmd.Parameters.AddWithValue("@建立者", 建立者);
        oCmd.Parameters.AddWithValue("@修改者", 修改者);
        oCmd.Parameters.AddWithValue("@資料狀態", "A");

        oCmd.Transaction = oTran;
        oCmd.ExecuteNonQuery();
    }

    public DataTable InsertFile()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
insert into 附件檔 (
guid,
父層guid,
年度,
檔案類型,
原檔名,
新檔名, 
附檔名, 
排序, 
版本, 
檔案大小,
建立者,
修改者,
資料狀態
) values (
@guid,
@父層guid,
@年度,
@檔案類型,
@原檔名,
@新檔名, 
@附檔名, 
@排序, 
@版本, 
@檔案大小,
@建立者,
@修改者,
@資料狀態 
) ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@年度", 年度);
        oCmd.Parameters.AddWithValue("@父層guid", 父層guid);
        oCmd.Parameters.AddWithValue("@檔案類型", 檔案類型);
        oCmd.Parameters.AddWithValue("@原檔名", 原檔名);
        oCmd.Parameters.AddWithValue("@新檔名", 新檔名);
        oCmd.Parameters.AddWithValue("@附檔名", 附檔名);
        oCmd.Parameters.AddWithValue("@排序", 排序);
        oCmd.Parameters.AddWithValue("@版本", 版本);
        oCmd.Parameters.AddWithValue("@檔案大小", 檔案大小);
        oCmd.Parameters.AddWithValue("@建立者", 建立者);
        oCmd.Parameters.AddWithValue("@修改者", 修改者);
        oCmd.Parameters.AddWithValue("@資料狀態", "A");

        oda.Fill(ds);
        return ds;
    }
}