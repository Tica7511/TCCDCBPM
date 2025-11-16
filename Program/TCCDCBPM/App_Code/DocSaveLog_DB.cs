using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;

/// <summary>
/// DocSaveLog_DB 的摘要描述
/// </summary>
public class DocSaveLog_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord { set { KeyWord = value; } }
    #region private
    string id = string.Empty;
    string guid = string.Empty;
    string 狀態 = string.Empty;
    string 類別 = string.Empty;
    string 建立者 = string.Empty;
    DateTime 建立日期;
    string 修改者 = string.Empty;
    DateTime 修改日期;
    string 資料狀態 = string.Empty;
    #endregion
    #region public
    public string _id { set { id = value; } }
    public string _guid { set { guid = value; } }
    public string _狀態 { set { 狀態 = value; } }
    public string _類別 { set { 類別 = value; } }
    public string _建立者 { set { 建立者 = value; } }
    public DateTime _建立日期 { set { 建立日期 = value; } }
    public string _修改者 { set { 修改者 = value; } }
    public DateTime _修改日期 { set { 修改日期 = value; } }
    public string _資料狀態 { set { 資料狀態 = value; } }
    #endregion

    public void Insert_Trans(SqlConnection oConn, SqlTransaction oTran)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(@"
insert into 公文儲存log (
guid,
狀態,
類別,
建立者,
修改者,
資料狀態
) values (
@guid,
@狀態,
@類別,
@建立者,
@修改者,
@資料狀態 
) 
 ");
        SqlCommand oCmd = oConn.CreateCommand();
        oCmd.CommandText = sb.ToString();

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@狀態", 狀態);
        oCmd.Parameters.AddWithValue("@類別", 類別);
        oCmd.Parameters.AddWithValue("@建立者", 建立者);
        oCmd.Parameters.AddWithValue("@修改者", 修改者);
        oCmd.Parameters.AddWithValue("@資料狀態", "A");

        oCmd.Transaction = oTran;
        oCmd.ExecuteNonQuery();
    }
}