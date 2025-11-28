using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;

/// <summary>
/// DocManage_DB 的摘要描述
/// </summary>
public class DocManage_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord { set { KeyWord = value; } }
    #region Private
    string id = string.Empty;
    string guid = string.Empty;
    string 父層guid = string.Empty;
    string 類別名稱 = string.Empty;
    string 階層 = string.Empty;
    string 排序 = string.Empty;
    DateTime 建立日期;
    string 建立者 = string.Empty;
    DateTime 修改日期;
    string 修改者 = string.Empty;
    string 資料狀態 = string.Empty;

    #endregion
    #region Public
    public string _id { set { id = value; } }
    public string _guid { set { guid = value; } }
    public string _父層guid { set { 父層guid = value; } }
    public string _類別名稱 { set { 類別名稱 = value; } }
    public string _階層 { set { 階層 = value; } }
    public string _排序 { set { 排序 = value; } }
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
WITH CategoryCTE AS (
    -- 第 1 層
    SELECT 
        guid,
        [父層guid],
        [類別名稱],
        CAST([類別名稱] AS NVARCHAR(200)) AS lv1,
        CAST(NULL AS NVARCHAR(200))        AS lv2,
        CAST(NULL AS NVARCHAR(200))        AS lv3,
        1 AS level,
        [排序] AS sort
    FROM 公文類別階層表
    WHERE [父層guid] IS NULL

    UNION ALL

    -- 第 2 層以後
    SELECT 
        c.guid,
        c.[父層guid],
        c.[類別名稱],
        p.lv1,
        CASE WHEN p.level = 1
             THEN CAST(c.[類別名稱] AS NVARCHAR(200))
             ELSE p.lv2
        END AS lv2,
        CASE WHEN p.level = 2
             THEN CAST(c.[類別名稱] AS NVARCHAR(200))
             ELSE p.lv3
        END AS lv3,
        p.level + 1 AS level,
        c.[排序] AS sort
    FROM 公文類別階層表 c
    INNER JOIN CategoryCTE p ON c.[父層guid] = p.guid
),

-- 針對附件檔：同一個 父層guid + 排序 只保留「版本最大」那一筆
LatestAttach AS (
    SELECT
        a.*,
        ROW_NUMBER() OVER(
            PARTITION BY a.[父層guid], a.[排序]
            ORDER BY CAST(a.版本 AS INT) DESC
        ) AS rn
    FROM 附件檔 a
)

SELECT 
    c.guid                     AS 類別guid,
    c.[類別名稱],
    c.lv1,
    c.lv2,
    c.lv3,
    a.guid                     AS 附件guid,
    CONVERT(nvarchar(100),a.修改日期, 20) as 上傳日期,
    a.檔案類型,
    a.原檔名,
    a.新檔名 + ISNULL(a.附檔名, '') AS 完整檔名,
    a.排序                     AS 附件排序,
	a.版本
FROM CategoryCTE c
INNER JOIN LatestAttach a
    ON a.[父層guid] = c.guid
   AND a.rn = 1                 -- 只取每組中版本最高的那一筆
ORDER BY
    c.lv1,
    c.lv2,
    c.lv3,
    a.排序;

");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        //oCmd.Parameters.AddWithValue("@userid", userid);
        //oCmd.Parameters.AddWithValue("@useridx", useridx);
        //oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        //oCmd.Parameters.AddWithValue("@StartFineDate", StartFineDate);
        //oCmd.Parameters.AddWithValue("@EndFineDate", EndFineDate);
        //oCmd.Parameters.AddWithValue("@pStart", pStart);
        //oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        //oCmd.Parameters.AddWithValue("@guid", guid);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oda.Fill(ds);
        return ds;
    }

    public DataTable GetDemoData()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
WITH CategoryCTE AS (
    -- 第 1 層
    SELECT 
        guid,
        [父層guid],
        [類別名稱],
        CAST([類別名稱] AS NVARCHAR(200)) AS lv1,
        CAST(NULL AS NVARCHAR(200))        AS lv2,
        CAST(NULL AS NVARCHAR(200))        AS lv3,
        1 AS level,
        [排序] AS sort
    FROM 公文類別階層表
    WHERE [父層guid] IS NULL

    UNION ALL

    -- 第 2 層以後
    SELECT 
        c.guid,
        c.[父層guid],
        c.[類別名稱],
        p.lv1,
        CASE WHEN p.level = 1
             THEN CAST(c.[類別名稱] AS NVARCHAR(200))
             ELSE p.lv2
        END AS lv2,
        CASE WHEN p.level = 2
             THEN CAST(c.[類別名稱] AS NVARCHAR(200))
             ELSE p.lv3
        END AS lv3,
        p.level + 1 AS level,
        c.[排序] AS sort
    FROM 公文類別階層表 c
    INNER JOIN CategoryCTE p ON c.[父層guid] = p.guid
),

-- 針對附件檔：同一個 父層guid + 排序 只保留「版本最大」那一筆
LatestAttach AS (
    SELECT
        a.*,
        ROW_NUMBER() OVER(
            PARTITION BY a.[父層guid], a.[排序]
            ORDER BY CAST(a.版本 AS INT) DESC
        ) AS rn
    FROM 附件檔 a
    WHERE 檔案類型='01'
)

SELECT 
    c.guid                     AS 類別guid,
    c.[類別名稱],
    c.lv1,
    c.lv2,
    c.lv3,
    a.guid                     AS 附件guid,
    CONVERT(nvarchar(100),a.修改日期, 20) as 上傳日期,
    a.檔案類型,
    a.原檔名,
    a.新檔名 + ISNULL(a.附檔名, '') AS 完整檔名,
    a.排序                     AS 附件排序,
	a.版本,
    ISNULL(c.lv1 + '\', '') +
    ISNULL(c.lv2 + '\', '') +
    ISNULL(c.lv3 + '\', '') +
    a.新檔名 + ISNULL(a.附檔名, '') AS RelativePath

FROM CategoryCTE c
INNER JOIN LatestAttach a
    ON a.[父層guid] = c.guid
   AND a.rn = 1                 -- 只取每組中版本最高的那一筆
WHERE a.guid = @guid and a.檔案類型='01' 
ORDER BY
    c.lv1,
    c.lv2,
    c.lv3,
    a.排序;
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        //oCmd.Parameters.AddWithValue("@userid", userid);
        //oCmd.Parameters.AddWithValue("@useridx", useridx);
        //oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        //oCmd.Parameters.AddWithValue("@StartFineDate", StartFineDate);
        //oCmd.Parameters.AddWithValue("@EndFineDate", EndFineDate);
        //oCmd.Parameters.AddWithValue("@pStart", pStart);
        //oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oCmd.Parameters.AddWithValue("@guid", guid);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oda.Fill(ds);
        return ds;
    }

    public DataTable GetDemoDataNoFile(string fileCode)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
WITH CategoryCTE AS (
    SELECT 
        guid,
        [父層guid],
        [類別名稱],
        CAST([類別名稱] AS NVARCHAR(200)) AS lv1,
        CAST(NULL AS NVARCHAR(200))        AS lv2,
        CAST(NULL AS NVARCHAR(200))        AS lv3,
        1 AS level,
        [排序] AS sort
    FROM 公文類別階層表
    WHERE [父層guid] IS NULL

    UNION ALL

    SELECT 
        c.guid,
        c.[父層guid],
        c.[類別名稱],
        p.lv1,
        CASE WHEN p.level = 1
             THEN CAST(c.[類別名稱] AS NVARCHAR(200))
             ELSE p.lv2
        END AS lv2,
        CASE WHEN p.level = 2
             THEN CAST(c.[類別名稱] AS NVARCHAR(200))
             ELSE p.lv3
        END AS lv3,
        p.level + 1 AS level,
        c.[排序] AS sort
    FROM 公文類別階層表 c
    INNER JOIN CategoryCTE p ON c.[父層guid] = p.guid
)

SELECT 
    c.guid                     AS 類別guid,
    c.[類別名稱],
    c.lv1,
    c.lv2,
    c.lv3,
    a.guid                     AS 附件guid,
    a.表單名稱,
    a.表單編號,
    a.原檔名,
    a.新檔名,
    a.附檔名,
    a.排序,
    a.新檔名 + ISNULL(a.附檔名, '') AS 完整檔名,

    -- ★ 這裡組出「相對路徑」
    ISNULL(c.lv1 + '\', '') +
    ISNULL(c.lv2 + '\', '') +
    ISNULL(c.lv3 + '\', '')  AS RelativePath

FROM CategoryCTE c
INNER JOIN 附件檔 a
    ON a.[父層guid] = c.guid
WHERE a.guid = @guid and a.檔案類型=@檔案類型;
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        oCmd.Parameters.AddWithValue("@guid", guid);
        oCmd.Parameters.AddWithValue("@檔案類型", fileCode);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oda.Fill(ds);
        return ds;
    }

    public DataSet GetIncludeFieldDemoList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
WITH CategoryCTE AS (
    -- 第 1 層
    SELECT 
        guid,
        [父層guid],
        [類別名稱],
        CAST([類別名稱] AS NVARCHAR(200)) AS lv1,
        CAST(NULL AS NVARCHAR(200))        AS lv2,
        CAST(NULL AS NVARCHAR(200))        AS lv3,
        1 AS level,
        [排序] AS sort
    FROM 公文類別階層表
    WHERE [父層guid] IS NULL

    UNION ALL

    -- 第 2 層以後
    SELECT 
        c.guid,
        c.[父層guid],
        c.[類別名稱],
        p.lv1,
        CASE WHEN p.level = 1
             THEN CAST(c.[類別名稱] AS NVARCHAR(200))
             ELSE p.lv2
        END AS lv2,
        CASE WHEN p.level = 2
             THEN CAST(c.[類別名稱] AS NVARCHAR(200))
             ELSE p.lv3
        END AS lv3,
        p.level + 1 AS level,
        c.[排序] AS sort
    FROM 公文類別階層表 c
    INNER JOIN CategoryCTE p ON c.[父層guid] = p.guid
),

-- 每一類別下只抓「版本最大」的附件
LatestAttach AS (
    SELECT
        a.*,
        ROW_NUMBER() OVER(
            PARTITION BY a.[父層guid], a.[排序]
            ORDER BY CAST(a.版本 AS INT) DESC
        ) AS rn
    FROM 附件檔 a
),

-- 欄位數（依附件 + 版本）
FieldCount AS (
    SELECT 
        guid AS 附件guid,
        版本,
        COUNT(*) AS 欄位數
    FROM 公文欄位定義表
    WHERE 是否已刪除 = 0
    GROUP BY guid, 版本
),

-- 最新修改者與修改時間（依附件 + 版本）
FieldLastUpdate AS (
    SELECT *
    FROM (
        SELECT
            guid AS 附件guid,
            版本,
            修改者,
            修改時間,
            ROW_NUMBER() OVER(
                PARTITION BY guid, 版本
                ORDER BY 修改時間 DESC
            ) AS rn
        FROM 公文欄位定義表
        WHERE 是否已刪除 = 0
    ) x
    WHERE rn = 1
)

SELECT 
    c.guid                                  AS 類別guid,
    c.[類別名稱],
    c.lv1,
    c.lv2,
    c.lv3,

    a.guid                                  AS 附件guid,
    a.原檔名                                AS 範本名稱,
    a.版本                                  AS 目前版本,
    a.排序									AS 附件排序,
    c.lv1 
        + ISNULL(' / ' + c.lv2, '') 
        + ISNULL(' / ' + c.lv3, '')         AS 所屬類別,

    ISNULL(fc.欄位數, 0)                     AS 欄位數,
    flu.修改者                              AS 最後修改者,
    CONVERT(nvarchar(100),flu.修改時間, 20) AS 最後修改時間,

    a.新檔名 + ISNULL(a.附檔名, '')        AS 完整檔名

FROM CategoryCTE c
JOIN LatestAttach a
      ON a.[父層guid] = c.guid
     AND a.rn = 1

LEFT JOIN FieldCount fc
       ON fc.附件guid = a.guid
      AND fc.版本 = a.版本

LEFT JOIN FieldLastUpdate flu
       ON flu.附件guid = a.guid
      AND flu.版本 = a.版本

ORDER BY
    c.lv1,
    c.lv2,
    c.lv3,
    a.排序;

");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        //oCmd.Parameters.AddWithValue("@userid", userid);
        //oCmd.Parameters.AddWithValue("@useridx", useridx);
        //oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        //oCmd.Parameters.AddWithValue("@StartFineDate", StartFineDate);
        //oCmd.Parameters.AddWithValue("@EndFineDate", EndFineDate);
        //oCmd.Parameters.AddWithValue("@pStart", pStart);
        //oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        //oCmd.Parameters.AddWithValue("@guid", guid);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oda.Fill(ds);
        return ds;
    }

    public DataSet GetDemoVersionList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
WITH Ver AS (
    SELECT 
        a2.guid,
        a2.版本,
        ROW_NUMBER() OVER(
            PARTITION BY a2.版本
            ORDER BY a2.修改日期 DESC
        ) rn
    FROM 附件檔 a1
    JOIN 附件檔 a2
       ON a1.父層guid = a2.父層guid
      AND a1.排序     = a2.排序
    WHERE a1.guid = @guid 
)
SELECT guid, 版本
FROM Ver
WHERE rn = 1
ORDER BY CONVERT(int, 版本);
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        oCmd.Parameters.AddWithValue("@guid", guid);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oda.Fill(ds);
        return ds;
    }

    public DataTable GetMenuList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"].ToString());
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select * from 公文類別階層表 
where (@父層guid='' or 父層guid=@父層guid) and (@階層='' or 階層=@階層)
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        oCmd.Parameters.AddWithValue("@父層guid", 父層guid);
        oCmd.Parameters.AddWithValue("@階層", 階層);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oda.Fill(ds);
        return ds;
    }
}