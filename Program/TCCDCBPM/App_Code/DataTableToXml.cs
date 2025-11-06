using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Xml;

/// <summary>
/// Summary description for DataTableToXml
/// </summary>
public class DataTableToXml
{
    static string ProcessDataTableToXml(DataTable dt, string NodeName, string RowsName)
    {
        string xmlstr = string.Empty;
        string tbsName = "DD" + Guid.NewGuid().ToString("N").Substring(0, 5);
        dt.TableName = tbsName;
        foreach (DataRow Rows in dt.Rows)
        {
            int CountRows = Rows.ItemArray.Length;
            int i = 0;
            for (i = 0; i < CountRows; i++)
            {

                if (string.IsNullOrEmpty(Rows[i].ToString()))
                {
                    try
                    {
                        Rows[i] = " ";
                    }
                    catch
                    {
                        if (dt.Columns[i].DataType == typeof(DateTime))
                            Rows[i] = DateTime.Parse("1900/01/01");
                        if (dt.Columns[i].DataType == typeof(int) ||
                            dt.Columns[i].DataType == typeof(Int16) ||
                            dt.Columns[i].DataType == typeof(Int32) ||
                             dt.Columns[i].DataType == typeof(Int64))
                            Rows[i] = 0;
                    }
                }

            }

        }

        System.IO.StringWriter writer = new System.IO.StringWriter();

        dt.WriteXml(writer, XmlWriteMode.WriteSchema, false);

        XmlDocument XmlDoc = new XmlDocument();
        XmlDoc.LoadXml(writer.ToString());
        XmlNodeList xnList = (XmlNodeList)XmlDoc.SelectNodes("//" + tbsName);

        foreach (XmlNode xn in xnList)
        {
            xmlstr += xn.OuterXml;
        }
        if (RowsName != null)
        {
            string x = @"<{0}>{1}</{0}>";
            x = string.Format(x, NodeName, xmlstr.Replace(tbsName, RowsName));
            return x;
        }
        else
        {
            string x = @"{0}";
            x = string.Format(x, xmlstr.Replace(tbsName, NodeName));
            return x;
        }
    }

    static string ProcessDataTableToXmlByAttribute(DataTable dt, string NodeName, string RowsName)
    {
        string xmlstr = string.Empty;
        XmlDocument doc = new XmlDocument();
        /// 根節點
        XmlElement XmlRoot = doc.CreateElement(NodeName);
        doc.AppendChild(XmlRoot);
        XmlElement xNode = doc.DocumentElement;
        if (dt.Rows.Count > 0)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                xNode = doc.CreateElement(RowsName);
                for (int j = 0; j < dt.Rows[i].ItemArray.Length; j++)
                {
                    //Format DateTime
                    if (dt.Columns[j].DataType == typeof(DateTime))
                    {
                        if (string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString()))
                            xNode.SetAttribute(dt.Columns[j].ColumnName, dt.Rows[i].ItemArray[j].ToString());
                        else
                            xNode.SetAttribute(dt.Columns[j].ColumnName, DateTime.Parse(dt.Rows[i].ItemArray[j].ToString()).ToString("yyyy/MM/dd"));
                    }
                    else
                        xNode.SetAttribute(dt.Columns[j].ColumnName, dt.Rows[i].ItemArray[j].ToString());
                }
                XmlRoot.AppendChild(xNode);
            }
        }

        xmlstr = XmlRoot.OuterXml;
        return xmlstr;
    }


    public static string ConvertDatatableToXML(DataTable dt, string NodeName, string RowsName)
    {
        return ProcessDataTableToXml(dt, NodeName, RowsName);
    }
    public static string ConvertDatatableToXML(DataTable dt, string NodeName)
    {
        return ProcessDataTableToXml(dt, NodeName, null);
    }
    public static string ConvertDatatableToXmlByAttribute(DataTable dt, string NodeName, string RowsName)
    {
        return ProcessDataTableToXmlByAttribute(dt, NodeName, RowsName);
    }
}