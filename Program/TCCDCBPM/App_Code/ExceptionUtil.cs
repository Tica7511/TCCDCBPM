using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml;

/// <summary>
/// ExceptionUtil 的摘要描述
/// </summary>
public class ExceptionUtil
{
    public static XmlDocument GetExceptionDocument(Exception err)
    {
        ///---------------------------------------------------------------------------
        /// 作    者: Nick
        /// 功    能: 把應用程式錯誤訊息封裝為 XmlDocument
        /// 說    明: 
        ///---------------------------------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        xDoc.LoadXml(@"<?xml version='1.0' encoding='utf-8'?><root />");
        XmlElement xNode = xDoc.CreateElement("Error");
        xNode.SetAttribute("Message", (string.IsNullOrWhiteSpace(err.Message) ? "" : err.Message.Trim().Replace("'", "")));
        xDoc.DocumentElement.AppendChild(xNode);
        return xDoc;
    }

    public static XmlDocument GetErrorMassageDocument(string eMsg)
    {
        ///---------------------------------------------------------------------------
        /// 作    者: Nick
        /// 功    能: 把錯誤訊息封裝為 XmlDocument
        /// 說    明: 
        ///---------------------------------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        xDoc.LoadXml(@"<?xml version='1.0' encoding='utf-8'?><root />");
        XmlElement xNode = xDoc.CreateElement("Error");
        xNode.SetAttribute("Message", (string.IsNullOrWhiteSpace(eMsg) ? "" : eMsg.Trim()));
        xDoc.DocumentElement.AppendChild(xNode);
        return xDoc;
    }


    public static object GetErrorMassageJson(string eMsg)
    {
        ///---------------------------------------------------------------------------
        /// 作    者: Nick
        /// 功    能: 把錯誤訊息封裝為 Json
        /// 說    明: 
        ///---------------------------------------------------------------------------
        var result = new
        {
            ErrorMassage = eMsg.Trim()
        };

        return result;
    }
}