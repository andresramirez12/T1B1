using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;
using System.Collections;
using System.Globalization;
using System.Windows.Forms;

namespace T1.B1.Helpers
{
    public class WithHoldingTaxes
    {
        static private WithHoldingTaxes objWTHelpers = null;
        static private string lastWTperItemCache = "";
        static private string lastSystemWTCache = "";

        private WithHoldingTaxes()
        {
            
        }

        static public void cacheAllItemsWTInfo(string strCacheName)
        {

            SAPbobsCOM.Recordset objRecordset = null;
            string strSql = "";
            XmlDocument objDocument = null;

            try
            {
                objRecordset = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                strSql = "select [@BYB_T1WHT100].Code ,U_Item, U_Type, Rate, PrctBsAmnt from[@BYB_T1WHT100] inner join[@BYB_T1WHT101] on [@BYB_T1WHT101].Code = [@BYB_T1WHT100].Code inner join OWHT on [@BYB_T1WHT101].U_Item = OWHT.WTCode where [@BYB_T1WHT101].U_Item is not null";
                objRecordset.DoQuery(strSql);
                objDocument = new XmlDocument();
                objDocument.LoadXml(objRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));
                //"WTCodesByItem"
                BYBCache.Instance.addToCache(strCacheName, objDocument, BYBCache.objCachePriority.Default);
                lastWTperItemCache = strCacheName;

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }



        }

        static public void cacheAllWTDefinitions(string strCacheName)
        {

            //TODO get date from form and not form system in case you can do WT for future dates

            SAPbobsCOM.Recordset objRecordset = null;
            XmlDocument objDocument = null;
            string strSql = "";

            try
            {
                objRecordset = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                strSql = "select OWHT.WTCode ,WTName ,WHT1.Rate ,PrctBsAmnt  ,Account from OWHT inner join WHT1 on OWHT.WTCode = WHT1.WTCode where getDate() >= WHT1.EffecDate";
                objRecordset.DoQuery(strSql);
                objDocument = new XmlDocument();
                objDocument.LoadXml(objRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));
                //"WTCodesDefinitions"
                BYBCache.Instance.addToCache(strCacheName, objDocument, BYBCache.objCachePriority.Default);
                lastSystemWTCache = strCacheName;


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }



        }

        static public double getPercentFromCode(string WTCode)
        {
            double dbPercent = 0;
            XmlDocument objDocument = null;

            

            try
            {
                objDocument = BYBCache.Instance.getFromCache(lastSystemWTCache);


                var nsm = new XmlNamespaceManager(objDocument.NameTable);
                nsm.AddNamespace("s", "http://www.sap.com/SBO/SDK/DI");
                string strXpath = "/s:Recordset/s:Rows/s:Row[./s:Fields/s:Field/s:Alias/text()='WTCode' and ./s:Fields/s:Field/s:Value/text()='" + WTCode + "']/s:Fields/s:Field[./s:Alias/text()='PrctBsAmnt']/s:Value";
                XmlNode oNode = objDocument.SelectSingleNode(strXpath,nsm);
                if(oNode != null)
                {
                    dbPercent = BYBHelpers.Instance.getStandarNumericValue(oNode.InnerText);

                }
                


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

            return dbPercent;
        }
    }
}
