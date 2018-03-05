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
    public class BusinessPartners
    {

        static private BusinessPartners objBPsHelpers = null;
        static private Hashtable hashWTCodesforBPCahed = null;

        private BusinessPartners()
        {
            hashWTCodesforBPCahed = new Hashtable();
        }

        static public void cleanCacheWTCodesForBP()
        {
            if(objBPsHelpers == null)
            {
                objBPsHelpers = new BusinessPartners();
            }

            try
            {
                foreach(string strKey in hashWTCodesforBPCahed.Keys)
                {
                    BYBCache.Instance.removeFromCache(strKey);
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
        }

        static public void cacheHashWTCodesForBP(string BP, string CacheName)
        {

            if (objBPsHelpers == null)
            {
                objBPsHelpers = new BusinessPartners();
            }

            SAPbobsCOM.BusinessPartners objBP = null;
            XmlDocument objDocument = null;
            XmlNode objWTNode = null;
            Hashtable BPHash = null;
            
            try
            {
                objBP = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                if (objBP.GetByKey(BP))
                {
                    BPHash = new Hashtable();
                    objDocument = new XmlDocument();
                    objDocument.LoadXml(objBP.GetAsXML());
                    objWTNode = objDocument.SelectSingleNode("/BOM/BO/BPWithholdingTax");
                    if (objWTNode != null)
                    {
                        BYBCache.Instance.addToCache(CacheName, objWTNode, BYBCache.objCachePriority.Default);
                        if(!hashWTCodesforBPCahed.ContainsKey(CacheName))
                        {
                            hashWTCodesforBPCahed.Add(CacheName, CacheName);
                        }
                    }
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
        }
    }
}
