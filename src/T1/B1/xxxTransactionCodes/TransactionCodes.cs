using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using T1.Classes;
using System.Runtime.InteropServices;

namespace T1.B1.TransactionCodes
{
    public class TransactionCodes
    {
        private static TransactionCodes objProjects = null;
        private static SAPbobsCOM.TransactionCodesService objTransactionCodesService = null;

        private TransactionCodes()
        {
            SAPbobsCOM.CompanyService objCompanyService = null;

            
        
            try
            {
                objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                objTransactionCodesService = objCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.TransactionCodesService);
                
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "TransactionCodes", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "TransactionCodes", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static private SAPbobsCOM.TransactionCodeParamsCollection getTransactionsCodesList()
        {

            SAPbobsCOM.TransactionCodeParamsCollection objTransactionCodesList = null;

            if (objProjects == null)
                objProjects = new TransactionCodes();

            try
            {
                try
                {
                    objTransactionCodesList = objTransactionCodesService.GetList();

                }
                catch (COMException comEx)
                {
                    if (comEx.ErrorCode == -2028)
                    {

                    }
                    else
                    {
                        Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                        BYBExceptionHandling.reportException(er.Message, "TransactionCodes.getTransactionsCodesList", er, 1, System.Diagnostics.EventLogEntryType.Error);
                    }
                    return null;
                }
                return objTransactionCodesList;
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "TransactionCodes.getTransactionsCodesList", er, 1, System.Diagnostics.EventLogEntryType.Error);
                return null;
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "TransactionCodes.getTransactionsCodesList", er, 1, System.Diagnostics.EventLogEntryType.Error);
                return null;
            }
            

        }

        static public void fillValidValuesFromDB(ref SAPbouiCOM.ComboBox oCombo)
        {
            SAPbobsCOM.TransactionCodeParamsCollection objTransactionParams = null;
            XmlDocument strXMLProjectList = null;
            
            string strXPath = "/TransactionCodeParamsCollection/TransactionCodeParams";
            XmlNodeList objNodes = null;

            if (objProjects == null)
                objProjects = new TransactionCodes();
            try
            {
                

                strXMLProjectList = new XmlDocument();

                objTransactionParams = getTransactionsCodesList();
                if (objTransactionParams != null)
                {
                    strXMLProjectList.LoadXml(objTransactionParams.ToXMLString());
                    if(oCombo.ValidValues.Count > 0)
                    {
                        while(oCombo.ValidValues.Count > 0)
                        {
                            oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }

                    }
                    objNodes =  strXMLProjectList.SelectNodes(strXPath);
                    if(objNodes != null)
                    {
                        foreach(XmlNode xn in objNodes)
                        {
                            string strCode = xn.SelectSingleNode("./Code").InnerText;
                            string strDescription = xn.SelectSingleNode("./Description").InnerText;
                            oCombo.ValidValues.Add(strCode,strDescription);

                        }
                    }
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "TransactionCodes.fillValidValuesFromDB", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "TransactionCodes.fillValidValuesFromDB", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

       
        

    }
}
