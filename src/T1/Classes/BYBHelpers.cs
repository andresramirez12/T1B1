using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Runtime.InteropServices;
using System.IO;
using System.Globalization;
using System.Windows.Forms;
using System.Threading;

namespace T1.Classes
{
    public sealed class BYBHelpers
    {

        private static readonly Lazy<BYBHelpers> lazy =
            new Lazy<BYBHelpers>(() => new BYBHelpers());

        private CultureInfo enUsCulture = CultureInfo.GetCultureInfo("en-US");

        

        public static BYBHelpers Instance
        {
            get
            {
                


                return lazy.Value;
            }
        }

        public double getDoubleFromMoney(string MoneyString)
        {
            double dbResult = 0;
            string strTempValue = "";
            CultureInfo ownCultrue = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            try
            {

                objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                SAPbobsCOM.AdminInfo objCompany = objCompanyService.GetAdminInfo();
                string strDecimal = objCompany.DecimalSeparator;
                string strMiles = objCompany.ThousandsSeparator;

                ownCultrue = new CultureInfo("en-US");
                ownCultrue.NumberFormat.NumberDecimalSeparator = ".";
                ownCultrue.NumberFormat.NumberGroupSeparator = ",";

                string[] strValue = MoneyString.Split(' ');
                if (strValue.Length > 0)
                {
                    for (int i = 0; i < strValue.Length; i++)
                    {
                        string strTemp = strValue[i];
                        if (strTemp.IndexOf(strDecimal) >= 0)
                        {
                            strTempValue = strTemp;
                            break;
                        }
                    }
                }
                if (strDecimal != ".")
                {
                    strTempValue = strTempValue.Replace(strMiles, "");
                    strTempValue = strTempValue.Replace(strDecimal, ".");

                }

                dbResult = Double.Parse(strTempValue, ownCultrue);
            }
            catch(Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.objButton_PressedAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return dbResult;
        }

        public string getVisualFormatFromDouble(double DoubleFormat)
        {
            string strResult = "";
            
            CultureInfo ownCultrue = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            try
            {

                objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                SAPbobsCOM.AdminInfo objCompany = objCompanyService.GetAdminInfo();
                string strDecimal = objCompany.DecimalSeparator;
                string strThousands = objCompany.ThousandsSeparator;
                int numberDecimals = objCompany.TotalsAccuracy;

                ownCultrue = new CultureInfo("es-CO");
                //ownCultrue.NumberFormat.NumberDecimalSeparator = ",";
                ownCultrue.NumberFormat.NumberDecimalSeparator = strDecimal;
                ownCultrue.NumberFormat.NumberGroupSeparator = strThousands;
                ownCultrue.NumberFormat.NumberDecimalDigits = numberDecimals;

                ownCultrue.NumberFormat.CurrencyDecimalSeparator = strDecimal;
                ownCultrue.NumberFormat.CurrencyGroupSeparator = strThousands;
                ownCultrue.NumberFormat.CurrencyDecimalDigits = numberDecimals;
                ownCultrue.NumberFormat.CurrencySymbol = "";

                strResult = DoubleFormat.ToString( "C"+ numberDecimals.ToString(),ownCultrue);
                //if(strResult.IndexOf(".") > 0)
                //{
                //    strResult = strResult.Replace(".", ",");
                //}
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.objButton_PressedAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strResult;
        }

        public string getClipboardString()
        {
            string strTemp = "";

            try
            {
                IDataObject idat = null;
                Exception threadEx = null;
                Thread staThread = new Thread(
                    delegate ()
                    {
                        try
                        {
                            strTemp = Clipboard.GetText();
                        }

                        catch (Exception ex)
                        {
                            threadEx = ex;
                        }
                    });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();

                

            }
            catch(Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.objButton_PressedAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

            return strTemp;
        }

        public XmlDocument getClipboardStringAsXML()
        {
            XmlDocument xmlDocument = null;
            string strTemp = "";

            try
            {
                Exception threadEx = null;
                Thread staThread = new Thread(
                    delegate ()
                    {
                        try
                        {
                            strTemp = Clipboard.GetText();
                        }

                        catch (Exception ex)
                        {
                            threadEx = ex;
                        }
                    });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();

                string[] strLines = strTemp.Split('\n');
                for (int i = 0; i < strLines.Length; i++)
                {

                }


            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.objButton_PressedAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

            return xmlDocument;
        }

        public DateTime getDateTimeFormString(string strDate)
        {
            return new DateTime(Convert.ToInt32(strDate.Substring(0, 4)), Convert.ToInt32(strDate.Substring(4, 2)), Convert.ToInt32(strDate.Substring(6, 2)));
        } 


        public double getStandarNumericValue(string Value)
        {

            double finalValue = -1;
            string strTemporalValue = Value;
            string strDecimal = BYBCache.Instance.getFromCache(CacheItemNames.Default.decimals);
            string strMiles = BYBCache.Instance.getFromCache(CacheItemNames.Default.thousands);

            CultureInfo objTest = new CultureInfo("en-US");
            objTest.NumberFormat.NumberDecimalSeparator = ".";
            objTest.NumberFormat.NumberGroupSeparator = ",";

            try
            {
                //if (strDecimal != ".")
                //{
                  //  strTemporalValue = strTemporalValue.Replace(strMiles, "");
                    //strTemporalValue = strTemporalValue.Replace(strDecimal, ".");
                    //finalValue = double.Parse(strTemporalValue, enUsCulture);
                

                //}
                //else
                //{
                finalValue = double.Parse(strTemporalValue, objTest);
                //}
            }
            catch(Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.objButton_PressedAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return finalValue;
        }

        public void getInstallStatus()
        {
            XmlDocument objDocument = null;
            SAPbobsCOM.UserTables oUserTables = null;
            SAPbobsCOM.UserTable oConfigTable = null;
            string strConfigValue = "";
            
            try
            {
                objDocument = new XmlDocument();
                oUserTables = BYBB1MainObject.Instance.B1Company.UserTables;
                try
                {
                    oConfigTable = oUserTables.Item(T1.Properties.Settings.Default.versionControlB1Table);
                    if (oConfigTable.GetByKey(T1.Properties.Settings.Default.versionControlKey))
                    {
                        strConfigValue = oConfigTable.UserFields.Fields.Item(T1.Properties.Settings.Default.versionControlField).Value;
                        XmlDocument oTempDoc = new XmlDocument();
                        oTempDoc.LoadXml(strConfigValue);
                        BYBCache.Instance.addToCache(T1.Properties.Settings.Default.VersionControlCacheName, oTempDoc, BYBCache.objCachePriority.NotRemovable);
                    }
                    else
                    {
                        addDefaultConfigValue();
                    }
                   
                    

                }
                catch (COMException comEx)
                {
                    if(comEx.ErrorCode == -1106)
                    {
                        createInstallationControlTable();
                    }
                }
                catch (Exception er)
                {
                    BYBExceptionHandling.reportException(er.Message, "BYBHelpers.addNSReports", er, 1, System.Diagnostics.EventLogEntryType.Error);
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BYBHelpers.addNSReports", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BYBHelpers.addNSReports", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        private bool createInstallationControlTable()
        {
            bool blResult = false;

            BYBB1MainObject.Instance.B1Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
            BYBB1MainObject.Instance.B1Company.XMLAsString = false;

            try
            {
                string strMetaDataXML = T1.Properties.Resources.MetaDataCreation;
                string strXMLHeader = T1.Properties.Resources.XMLHeader;
                XmlDocument objDocument = new XmlDocument();
                objDocument.LoadXml(strMetaDataXML);
                XmlNodeList objTables = objDocument.SelectNodes(T1.Properties.Settings.Default.mdTablePath);
                XmlNodeList objUserFields = objDocument.SelectNodes(T1.Properties.Settings.Default.mdUserFieldsPath);
                XmlNodeList objUDO = objDocument.SelectNodes(T1.Properties.Settings.Default.mdUDOPath);

                if (objTables != null && objTables.Count > 0)
                {
                    foreach (XmlNode xn in objTables)
                    {
                        string strXML = strXMLHeader + xn.InnerXml;
                        using (StreamWriter sr = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\temp.xml", false))
                        {
                            sr.Write(strXML);
                        }
                        SAPbobsCOM.UserTablesMD objUMD = BYBB1MainObject.Instance.B1Company.GetBusinessObjectFromXML(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\temp.xml", 0);
                        int iResult = objUMD.Add();
                        if (iResult == 0 || iResult == -2035)
                        {
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUMD);
                            objUMD = null;
                        }
                        else
                        {
                            Exception er = new Exception(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());
                            BYBExceptionHandling.reportException(er.Message, "BYBHelpers.createInstallationControlTable", er, 1, System.Diagnostics.EventLogEntryType.Error);
                        }



                    }

                }
                if (objUserFields != null && objUserFields.Count > 0)
                {
                    foreach (XmlNode xn in objUserFields)
                    {
                        string strXML = strXMLHeader + xn.InnerXml;
                        using (StreamWriter sr = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\temp.xml", false))
                        {
                            sr.Write(strXML);
                        }
                        SAPbobsCOM.UserFieldsMD objUMD = BYBB1MainObject.Instance.B1Company.GetBusinessObjectFromXML(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\temp.xml", 0);
                        int iResult = objUMD.Add();
                        if (iResult == 0 || iResult == -2035)
                        {
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUMD);
                            objUMD = null;
                        }
                        else
                        {
                            Exception er = new Exception(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());
                            BYBExceptionHandling.reportException(er.Message, "BYBHelpers.createInstallationControlTable", er, 1, System.Diagnostics.EventLogEntryType.Error);
                        }


                    }

                }
                if (objUDO != null && objUDO.Count > 0)
                {
                    foreach (XmlNode xn in objUDO)
                    {
                        string strXML = strXMLHeader + xn.InnerXml;
                        using (StreamWriter sr = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\temp.xml", false))
                        {
                            sr.Write(strXML);
                        }
                        SAPbobsCOM.UserObjectsMD objUMD = BYBB1MainObject.Instance.B1Company.GetBusinessObjectFromXML(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\temp.xml", 0);
                        int iResult = objUMD.Add();
                        if (iResult == 0 || iResult == -2035)
                        {
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUMD);
                            objUMD = null;
                        }
                        else
                        {
                            Exception er = new Exception(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());
                            BYBExceptionHandling.reportException(er.Message, "BYBHelpers.createInstallationControlTable", er, 1, System.Diagnostics.EventLogEntryType.Error);
                        }


                    }

                }

                addDefaultConfigValue();
                


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BYBHelpers.createInstallationControlTable", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BYBHelpers.createInstallationControlTable", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }


            return blResult;



            
        }


        static private void addDefaultConfigValue()
        {

            SAPbobsCOM.UserTables oUserTables = null;
            SAPbobsCOM.UserTable oUserTable = null;

            try
            {
                oUserTables = BYBB1MainObject.Instance.B1Company.UserTables;
                oUserTable = oUserTables.Item(T1.Properties.Settings.Default.versionControlSQLTable);

                oUserTable.Code = T1.Properties.Settings.Default.versionControlKey;
                oUserTable.Name = T1.Properties.Settings.Default.versionControlKey;
                oUserTable.UserFields.Fields.Item(T1.Properties.Settings.Default.versionControlField).Value = T1.Properties.Resources.BaseVersionControl;
                if (oUserTable.Add() != 0)
                {
                    Exception er = new Exception(Convert.ToString("COM Error::" + BYBB1MainObject.Instance.B1Company.GetLastErrorCode().ToString() + "::" + BYBB1MainObject.Instance.B1Company.GetLastErrorDescription() + "::" ));

                }
                else
                {
                    XmlDocument oTempDoc = new XmlDocument();
                    oTempDoc.LoadXml(T1.Properties.Resources.BaseVersionControl);
                    BYBCache.Instance.addToCache(T1.Properties.Settings.Default.VersionControlCacheName, oTempDoc, BYBCache.objCachePriority.NotRemovable);
                }


                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "FSNotes.addDefaultConfigValue", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "FSNotes.addDefaultConfigValue", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        public void getB1LocalConfiguration()
        {
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.AdminInfo objAdmInfo = null;
            try
            {
                objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                objAdmInfo = objCompanyService.GetAdminInfo();
                BYBCache.Instance.addToCache(CacheItemNames.Default.decimals, objAdmInfo.DecimalSeparator, BYBCache.objCachePriority.NotRemovable);
                BYBCache.Instance.addToCache(CacheItemNames.Default.thousands, objAdmInfo.ThousandsSeparator, BYBCache.objCachePriority.NotRemovable);
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BYBHelpers.getB1LocalCOnfiguration", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BYBHelpers.getB1LocalCOnfiguration", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            finally
            {
                if(objAdmInfo != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objAdmInfo);
                    objAdmInfo = null;
                }

                if (objAdmInfo != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objAdmInfo);
                    objCompanyService = null;
                }
            }
        }


        static public void getAllCurrencies()
        {
            SAPbobsCOM.Recordset objRecordset = null;
            SAPbobsCOM.Currencies objCurrencies = null;
            string strSql = "";
            Hashtable objHashInformation = null;
            //ArrayList objArrayList = null;
            

            try
            {
                objRecordset = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                strSql = "select CurrCode from OCRN";
                objRecordset.DoQuery(strSql);
                if (objRecordset.RecordCount > 0)
                {
                    objHashInformation = new Hashtable();
                    objCurrencies = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCurrencyCodes);
                    objCurrencies.Browser.Recordset = objRecordset;
                    objCurrencies.Browser.MoveFirst();
                    while (!objCurrencies.Browser.EoF)
                    {
                        objHashInformation.Add(objCurrencies.Code, objCurrencies.Name);
                        objCurrencies.Browser.MoveNext();
                    }
                }
                
                BYBCache.Instance.addToCache("MoneyCodes", objHashInformation, BYBCache.objCachePriority.Default);

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
