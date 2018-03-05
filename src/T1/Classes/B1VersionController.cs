using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Xml;
using System.IO;
using System.Resources;
using System.Globalization;

using System.Collections;

namespace T1.Classes
{
    class B1VersionController
    {
        private bool blInstallationFound = false;

        private bool blResetConfiguration = false;
        




        public bool InstallationFound
        {
            get { return blInstallationFound; }
        }

        public bool ResetConfiguration
        {
            get { return blResetConfiguration; }
        }

        
        
        public B1VersionController()
        {
            
            
            try
            {
                if(BYBCache.Instance.getFromCache(CacheItemNames.Default.currentVersion) == null)
                {
                    //blInstallationFound = getInstallationInfo();
                }
            }
            catch(Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "B1VersionController", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        private bool getInstallationInfo()
        {
            bool blResult = false;
            
            SAPbobsCOM.Recordset objRecordset = null;
            SAPbobsCOM.Fields objFields = null;
            SAPbobsCOM.Field objField = null;
            string strQuery = "";
            

            XmlDocument objDocument = null;
            string strVersionXML = "";
            
            try
            {
                strQuery = string.Format(Classes.Queries.DBQueries.Default.getVersionControlSQL
                        , T1.Properties.Settings.Default.versionControlField
                        , T1.Properties.Settings.Default.versionControlSQLTable
                        , T1.Properties.Settings.Default.versionControlKey);
                if(BYBB1MainObject.Instance.B1Company != null &&  BYBB1MainObject.Instance.B1Company.Connected)
                {
                    try
                    {
                        objRecordset = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        objRecordset.DoQuery(strQuery);
                        if(objRecordset.RecordCount == 1)
                        {
                            objDocument = new XmlDocument();
                            objFields = objRecordset.Fields;
                            objField = objFields.Item(0);
                            strVersionXML = objField.Value;
                            objDocument.LoadXml(strVersionXML);
                            BYBCache.Instance.addToCache(T1.Properties.Settings.Default.VersionControlCacheName, objDocument, BYBCache.objCachePriority.NotRemovable);
                            BYBCache.Instance.addToCache(T1.Classes.CacheItemNames.Default.installationInfo, strVersionXML, BYBCache.objCachePriority.Default);
                            blResult = true;
                        }
                        else
                        {
                            blResetConfiguration = true;
                        }
                        
                        
                    }
                    catch (COMException comEx)
                    {
                        if (comEx.ErrorCode == -2000)
                        {
                            blResetConfiguration = true;
                        }
                        else
                        {

                            Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                            BYBExceptionHandling.reportException(er.Message, "B1VersionController.checkVersionTableExistance", er, 2, System.Diagnostics.EventLogEntryType.Information);
                        }
                        
                    }
                    catch (Exception er)
                    {
                        BYBExceptionHandling.reportException(er.Message, "B1VersionController.checkVersionTableExistance", er, 3, System.Diagnostics.EventLogEntryType.Error);
                    }
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "B1VersionController.checkVersionTableExistance", er, 2, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "B1VersionController.checkVersionTableExistance", er, 3, System.Diagnostics.EventLogEntryType.Error);
            }
            finally
            {
                if (objRecordset != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordset);
                    objRecordset = null;
                }

                if (objFields != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objFields);
                    objFields = null;
                }

                if (objField != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objField);
                    objField = null;
                }
            }

            return blResult;

        }

        public bool resetInstallInformation()
        {
            bool blResult = true;

            try
            {
                //blResult = createVersionMetaData();
                if(blResult)
                {
                    blResult = createAllModulesMetaData();
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "B1VersionController.resetInstallInformation", er, 2, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "B1VersionController.resetInstallInformation", er, 3, System.Diagnostics.EventLogEntryType.Error);
            }
            finally
            {
                /*if (objRecordset != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordset);
                    objRecordset = null;
                }

                if (objFields != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordset);
                    objRecordset = null;
                }

                if (objField != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordset);
                    objRecordset = null;
                }*/
            }





            return blResult;
        }

        private bool createAllModulesMetaData()
        {
            bool blResult = false;
            ResourceSet oResources = null;
            try
            {
                oResources = T1.Classes.Resources.MetaDataResources.ResourceManager.GetResourceSet(CultureInfo.CurrentUICulture, true, true);
                foreach (DictionaryEntry entry in oResources)
                {
                    if (entry.Key.ToString().IndexOf("MD") == 0)
                    {
                        blResult = createMetaData(entry.Value.ToString(), entry.Key.ToString());
                        if (!blResult)
                        {
                            break;
                        }
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "B1VersionController.resetInstallInformation", er, 2, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "B1VersionController.resetInstallInformation", er, 3, System.Diagnostics.EventLogEntryType.Error);
            }


            return blResult;
        }

        private bool createVersionMetaData()
        {
            bool blResult = false;
            string strMetaDataXML = "";
            string strXMLHeader = "";
            XmlDocument objDocument = null;

            try
            {
                BYBB1MainObject.Instance.B1Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                BYBB1MainObject.Instance.B1Company.XMLAsString = false;
                strMetaDataXML = "";// T1.Classes.Resources.MetaDataResources.VersionControlMetaData;
                strXMLHeader = T1.Properties.Resources.XMLHeader;
                objDocument = new XmlDocument();    
                objDocument.LoadXml(strMetaDataXML);
                XmlNodeList objTables = objDocument.SelectNodes(T1.Properties.Settings.Default.mdTablePath);
                XmlNodeList objUserFields = objDocument.SelectNodes(T1.Properties.Settings.Default.mdUserFieldsPath);
                
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
                            BYBExceptionHandling.reportException(er.Message, "B1VersionController.createVersionMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
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
                            BYBExceptionHandling.reportException(er.Message, "B1VersionController.createVersionMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
                        }
                    }

                }
                blResult = addDefaultConfigValue();
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "B1VersionController.createVersionMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "B1VersionController.createVersionMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

            return blResult;
        }



        private bool addDefaultConfigValue()
        {

            bool blResult = false;
            SAPbobsCOM.UserTables oUserTables = null;
            SAPbobsCOM.UserTable oUserTable = null;
            SAPbobsCOM.UserFields oUserFields = null;
            SAPbobsCOM.Fields oFields = null;
            SAPbobsCOM.Field oField = null;

            try
            {
                oUserTables = BYBB1MainObject.Instance.B1Company.UserTables;
                oUserTable = oUserTables.Item(T1.Properties.Settings.Default.versionControlB1Table);



                oUserTable.Code = T1.Properties.Settings.Default.versionControlKey;
                oUserTable.Name = T1.Properties.Settings.Default.versionControlKey;

                oUserFields = oUserTable.UserFields;
                oFields = oUserFields.Fields;
                oField = oFields.Item(T1.Properties.Settings.Default.versionControlField);
                oField.Value = T1.Properties.Resources.BaseVersionControl;
                if (oUserTable.Add() != 0)
                {
                    Exception er = new Exception(Convert.ToString("COM Error::" + BYBB1MainObject.Instance.B1Company.GetLastErrorCode().ToString() + "::" + BYBB1MainObject.Instance.B1Company.GetLastErrorDescription() + "::" ));
                    BYBExceptionHandling.reportException(er.Message, "B1VersionController.addDefaultConfigValue", er, 1, System.Diagnostics.EventLogEntryType.Error);
                    
                }
                else
                {
                    XmlDocument oTempDoc = new XmlDocument();
                    oTempDoc.LoadXml(T1.Properties.Resources.BaseVersionControl);
                    BYBCache.Instance.addToCache(T1.Properties.Settings.Default.VersionControlCacheName, oTempDoc, BYBCache.objCachePriority.NotRemovable);
                    blResult = true;
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
            finally
            {
                if (oUserTables != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTables);
                    oUserTables = null;
                }

                if (oUserTable != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                    oUserTable = null;
                }

                if (oUserFields != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFields);
                    oUserFields = null;
                }

                if (oFields != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oFields);
                    oUserFields = null;
                }

                if (oField != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oField);
                    oField = null;
                }
            }
            return blResult;
        }

        
        private bool createMetaData(string strXMLString, string strModule)
        {
           

            bool blResult = false;
            BYBB1MainObject.Instance.B1Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
            BYBB1MainObject.Instance.B1Company.XMLAsString = false;
            bool blContinue = false;
            string strXMLHeader = "";
            XmlDocument objDocument = null;
            XmlNodeList objTables = null;
            XmlNodeList objUserFields = null;
            XmlNodeList objUDO = null;

            try
            {
                
                strXMLHeader = T1.Properties.Resources.XMLHeader;
                objDocument = new XmlDocument();
                objDocument.LoadXml(strXMLString);
                objTables = objDocument.SelectNodes(T1.Properties.Settings.Default.mdTablePath);
                objUserFields = objDocument.SelectNodes(T1.Properties.Settings.Default.mdUserFieldsPath);
                objUDO = objDocument.SelectNodes(T1.Properties.Settings.Default.mdUDOPath);

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
                            blContinue = true;
                        }
                        else
                        {
                            Exception er = new Exception(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());
                            BYBExceptionHandling.reportException(er.Message, "createMetaData.createMetaData." + strModule, er, 1, System.Diagnostics.EventLogEntryType.Error);
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUMD);
                            objUMD = null;
                            //blContinue = false;
                            //break;
                        }
                    }

                }
                if (blContinue)
                {
                    blContinue = false;
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
                                blContinue = true;
                            }
                            else
                            {
                                Exception er = new Exception(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());
                                BYBExceptionHandling.reportException(er.Message, "createMetaData.createMetaData." + strModule, er, 1, System.Diagnostics.EventLogEntryType.Error);
                                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUMD);
                                objUMD = null;
                                //blContinue = false;
                                //break;
                            }


                        }

                    }
                }
                if (blContinue)
                {
                    blContinue = false;

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
                                blContinue = true;

                            }
                            else
                            {
                                Exception er = new Exception(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());
                                BYBExceptionHandling.reportException(er.Message, "createMetaData.createMetaData." + strModule, er, 1, System.Diagnostics.EventLogEntryType.Error);
                                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUMD);
                                objUMD = null;
                                //blContinue = false;
                                //break;
                            }


                        }

                    }
                }
                blResult = blContinue;



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "FSNotes.createMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "FSNotes.createMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }


            return blResult;
        }
         
    }
}
