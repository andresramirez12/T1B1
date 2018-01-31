using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;
using System.IO;
using System.Xml;

namespace T1.B1.MetaData
{
    public class Operations
    {
        static private Operations objMD = null;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private Operations()
        {

        }

        static private void insertXML(string strPath)
        {
            SAPbobsCOM.UserTablesMD objUTMD = null;
            SAPbobsCOM.UserFieldsMD objUFMD = null;
            SAPbobsCOM.UserObjectsMD objUOMD = null;
            

            try
            {
                
                    T1.B1.MainObject.Instance.B1Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                
                int iRes = -1;
                    int intTotal = T1.B1.MainObject.Instance.B1Company.GetXMLelementCount(strPath);
                    for (int i = 0; i < intTotal; i++)
                    {
                        if (T1.B1.MainObject.Instance.B1Company.GetXMLobjectType(strPath, i) == SAPbobsCOM.BoObjectTypes.oUserTables)
                        {
                        #region Create Tables
                        try
                        {
                            objUTMD = T1.B1.MainObject.Instance.B1Company.GetBusinessObjectFromXML(strPath, i);
                            if (!isTableCreated(objUTMD.TableName))
                            {

                                iRes = objUTMD.Add();
                                if (iRes != 0 && iRes != -2035)
                                {
                                    Exception er = new Exception("Could not create MD:" + objUTMD.TableName);
                                    _Logger.Error("Could not create MD " + objUTMD.TableName, er);
                                    break;
                                }
                                iRes = -1;

                            }
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUTMD);
                            objUTMD = null;
                        }
                        catch(Exception er)
                        {
                            _Logger.Error(strPath, er);
                            if(objUTMD != null)
                            {
                                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUTMD);
                                objUTMD = null;
                            }
                        }
                        #endregion
                    }
                        else if (T1.B1.MainObject.Instance.B1Company.GetXMLobjectType(strPath, i) == SAPbobsCOM.BoObjectTypes.oUserFields)
                        {
                        #region create Fields
                        try { 
                            objUFMD = T1.B1.MainObject.Instance.B1Company.GetBusinessObjectFromXML(strPath, i);
                            if (!isFieldCreated(objUFMD.Name, objUFMD.TableName))
                            {
                                iRes = objUFMD.Add();
                                if (iRes != 0 && iRes != -2035)
                                {
                                    Exception er = new Exception("Could not create MD:" + objUFMD.TableName + " Field:" + objUFMD.Name);
                                    string strError = T1.B1.MainObject.Instance.B1Company.GetLastErrorDescription();
                                    _Logger.Error("Could not create MD " + objUFMD.Name, er);
                                    break;
                                }
                                iRes = -1;

                            }
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUFMD);
                            objUFMD = null;
                        }
                        catch (Exception er)
                        {
                            _Logger.Error(strPath, er);
                            if (objUFMD != null)
                            {
                                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUFMD);
                                objUFMD = null;
                            }
                        }
                        #endregion
                    }
                        else if (T1.B1.MainObject.Instance.B1Company.GetXMLobjectType(strPath, i) == SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                        {
                        #region Create UDOs
                        try
                        {
                            objUOMD = T1.B1.MainObject.Instance.B1Company.GetBusinessObjectFromXML(strPath, i);
                            if (!isUDOCreated(objUOMD.Name))
                            {
                                iRes = objUOMD.Add();
                                if (iRes != 0 && iRes != -2035 && iRes != -5002)
                                {
                                    Exception er = new Exception("Could not create UDO:" + objUOMD.Code + T1.B1.MainObject.Instance.B1Company.GetLastErrorDescription());
                                    _Logger.Error("Could not create UDO " + objUOMD.Name, er);
                                    break;
                                }
                                iRes = -1;

                            }
                            else
                            {
                                if (Settings._Main.updateUDOForm)
                                {
                                    string strSRF = objUOMD.FormSRF;
                                    string strCode = objUOMD.Code;
                                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUOMD);
                                    objUOMD = null;

                                    objUOMD = T1.B1.MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                                    if (objUOMD.GetByKey(strCode))
                                    {
                                        objUOMD.FormSRF = strSRF;
                                        iRes = objUOMD.Update();
                                        if (iRes != 0)
                                        {
                                            Exception er = new Exception("Could not update UDO:" + objUOMD.Code + T1.B1.MainObject.Instance.B1Company.GetLastErrorDescription());
                                            _Logger.Error("Could not update UDO " + objUOMD.Name, er);
                                            break;
                                        }
                                    }
                                    T1.Config.Settings._T1B1Metadata.updateUDOForm = false;
                                    T1.Config.Settings._T1B1Metadata.Write();
                                }
                            }
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUOMD);
                            objUOMD = null;
                        }
                        catch (Exception er)
                        {
                            _Logger.Error(strPath, er);
                            if (objUOMD != null)
                            {
                                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUOMD);
                                objUOMD = null;
                            }
                        }
                        #endregion

                    }

                    }
                
            }
            catch(Exception er)
            {
                _Logger.Error(strPath, er);
                
            }
        }
        static public bool blCreateMD(bool CreateMD)
        {
            if (objMD == null)
                objMD = new Operations();

            bool blResult = false;
            string strPath = "";
            
            try
            {
                if (CreateMD)
                {

                    strPath = AppDomain.CurrentDomain.BaseDirectory + Settings._Main.resourceFodler;

                    if (Settings._Main.singleFile)
                    {
                        T1.B1.MainObject.Instance.B1Company.XMLAsString = false;

                        insertXML(strPath + "\\" + Settings._Main.singleFileName);

                        T1.B1.MainObject.Instance.B1Company.XMLAsString = true;
                    }
                    else
                    {
                        string[] strTableFiles = Directory.GetFiles(strPath, "TABLE*.xml");
                        string[] strFieldFiles = Directory.GetFiles(strPath, "FIELD*.xml");
                        string[] strUDOFiles = Directory.GetFiles(strPath, "UDO*.xml");
                        List<string[]> udoConfig = new List<string[]>();
                        udoConfig.Add(strTableFiles);
                        udoConfig.Add(strFieldFiles);
                        udoConfig.Add(strUDOFiles);

                        for (int i = 0; i < udoConfig.Count; i++)
                        {
                            string[] strTemp = udoConfig[i];
                            T1.B1.MainObject.Instance.B1Company.XMLAsString = false;
                            for (int k = 0; k < strTemp.Length; k++)
                            {
                                insertXML(strTemp[k]);
                            }
                            T1.B1.MainObject.Instance.B1Company.XMLAsString = true;
                        }
                    }
                }


                blResult = true;

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("COM Error", er);
            }
            catch (Exception er)
            {
                _Logger.Error("MD Creation Error", er);
            }
            return blResult;
        }

        static private bool isTableCreated(string TableName)
        {
            bool blResult = false;
            SAPbobsCOM.Recordset objRecord = null;
            try
            {
                objRecord = T1.B1.MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string strSQL = "select count(TableName) from OUTB where TableName='" + TableName + "'";
                objRecord.DoQuery(strSQL);
                if (objRecord.RecordCount > 0)
                {
                    int intCount = objRecord.Fields.Item(0).Value;
                    if (intCount == 1)
                    {
                        blResult = true;
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("COM Error", er);


            }
            catch (Exception er)
            {
                _Logger.Error("MD Creation Error", er);
            }
            finally
            {
                if (objRecord != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objRecord);
                    objRecord = null;
                }
            }

            return blResult;
        }

        static private bool isFieldCreated(string FieldName, string TableName)
        {
            bool blResult = false;
            SAPbobsCOM.Recordset objRecord = null;
            try
            {
                objRecord = T1.B1.MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string strSQL = "select count(AliasID) from CUFD where AliasID='" + FieldName + "' and TableID='" + TableName + "'";
                objRecord.DoQuery(strSQL);
                if (objRecord.RecordCount > 0)
                {
                    int intCount = objRecord.Fields.Item(0).Value;
                    if (intCount == 1)
                    {
                        blResult = true;
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("COM Error", er);


            }
            catch (Exception er)
            {
                _Logger.Error("MD Creation Error", er);
            }
            finally
            {
                if (objRecord != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objRecord);
                    objRecord = null;
                }
            }

            return blResult;
        }

        static private bool isUDOCreated(string UDOName)
        {
            bool blResult = false;
            SAPbobsCOM.Recordset objRecord = null;
            try
            {
                objRecord = T1.B1.MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string strSQL = "select count(Name) from OUDO where Name='" + UDOName + "'";
                objRecord.DoQuery(strSQL);
                if (objRecord.RecordCount > 0)
                {
                    int intCount = objRecord.Fields.Item(0).Value;
                    if (intCount == 1)
                    {
                        blResult = true;
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("COM Error", er);


            }
            catch (Exception er)
            {
                _Logger.Error("MD Creation Error", er);
            }
            finally
            {
                if (objRecord != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objRecord);
                    objRecord = null;
                }
            }

            return blResult;
        }

        static public void loadMuni()
        {
            
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService UDOService = null;
            SAPbobsCOM.GeneralData headerInfo = null;
            SAPbobsCOM.GeneralDataParams addResult = null;
            SAPbobsCOM.GeneralCollectionParams objList = null;

            try
            {

                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                UDOService = objCompanyService.GetGeneralService(Settings._Main.loadMuniUDO);
                headerInfo = UDOService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                try
                {
                    objList = UDOService.GetList();
                }
                catch(COMException comEx)
                {
                    if(comEx.ErrorCode == -2028)
                    {
                        objList = null;
                    }
                }
                catch(Exception er)
                {
                    _Logger.Error("", er);
                    objList = null;
                }
                if (objList == null || objList.Count == 0)
                {
                    using (StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "LoadData\\" + Settings._Main.loadMuniUDO))
                    {
                        while (!sr.EndOfStream)
                        {
                            string strLine = sr.ReadLine();
                            string[] strResult = strLine.Split('|');
                            headerInfo.SetProperty("Code", strResult[0].Trim());
                            headerInfo.SetProperty("Name", strResult[1].Trim());
                            addResult = UDOService.Add(headerInfo);

                        }
                    }
                }

                
            }
            catch (Exception er)
            {

                _Logger.Error("", er);
            }

        }

        static public void loadDepto()
        {

            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService UDOService = null;
            SAPbobsCOM.GeneralData headerInfo = null;
            SAPbobsCOM.GeneralDataParams addResult = null;
            SAPbobsCOM.GeneralCollectionParams objList = null;

            try
            {

                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                UDOService = objCompanyService.GetGeneralService(Settings._Main.loadDeptUDO);
                headerInfo = UDOService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                try
                {
                    objList = UDOService.GetList();
                }
                catch (COMException comEx)
                {
                    if (comEx.ErrorCode == -2028)
                    {
                        objList = null;
                    }
                }
                catch (Exception er)
                {
                    _Logger.Error("", er);
                    objList = null;
                }
                if (objList == null || objList.Count == 0)
                {
                    using (StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "LoadData\\" + Settings._Main.loadDeptUDO))
                    {
                        while (!sr.EndOfStream)
                        {
                            string strLine = sr.ReadLine();
                            string[] strResult = strLine.Split('|');
                            headerInfo.SetProperty("Code", strResult[0].Trim());
                            headerInfo.SetProperty("Name", strResult[1].Trim());
                            addResult = UDOService.Add(headerInfo);

                        }
                    }
                }


            }
            catch (Exception er)
            {

                _Logger.Error("", er);
            }

        }

        static public void loadGenericUDO()
        {

            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService UDOService = null;
            SAPbobsCOM.GeneralData headerInfo = null;
            
            SAPbobsCOM.GeneralDataParams oResult = null;
            SAPbobsCOM.GeneralCollectionParams objList = null;
            XmlDocument oXMLLod = new XmlDocument();

            try
            {
                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "LoadData\\AutomaticLoad.xml"))
                {
                    oXMLLod.Load(AppDomain.CurrentDomain.BaseDirectory + "LoadData\\AutomaticLoad.xml");
                    XmlNodeList oList = oXMLLod.SelectNodes("/LoadFiles/File");
                    if (oList != null && oList.Count > 0)
                    {
                        foreach (XmlNode oFileNode in oList)
                        {
                            string strFilePath = oFileNode.SelectSingleNode("definition/location").InnerText;
                            string strUDOName = oFileNode.SelectSingleNode("definition/UDO").InnerText;

                            objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                            UDOService = objCompanyService.GetGeneralService(strUDOName);
                            headerInfo = UDOService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                            try
                            {
                                objList = UDOService.GetList();
                            }

                            catch (COMException comEx)
                            {
                                if (comEx.ErrorCode == -2028)
                                {
                                    objList = null;
                                    _Logger.Error("The UDO with name " + strUDOName + " was not found");
                                }
                            }
                            catch (Exception er)
                            {
                                _Logger.Error("", er);
                                objList = null;
                            }

                            using (StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + strFilePath))
                            {
                                while (!sr.EndOfStream)
                                {
                                    string strLine = sr.ReadLine();
                                    string[] strResult = strLine.Split('|');
                                    bool runAdd = false;
                                    
                                    string strCode = "";
                                    
                                    for (int i = 0; i < strResult.Length; i++)
                                    {
                                        string strPropertyName = oFileNode.SelectSingleNode("mapping/column[@index='" + i.ToString() + "']") != null ? oFileNode.SelectSingleNode("mapping/column[@index='" + i.ToString() + "']/@property").InnerText : "";
                                        if (i == 0)
                                        {
                                            oResult = UDOService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                            strCode = strResult[i];
                                            oResult.SetProperty("Code", strResult[i]);
                                            try
                                            {
                                                headerInfo = UDOService.GetByParams(oResult);
                                            }
                                            catch (COMException comEx)
                                            {
                                                if (comEx.ErrorCode == -2028)
                                                {
                                                    runAdd = true;
                                                    headerInfo = UDOService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                                                    headerInfo.SetProperty("Code", strResult[i].Trim());
                                                }
                                                else
                                                {
                                                    headerInfo = null;
                                                }
                                            }
                                            catch (Exception er)
                                            {
                                                _Logger.Error("", er);
                                                headerInfo = null;
                                            }

                                        }
                                        else
                                        {

                                            if (strPropertyName.Trim().Length > 0 && headerInfo != null)
                                            {
                                                headerInfo.SetProperty(strPropertyName, strResult[i].Trim());


                                            }
                                        }
                                    }
                                    if(runAdd)
                                    {
                                        oResult = UDOService.Add(headerInfo);
                                        _Logger.Info(strCode + " added OK");
                                    }
                                    else
                                    {
                                        try
                                        {
                                            UDOService.Update(headerInfo);
                                            _Logger.Info(strCode + " updated OK");
                                        }
                                        catch (COMException comEx)
                                        {
                                            _Logger.Error("Could not update value " + strCode + " of UDO " + strUDOName, comEx);
                                        }
                                        catch (Exception er)
                                        {
                                            _Logger.Error("Could not update value " + strCode + " of UDO " + strUDOName, er);
                                        }
                                    }


                                }
                            }
                        }
                    }
                    
                }
                
            }
            catch (Exception er)
            {

                _Logger.Error("", er);
            }

        }

    }
}
