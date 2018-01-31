using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;
using System.Xml;
using System.Globalization;
using System.Drawing;

namespace T1.B1.ReletadParties
{
    public class Instance
    {
        //private static Instance objInstance;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private Instance()
        {
            
        }
        #region Related Parties UDO Form
        public static void loadRelatedPartiesUDOForm()
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.FormCreationParams objParams = null;
            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objParams.XmlData = ReletadParties.RelatedPartiesRes.BYB_Terceros_Relacionados;
                objParams.FormType = "BYB_T1RPA100UDO";
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objForm.VisibleEx = true;

            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void addInsertRowRelationMenuUDO(SAPbouiCOM.Form objForm, SAPbouiCOM.ContextMenuInfo eventInfo)
        {
            SAPbouiCOM.MenuCreationParams objParams = null;

            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objParams.String = "Agregar línea";
                objParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                objParams.UniqueID = "BYB_MRPARU";
                objParams.Position = -1;
                objForm.Menu.AddEx(objParams);
                EventInfoClass objEvent = new EventInfoClass();
                objEvent.ColUID = eventInfo.ColUID;
                objEvent.FormUID = eventInfo.FormUID;
                objEvent.ItemUID = eventInfo.ItemUID;
                objEvent.Row = eventInfo.Row;
                CacheManager.CacheManager.Instance.addToCache(Settings._Main.lastRightClickEventInfo, objEvent, CacheManager.CacheManager.objCachePriority.Default);



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void removeInsertRowRelationMenuUDO()
        {


            try
            {
                if (MainObject.Instance.B1Application.Menus.Exists("BYB_MRPARU"))
                {
                    MainObject.Instance.B1Application.Menus.RemoveEx("BYB_MRPARU");
                }
                CacheManager.CacheManager.Instance.removeFromCache(Settings._Main.lastRightClickEventInfo);


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void addDeleteRowRelationMenuUDO(SAPbouiCOM.Form objForm, SAPbouiCOM.ContextMenuInfo eventInfo)
        {
            SAPbouiCOM.MenuCreationParams objParams = null;

            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objParams.String = "Eliminar línea";
                objParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                objParams.UniqueID = "BYB_MRPDRU";
                objParams.Position = -1;
                objForm.Menu.AddEx(objParams);
                EventInfoClass objEvent = new EventInfoClass();
                objEvent.ColUID = eventInfo.ColUID;
                objEvent.FormUID = eventInfo.FormUID;
                objEvent.ItemUID = eventInfo.ItemUID;
                objEvent.Row = eventInfo.Row;
                CacheManager.CacheManager.Instance.addToCache(Settings._Main.lastRightClickEventInfo, objEvent, CacheManager.CacheManager.objCachePriority.Default);


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void removeDeleteRowRelationMenuUDO()
        {


            try
            {
                if (MainObject.Instance.B1Application.Menus.Exists("BYB_MRPDRU"))
                {
                    MainObject.Instance.B1Application.Menus.RemoveEx("BYB_MRPDRU");
                }
                CacheManager.CacheManager.Instance.removeFromCache(Settings._Main.lastRightClickEventInfo);


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void relatedPartiedMatrixOperationUDO(EventInfoClass eventInfo, string Action)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;
            
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                objMatrix = objForm.Items.Item("Item_1").Specific;
                
                int intRow = eventInfo.Row;
                switch (Action)
                {
                    case "Add":
                        objMatrix.AddRow(1, intRow);

                        objMatrix.SetCellWithoutValidation(intRow + 1, "Col_0", "");
                        objMatrix.FlushToDataSource();

                        objMatrix.SetCellFocus(intRow + 1, 1);


                        break;
                    case "Delete":
                        objMatrix.DeleteRow(intRow);
                        objMatrix.FlushToDataSource();
                        break;

                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
        #endregion

        #region Related Parties BP Form

        static public void BYBRelatedPartiesFolderAdd(string strFormUID)
        {

            SAPbouiCOM.Form objForm = null;
            int intLeft = 0;
            string strUID = "";
            SAPbouiCOM.Item objItemBase = null;
            SAPbouiCOM.Item objItem = null;
            SAPbouiCOM.Matrix objMatrix = null;
            //SAPbouiCOM.DBDataSource oDbDS = null;
            SAPbouiCOM.Folder objFolder = null;
            XmlDocument xmlResult = null;
            SAPbouiCOM.BoFormMode objMode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            bool blFolderFound = false;
            XmlDocument objFormXML = null;
            XmlNode objNode = null;

            try
            {

                objForm = MainObject.Instance.B1Application.Forms.Item(strFormUID);
                objMode = objForm.Mode;

                objFormXML = new XmlDocument();
                objFormXML.LoadXml(objForm.GetAsXML());
                objNode = objFormXML.SelectSingleNode("/Application/forms/action/form/items/action/item[@uid='BYB_FLRP']");
                if(objNode != null)
                {
                    blFolderFound = true;
                }
                objForm.Freeze(true);
                if (blFolderFound)
                { 
                    objForm.Freeze(true);
                    objItem = objForm.Items.Item(Settings._Main.RelatedPartiesFolderId);
                }
                else
                {
                    objForm.Freeze(true);
                    string strFolderXML = RelatedPartiesRes.BYB_Folder_Terceros_Relacionados;
                    strFolderXML = strFolderXML.Replace("[--UniqueId--]", strFormUID);
                    MainObject.Instance.B1Application.LoadBatchActions(ref strFolderXML);
                    string strResult = MainObject.Instance.B1Application.GetLastBatchResults();
                    xmlResult = new XmlDocument();
                    xmlResult.LoadXml(strResult);
                    bool errors = xmlResult.SelectSingleNode("/result/errors").HasChildNodes != true ? false : true;
                    if (!errors)
                    {
                        objItem = objForm.Items.Item(Settings._Main.RelatedPartiesFolderId);
                        
                    }
                    else
                    {
                        objItem = null;
                    }
                }
                
                objForm.Freeze(false);
                
                if (objItem != null)
                {
                    objForm.Freeze(true);
                    #region Folder
                    
                    objItemBase = objForm.Items.Item(Settings._Main.lastFolderId);
                    if (objItemBase != null)
                    {
                        if (objItemBase.Visible)
                        {

                            intLeft = objItemBase.Left;
                            strUID = objItemBase.UniqueID;
                            objItem.Left = intLeft + 1;
                            objItem.FromPane = 0;
                            objItem.ToPane = 0;
                            objFolder = objItem.Specific;
                            objFolder.GroupWith(strUID);



                            objItem.Visible = true;
                        }
                    }
                    #endregion Folder;
                    
                    objMatrix = objForm.Items.Item("BYB_I47").Specific;
                    objMatrix.LoadFromDataSource();

                    objForm.Mode = objMode;
                    objForm.Freeze(false);
                }
                else
                {
                    objForm.Mode = objMode;
                }

                //}



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

            if (objForm != null)
            {
                objForm.Freeze(false);

            }
        }

        static public string gotRPInfo(string strCardCode)
        {
            string strResult = "";
            
            SAPbobsCOM.Recordset objRecordSet = null;
            
            string strSQL = "";
            try
            {
                objRecordSet = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if(CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName))
                {
                    strSQL = Settings._HANA.getCodeFromCardCode;
                }
                else
                {
                    strSQL = Settings._SQL.getCodeFromCardCode;
                }
                strSQL = strSQL.Replace("[--CardCode--]", strCardCode);
                objRecordSet.DoQuery(strSQL);
                if(objRecordSet.RecordCount > 0)
                {
                    strResult = objRecordSet.Fields.Item(0).Value;
                }
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                strResult = "";

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                strResult = "";
            }
            finally
            {
                if(objRecordSet != null)
                {
                    objRecordSet = null;
                }
            }
            return strResult;
        }

        static public void getRelatedpartyInfo(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            SAPbouiCOM.DBDataSource objDBDS = null;
            string strInternalCode = "";
            

            SAPbobsCOM.BusinessPartners objBP = null;
            string strCardCode = "";
            SAPbouiCOM.Conditions objCOnditions = null;
            SAPbouiCOM.Condition objCond = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;
            

            try
            {
                

                objBP = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                if(objBP.Browser.GetByKeys(BusinessObjectInfo.ObjectKey))
                {
                    strCardCode = objBP.CardCode;
                    strInternalCode = gotRPInfo(strCardCode).Trim();
                    objForm = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                    if (strInternalCode.Length > 0)
                    {
                        objForm.Freeze(true);
                        objCOnditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                        objCond = objCOnditions.Add();
                        objCond.Alias = "Code";
                        objCond.CondVal = strInternalCode;
                        objCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        objDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1RPA100");
                        objDBDS.Query(objCOnditions);

                        objCOnditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                        objCond = objCOnditions.Add();
                        objCond.Alias = "Code";
                        objCond.CondVal = strInternalCode;
                        objCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        objDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1RPA101");
                        objDBDS.Query(objCOnditions);
                        objMatrix = objForm.Items.Item("BYB_I47").Specific;
                        objMatrix.LoadFromDataSource();
                    }
                    else
                    {
                        cleanEditTexts(objForm.UniqueID);
                    }

                }
                else
                {
                    _Logger.Error("Could not get BP information from ObjectKey");
                    MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + "Could not get BP information from ObjectKey", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

            if (objForm != null)
            {
                objForm.Freeze(false);

            }
        }
        
        static public void addRelatedPartyInfo(SAPbouiCOM.BusinessObjectInfo objBusinessObjectInfo)
        {
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objGeneralService = null;
            SAPbobsCOM.GeneralData objGeneralData = null;
            SAPbobsCOM.GeneralData objRelationLine = null;
            SAPbobsCOM.GeneralDataCollection objRelationData = null;
            SAPbouiCOM.Form objForm = null;
            SAPbobsCOM.BusinessPartners objBP = null;
            SAPbouiCOM.DBDataSource objMainSource = null;
            SAPbouiCOM.DBDataSource objRelation = null;

            string strInternalCode = "";
            string strLegalName = "";
            string strID = "";
            int intOffSet = 0;
            


            try
            {
                objBP = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                if (objBP.Browser.GetByKeys(objBusinessObjectInfo.ObjectKey))
                {
                    objForm = MainObject.Instance.B1Application.Forms.Item(objBusinessObjectInfo.FormUID);
                    objMainSource = objForm.DataSources.DBDataSources.Item("@BYB_T1RPA100");
                    objRelation = objForm.DataSources.DBDataSources.Item("@BYB_T1RPA101");

                    if (objMainSource.Size > 0)
                    {


                        strInternalCode = objMainSource.GetValue("Code", 0).Trim() == "" ? objBP.CardCode : objMainSource.GetValue("Code", 0).Trim();
                        strLegalName = objMainSource.GetValue("U_LEGALNAME", 0).Trim() == "" ? objBP.CardName : objMainSource.GetValue("U_LEGALNAME", 0).Trim();
                        strID = objMainSource.GetValue("U_IDNUM", 0).Trim() == "" ? objBP.FederalTaxID : objMainSource.GetValue("U_IDNUM", 0).Trim();


                        objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                        objGeneralService = objCompanyService.GetGeneralService("BYB_T1RPA100");
                        objGeneralData = objGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);



                        objGeneralData.SetProperty("Code", strInternalCode);
                        objGeneralData.SetProperty("Name", strInternalCode);
                        objGeneralData.SetProperty("U_TIPODOC", objMainSource.GetValue("U_TIPODOC", 0).Trim());
                        objGeneralData.SetProperty("U_CARDCODE", objBP.CardCode);
                        objGeneralData.SetProperty("U_PHONE1", objMainSource.GetValue("U_PHONE1", 0).Trim());
                        objGeneralData.SetProperty("U_PHONE2", objMainSource.GetValue("U_PHONE2", 0).Trim());
                        objGeneralData.SetProperty("U_FIRSTNAME1", objMainSource.GetValue("U_FIRSTNAME1", 0).Trim());
                        objGeneralData.SetProperty("U_FIRSTNAME2", objMainSource.GetValue("U_FIRSTNAME2", 0).Trim());
                        objGeneralData.SetProperty("U_LASTNAME1", objMainSource.GetValue("U_LASTNAME1", 0).Trim());
                        objGeneralData.SetProperty("U_LASTNAME2", objMainSource.GetValue("U_LASTNAME2", 0).Trim());
                        objGeneralData.SetProperty("U_ONAME", objMainSource.GetValue("U_ONAME", 0).Trim());
                        objGeneralData.SetProperty("U_ADDRESS1", objMainSource.GetValue("U_ADDRESS1", 0).Trim());
                        objGeneralData.SetProperty("U_ADDRESS2", objMainSource.GetValue("U_ADDRESS2", 0).Trim());
                        objGeneralData.SetProperty("U_MUNICODE", objMainSource.GetValue("U_MUNICODE", 0).Trim());
                        objGeneralData.SetProperty("U_COUCODE", objMainSource.GetValue("U_COUCODE", 0).Trim());
                        objGeneralData.SetProperty("U_DEPTCODE", objMainSource.GetValue("U_DEPTCODE", 0).Trim());
                        objGeneralData.SetProperty("U_NATURE", objMainSource.GetValue("U_NATURE", 0).Trim());
                        objGeneralData.SetProperty("U_LEGALNAME", strLegalName);
                        objGeneralData.SetProperty("U_IDNUM", strID);
                        objGeneralData.SetProperty("U_EMAIL", objMainSource.GetValue("U_EMAIL", 0).Trim());
                        objGeneralData.SetProperty("U_DIGVER", objMainSource.GetValue("U_DIGVER", 0).Trim());
                        objGeneralData.SetProperty("U_REGIMEN", objMainSource.GetValue("U_REGIMEN", 0).Trim());

                        objRelationData = objGeneralData.Child("BYB_T1RPA101");

                        intOffSet = objRelation.Offset;
                        for (int i = 0; i < objRelation.Size; i++)
                        {
                            objRelationLine = objRelationData.Add();
                            objRelationLine.SetProperty("U_RELCOD", objRelation.GetValue("U_RELCOD", i).Trim());

                        }

                        objGeneralService.Add(objGeneralData);
                        
                    }
                    else
                    {
                        strInternalCode =  objBP.CardCode ;
                        strLegalName = objBP.CardName;
                        strID = objBP.FederalTaxID;


                        objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                        objGeneralService = objCompanyService.GetGeneralService("BYB_T1RPA100");
                        objGeneralData = objGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);



                        objGeneralData.SetProperty("Code", strInternalCode);
                        objGeneralData.SetProperty("Name", strInternalCode);
                        objGeneralData.SetProperty("U_CARDCODE", objBP.CardCode);
                        objGeneralData.SetProperty("U_LEGALNAME", strLegalName);
                        objGeneralData.SetProperty("U_IDNUM", strID);
                        

                        objRelationData = objGeneralData.Child("BYB_T1RPA101");

                        intOffSet = objRelation.Offset;
                        for (int i = 0; i < objRelation.Size; i++)
                        {
                            objRelationLine = objRelationData.Add();
                            objRelationLine.SetProperty("U_RELCOD", objRelation.GetValue("U_RELCOD", i).Trim());

                        }

                        objGeneralService.Add(objGeneralData);
                        
                    }
                }
                else
                {
                    _Logger.Error("Could not retrieve BP Information from key");
                    MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + "Could not retrieve BP Information from key", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }









            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

            
        }

        static public void updateRelatedPartyInfo(SAPbouiCOM.BusinessObjectInfo objBusinessObjectInfo)
        {
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objGeneralService = null;
            SAPbobsCOM.GeneralData objGeneralData = null;
            SAPbobsCOM.GeneralData objRelationLine = null;
            SAPbobsCOM.GeneralDataCollection objRelationData = null;
            SAPbobsCOM.GeneralDataParams objGeneralDataParams = null;

            SAPbouiCOM.Matrix objMatrix = null;


            SAPbouiCOM.Form objForm = null;
            SAPbobsCOM.BusinessPartners objBP = null;
            SAPbouiCOM.DBDataSource objMainSource = null;
            SAPbouiCOM.DBDataSource objRelation = null;

            string strInternalCode = "";
            string strLegalName = "";
            string strID = "";
            int intOffSet = 0;
            
            try
            {
                objBP = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                if (objBP.Browser.GetByKeys(objBusinessObjectInfo.ObjectKey))
                {
                    objForm = MainObject.Instance.B1Application.Forms.Item(objBusinessObjectInfo.FormUID);
                    objMatrix = objForm.Items.Item("BYB_I47").Specific;
                    objMatrix.FlushToDataSource();
                    objMainSource = objForm.DataSources.DBDataSources.Item("@BYB_T1RPA100");
                    objRelation = objForm.DataSources.DBDataSources.Item("@BYB_T1RPA101");

                    if (objMainSource.Size > 0)
                    {


                        strInternalCode = objMainSource.GetValue("Code", 0).Trim() == "" ? objBP.CardCode : objMainSource.GetValue("Code", 0).Trim();
                        strLegalName = objMainSource.GetValue("U_LEGALNAME", 0).Trim() == "" ? objBP.CardName : objMainSource.GetValue("U_LEGALNAME", 0).Trim();
                        strID = objMainSource.GetValue("U_IDNUM", 0).Trim() == "" ? objBP.FederalTaxID : objMainSource.GetValue("U_IDNUM", 0).Trim();


                        objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                        objGeneralService = objCompanyService.GetGeneralService("BYB_T1RPA100");
                        objGeneralDataParams = objGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        objGeneralDataParams.SetProperty("Code", strInternalCode);
                        try
                        {
                            objGeneralData = objGeneralService.GetByParams(objGeneralDataParams);



                            objGeneralData.SetProperty("Code", strInternalCode);
                            objGeneralData.SetProperty("Name", strInternalCode);
                            objGeneralData.SetProperty("U_TIPODOC", objMainSource.GetValue("U_TIPODOC", 0).Trim());
                            objGeneralData.SetProperty("U_CARDCODE", objBP.CardCode);
                            objGeneralData.SetProperty("U_PHONE1", objMainSource.GetValue("U_PHONE1", 0).Trim());
                            objGeneralData.SetProperty("U_PHONE2", objMainSource.GetValue("U_PHONE2", 0).Trim());
                            objGeneralData.SetProperty("U_FIRSTNAME1", objMainSource.GetValue("U_FIRSTNAME1", 0).Trim());
                            objGeneralData.SetProperty("U_FIRSTNAME2", objMainSource.GetValue("U_FIRSTNAME2", 0).Trim());
                            objGeneralData.SetProperty("U_LASTNAME1", objMainSource.GetValue("U_LASTNAME1", 0).Trim());
                            objGeneralData.SetProperty("U_LASTNAME2", objMainSource.GetValue("U_LASTNAME2", 0).Trim());
                            objGeneralData.SetProperty("U_ONAME", objMainSource.GetValue("U_ONAME", 0).Trim());
                            objGeneralData.SetProperty("U_ADDRESS1", objMainSource.GetValue("U_ADDRESS1", 0).Trim());
                            objGeneralData.SetProperty("U_ADDRESS2", objMainSource.GetValue("U_ADDRESS2", 0).Trim());
                            objGeneralData.SetProperty("U_MUNICODE", objMainSource.GetValue("U_MUNICODE", 0).Trim());
                            objGeneralData.SetProperty("U_COUCODE", objMainSource.GetValue("U_COUCODE", 0).Trim());
                            objGeneralData.SetProperty("U_DEPTCODE", objMainSource.GetValue("U_DEPTCODE", 0).Trim());
                            objGeneralData.SetProperty("U_NATURE", objMainSource.GetValue("U_NATURE", 0).Trim());
                            objGeneralData.SetProperty("U_LEGALNAME", strLegalName);
                            objGeneralData.SetProperty("U_IDNUM", strID);
                            objGeneralData.SetProperty("U_EMAIL", objMainSource.GetValue("U_EMAIL", 0).Trim());
                            objGeneralData.SetProperty("U_DIGVER", objMainSource.GetValue("U_DIGVER", 0).Trim());
                            objGeneralData.SetProperty("U_REGIMEN", objMainSource.GetValue("U_REGIMEN", 0).Trim());

                            objRelationData = objGeneralData.Child("BYB_T1RPA101");
                            for (int i = 0; i < objRelationData.Count; i++)
                            {
                                objRelationData.Remove(i);
                            }
                            for (int i = 0; i < objRelation.Size; i++)
                            {
                                string strRel = objRelation.GetValue("U_RELCOD", i).Trim();
                                if (strRel.Length > 0)
                                {
                                    objRelationLine = objRelationData.Add();
                                    objRelationLine.SetProperty("U_RELCOD", objRelation.GetValue("U_RELCOD", i).Trim());
                                }

                            }

                            objGeneralService.Update(objGeneralData);
                            getRelatedpartyInfo(objBusinessObjectInfo);
                        }
                        catch (COMException comEx)
                        {
                            if (comEx.ErrorCode == -2028)
                            {
                                addRelatedPartyInfo(objBusinessObjectInfo);
                                getRelatedpartyInfo(objBusinessObjectInfo);
                            }

                        }
                        
                        
                    }
                    else
                    {
                        strInternalCode = objBP.CardCode;
                        strLegalName = objBP.CardName;
                        strID = objBP.FederalTaxID;


                        objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                        objGeneralService = objCompanyService.GetGeneralService("BYB_T1RPA100");
                        objGeneralData = objGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);



                        objGeneralData.SetProperty("Code", strInternalCode);
                        objGeneralData.SetProperty("Name", strInternalCode);
                        objGeneralData.SetProperty("U_CARDCODE", objBP.CardCode);
                        objGeneralData.SetProperty("U_LEGALNAME", strLegalName);
                        objGeneralData.SetProperty("U_IDNUM", strID);


                        objRelationData = objGeneralData.Child("BYB_T1RPA101");

                        intOffSet = objRelation.Offset;
                        for (int i = 0; i < objRelation.Size; i++)
                        {
                            objRelationLine = objRelationData.Add();
                            objRelationLine.SetProperty("U_RELCOD", objRelation.GetValue("U_RELCOD", i).Trim());

                        }

                        objGeneralService.Add(objGeneralData);
                        
                    }
                }
                else
                {
                    _Logger.Error("Could not retrieve BP information from Key");
                    MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + "Could not retrieve BP information from Key", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }









            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB:" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }


        }

        static public void cleanEditTexts(string FormUID)
        {
            SAPbouiCOM.EditText objEdit = null;
            SAPbouiCOM.Form objForm = null;
            //SAPbouiCOM.DBDataSource objDS = null;
            SAPbouiCOM.Matrix objMatrix = null;
            string[] strEditTexts;
            int intActualRowCount;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
                objForm.Freeze(true);
                strEditTexts = Settings._Main.BPFormBYBEditTextItems.Split(',');

                foreach (string strField in strEditTexts)
                {
                    objEdit = objForm.Items.Item(strField).Specific;
                    objEdit.String = "";
                }
                objMatrix = objForm.Items.Item(Settings._Main.BPFormMatrixId).Specific;
                intActualRowCount = objMatrix.RowCount;
                for (int i=1; i <= intActualRowCount; i++)
                {
                    objMatrix.DeleteRow(1);
                }
                objMatrix.FlushToDataSource();
                objForm.Freeze(false);
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            if(objForm != null)
            {
                objForm.Freeze(false);
            }
        }

        static public void deleteRelatedPartyInfo(SAPbouiCOM.BusinessObjectInfo objBusinessObjectInfo)
        {
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objGeneralService = null;
            SAPbobsCOM.GeneralDataParams objGeneralDataParams = null;
            SAPbouiCOM.Form objForm = null;
            
            SAPbouiCOM.DBDataSource objMainSource = null;
            string strRPCode = "";

            //SAPbouiCOM.DBDataSource oDbDS = null;
            
            
            
            SAPbouiCOM.Matrix objMatrix = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(objBusinessObjectInfo.FormUID);
                objMainSource = objForm.DataSources.DBDataSources.Item("@BYB_T1RPA100");
                if (objMainSource.Size > 0)
                {
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objGeneralService = objCompanyService.GetGeneralService("BYB_T1RPA100");
                    
                        strRPCode = objMainSource.GetValue("Code", 0).Trim();
                        objGeneralDataParams = objGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        objGeneralDataParams.SetProperty("Code", strRPCode);
                    objGeneralService.Delete(objGeneralDataParams);
                    objForm.Freeze(true);
                    cleanEditTexts(objBusinessObjectInfo.FormUID);
                    objForm.Freeze(false);

                       



                //



                //objCOnditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                //objCond = objCOnditions.Add();
                //objCond.Alias = "Code";
                //objCond.CondVal = strRPCode;
                //objCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;

                //objMainSource.Query(objCOnditions);

                //objDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1RPA101");

                //for (int i = 0; i < objDBDS.Size; i++)
                //{
                //    for (int j = 0; j < objDBDS.Fields.Count; j++)
                //    {
                //        if (objDBDS.Fields.Item(j).Name.IndexOf("U_") == 0)
                //        {
                //            objDBDS.SetValue(j, i, "");
                //        }
                //    }
                //}

                //objCOnditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                //objCond = objCOnditions.Add();
                //objCond.Alias = "Code";
                //objCond.CondVal = strRPCode;
                //objCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //objDBDS.Query(objCOnditions);


                objMatrix = objForm.Items.Item("BYB_I47").Specific;
                    objMatrix.LoadFromDataSource();


                }
                    
                
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

            if(objForm != null)
            {
                objForm.Freeze(false);
            }


        }

        static public void addLineToDS()
        {
            SAPbouiCOM.Form objForm = null;
            //SAPbouiCOM.DBDataSource oDbDS = null;


            try
            {
                objForm = MainObject.Instance.B1Application.Forms.ActiveForm;
                if(objForm != null && objForm.TypeEx == Settings._Main.BPFormTypeEx)
                {
                    objForm.Freeze(true);
                    Instance.BYBRelatedPartiesFolderAdd(objForm.UniqueID);
                    
                    cleanEditTexts(objForm.UniqueID);
                    objForm.Freeze(false);
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

       static public void addInsertRowRelationMenu(SAPbouiCOM.Form objForm, SAPbouiCOM.ContextMenuInfo eventInfo)
        {
            SAPbouiCOM.MenuCreationParams objParams = null;

            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objParams.String = "Agregar línea";
                objParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                objParams.UniqueID = "BYB_MRPAR";
                objParams.Position = -1;
                objForm.Menu.AddEx(objParams);
                EventInfoClass objEvent = new EventInfoClass();
                objEvent.ColUID = eventInfo.ColUID;
                objEvent.FormUID = eventInfo.FormUID;
                objEvent.ItemUID = eventInfo.ItemUID;
                objEvent.Row = eventInfo.Row;
                CacheManager.CacheManager.Instance.addToCache(Settings._Main.lastRightClickEventInfo, objEvent, CacheManager.CacheManager.objCachePriority.Default);
                

                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void removeInsertRowRelationMenu()
        {
            

            try
            {
                if (MainObject.Instance.B1Application.Menus.Exists("BYB_MRPAR"))
                {
                    MainObject.Instance.B1Application.Menus.RemoveEx("BYB_MRPAR");
                }
                CacheManager.CacheManager.Instance.removeFromCache(Settings._Main.lastRightClickEventInfo);


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void addDeleteRowRelationMenu(SAPbouiCOM.Form objForm, SAPbouiCOM.ContextMenuInfo eventInfo)
        {
            SAPbouiCOM.MenuCreationParams objParams = null;

            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objParams.String = "Eliminar línea";
                objParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                objParams.UniqueID = "BYB_MRPDR";
                objParams.Position = -1;
                objForm.Menu.AddEx(objParams);
                EventInfoClass objEvent = new EventInfoClass();
                objEvent.ColUID = eventInfo.ColUID;
                objEvent.FormUID = eventInfo.FormUID;
                objEvent.ItemUID = eventInfo.ItemUID;
                objEvent.Row = eventInfo.Row;
                CacheManager.CacheManager.Instance.addToCache(Settings._Main.lastRightClickEventInfo, objEvent, CacheManager.CacheManager.objCachePriority.Default);


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void removeDeleteRowRelationMenu()
        {


            try
            {
                if (MainObject.Instance.B1Application.Menus.Exists("BYB_MRPDR"))
                {
                    MainObject.Instance.B1Application.Menus.RemoveEx("BYB_MRPDR");
                }
                CacheManager.CacheManager.Instance.removeFromCache(Settings._Main.lastRightClickEventInfo);


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void relatedPartiedMatrixOperation(EventInfoClass eventInfo, string Action)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;
            int intTotalLines = -1;
            //SAPbouiCOM.DBDataSource objDS = null;
            //int intSize = -1;
            


            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                objMatrix = objForm.Items.Item(Settings._Main.BPFormMatrixId).Specific;
                intTotalLines = objMatrix.RowCount;
                int intRow = eventInfo.Row;
                switch (Action)
                {
                    case "Add":
                        objMatrix.AddRow(1, intRow);
                        
                        objMatrix.SetCellWithoutValidation(intRow + 1, "Col_0", "");
                        objMatrix.FlushToDataSource();

                        objMatrix.SetCellFocus(intRow + 1, 1);
                        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                }

                        break;
                    case "Delete":
                        objMatrix.DeleteRow(intRow);
                        objMatrix.FlushToDataSource();
                        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                        break;
                        
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        #endregion

        #region Add missing Related Parties

        static public void createMissingRelatedParties(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.DataTable objDtRes = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Grid objGrid = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objGeneralService = null;
            SAPbobsCOM.GeneralData objGeneralData = null;
            SAPbobsCOM.GeneralDataParams objResult = null;
            SAPbouiCOM.Item objItem = null;
            string strInternalCode = "";
            string strLegalName = "";
            string strID = "";
            SAPbouiCOM.GridColumn objGridColumn;
            SAPbouiCOM.EditTextColumn oEditTExt;



            SAPbobsCOM.BusinessPartners objBP = null;
            
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objDT = objForm.DataSources.DataTables.Item("DT_TRA");
                objDtRes = objForm.DataSources.DataTables.Item("DT_RES");
                if (objDT.Rows.Count > 0)
                {
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    objGeneralService = objCompanyService.GetGeneralService("BYB_T1RPA100");
                    objGeneralData = objGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    objGrid = objForm.Items.Item("grTRA").Specific;
                    T1.B1.Base.UIOperations.Operations.startProgressBar("Procesando...", objDT.Rows.Count);
                    for (int i = 0; i < objDT.Rows.Count; i++)
                    {
                        string strCardCode = objDT.GetValue(0, i);
                        string strMessage = "";
                        objBP = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                        if (objBP.GetByKey(strCardCode))
                        {
                            strInternalCode = objBP.CardCode;
                            strLegalName = objBP.CardName;
                            strID = objBP.FederalTaxID;



                            objGeneralData.SetProperty("Code", strInternalCode);
                            objGeneralData.SetProperty("Name", strInternalCode);
                            objGeneralData.SetProperty("U_CARDCODE", objBP.CardCode);
                            objGeneralData.SetProperty("U_LEGALNAME", strLegalName);
                            objGeneralData.SetProperty("U_IDNUM", strID);
                            try
                            {
                                objResult = objGeneralService.Add(objGeneralData);
                                if(objResult != null)
                                {
                                    strMessage = "OK";
                                }
                            }
                            catch(Exception er)
                            {
                                strMessage = er.Message;
                            }
                            objDtRes.Rows.Add(1);
                            objDtRes.SetValue(0, objDtRes.Rows.Count - 1, strInternalCode);
                            objDtRes.SetValue(1, objDtRes.Rows.Count - 1, strLegalName);
                            objDtRes.SetValue(2, objDtRes.Rows.Count - 1, strID);
                            objDtRes.SetValue(3, objDtRes.Rows.Count - 1, strMessage);
                            T1.B1.Base.UIOperations.Operations.setProgressBarMessage(strInternalCode + " procesado.", i + 1);


                        }
                    }
                    objForm.Freeze(true);
                    objGrid.DataTable = objDtRes;


                    objGrid = objForm.Items.Item("grTRA").Specific;

                    //objGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;


                    objGridColumn = objGrid.Columns.Item(0);
                    objGridColumn.Editable = false;
                    
                    objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                    oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                    oEditTExt.LinkedObjectType = "2";


                    objGridColumn = objGrid.Columns.Item(1);
                    objGridColumn.Editable = false;
                    

                    objGridColumn = objGrid.Columns.Item(2);
                    objGridColumn.Editable = false;
                    

                    objGridColumn = objGrid.Columns.Item(3);
                    objGridColumn.Editable = false;
                    

                    for(int i=0; i < objDtRes.Rows.Count; i++)
                    {
                        string strResult = objDtRes.GetValue("Result", i);
                        if(strResult == "OK")
                        {
                            objGrid.CommonSetting.SetCellBackColor(i + 1, 4, Color.Green.R | (Color.Green.G << 8) | (Color.Green.B << 16));

                        }
                        else
                        {
                            objGrid.CommonSetting.SetCellBackColor(i + 1, 4, Color.Red.R | (Color.Red.G << 8) | (Color.Red.B << 16));

                        }
                    }






                    objGrid.AutoResizeColumns();
                    



                    T1.B1.Base.UIOperations.Operations.stopProgressBar();

                    objItem = objForm.Items.Item("btnAdd");
                    objItem.Enabled = false;
                    objForm.Freeze(false);
                }
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
                if (objForm != null)
                {
                    objForm.Freeze(false);
                }
            }
        }

        static public void loadMissingRelatedPartiesForm()
        {
            string strSQL = "";
            SAPbobsCOM.Recordset objRS = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.FormCreationParams objParams = null;
            SAPbouiCOM.DataTable objDT = null;
            //SAPbouiCOM.Item objItem = null;
            SAPbouiCOM.Grid objGrid = null;
            SAPbouiCOM.GridColumn objGridColumn = null;
            SAPbouiCOM.EditTextColumn oEditTExt = null;


            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1,20);
                objParams.XmlData = RelatedPartiesRes.BYB_Terceros_Relacionados_Faltantes;
                objParams.FormType = "BYB_FTRA1";
                objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objDT = objForm.DataSources.DataTables.Item("DT_TRA");


                objRS = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName))
                {
                    strSQL = Settings._HANA.getMissingRP;
                }
                else
                {
                    strSQL = Settings._SQL.getMissingRP;
                }

                objDT.ExecuteQuery(strSQL);

                #region Format Grid
                objGrid = objForm.Items.Item("grTRA").Specific;

                //objGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;


                objGridColumn = objGrid.Columns.Item(0);
                objGridColumn.Editable = false;
                
                objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                oEditTExt.LinkedObjectType = "2";


                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;
                

                objGridColumn = objGrid.Columns.Item(2);
                objGridColumn.Editable = false;
                

                


                objGrid.AutoResizeColumns();


                #endregion


                objForm.Visible = true;

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

        }
        #endregion

        #region getBPThirdPartyRelation

        public static Dictionary<string,string> getBPThirdPartyRelation()
        {
            SAPbobsCOM.Recordset objRecordSet = null;
            Dictionary<string, string> objBPList = null;
            string strSQL = "";
            try
            {
                objBPList = new Dictionary<string, string>();
                if (CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName))
                {
                    strSQL = Settings._HANA.getBPCodeRelation;
                }
                else
                {
                    strSQL = Settings._SQL.getBPCodeRelation;
                }
                objRecordSet = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if(objRecordSet != null)
                {
                    objRecordSet.DoQuery(strSQL);
                    if(objRecordSet != null && objRecordSet.RecordCount > 0)
                    {
                        while(!objRecordSet.EoF)
                        {
                            string strTPCode = objRecordSet.Fields.Item(0).Value;
                            string strCardCode = objRecordSet.Fields.Item(1).Value;
                            if(!objBPList.ContainsKey(strCardCode))
                            {
                                objBPList.Add(strCardCode, strTPCode);
                            }


                            objRecordSet.MoveNext();
                        }
                    }
                    
                }

            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                objBPList = new Dictionary<string, string>();
            }
            return objBPList;
        }

        #endregion
    }
}
