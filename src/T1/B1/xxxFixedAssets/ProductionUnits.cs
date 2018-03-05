using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using T1.Classes;
using System.Xml;
using System.IO;

namespace T1.B1.FixedAssets
{
    
    public class ProductionUnits
    {
        static private ProductionUnits objProductionUnits = null;
        

        private ProductionUnits()
        {
        }
        static public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            if (objProductionUnits == null)
            {
                objProductionUnits = new ProductionUnits();
            }
            
            BubbleEvent = true;



        }

        static public void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            if (objProductionUnits == null)
            {
                objProductionUnits = new ProductionUnits();
            }

            BubbleEvent = true;

            SAPbouiCOM.Form objForm = null;

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                if (objForm.TypeEx == T1.B1.FixedAssets.InteractionId.Default.fixedAssetsMasterFormId)
                {
                    if (objForm.BusinessObject.Key.Length > 0)
                    {
                        if (eventInfo.BeforeAction)
                        {
                            if (checkNoAmortization(objForm.BusinessObject.Key))
                            {
                                BYBCache.Instance.addToCache(T1.B1.FixedAssets.CacheItemNames.Default.fixedAssetsCurrentKey, objForm.BusinessObject.Key, BYBCache.objCachePriority.NotRemovable);
                                BYBCache.Instance.addToCache(T1.B1.FixedAssets.CacheItemNames.Default.fixedAssetsRightClickForm, objForm.GetAsXML(), BYBCache.objCachePriority.Default);

                                addMainMenu();
                            }
                        }
                        else
                        {
                            removeMainMenu();
                        }
                    }

                    objForm = null;

                }
                else if (objForm.TypeEx == T1.B1.FixedAssets.InteractionId.Default.productionUnitsFormType)
                {
                    if (objForm.BusinessObject.Key.Length > 0)
                    {
                        if (eventInfo.BeforeAction && eventInfo.ItemUID=="Item_9")
                        {
                            addRowMenu(objForm);
                        }
                        else
                        {
                            removeRowMenu();
                        }
                    }
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.RightClickEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.RightClickEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form objForm = null;
            try
            {
                if(!pVal.BeforeAction)
                {
                    if(pVal.MenuUID == B1.FixedAssets.InteractionId.Default.fixedAssetsActivateMenuId)
                    {
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm(B1.FixedAssets.InteractionId.Default.fixedAssetsMasterFormId);
                        objFormCreationParams.FormType = InteractionId.Default.productionUnitsFormType;
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        objFormCreationParams.ObjectType = "SAPFXPU";
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);

                        XmlDocument objDoc = new XmlDocument();
                        string strAssetCode = BYBCache.Instance.getFromCache(B1.FixedAssets.CacheItemNames.Default.fixedAssetsCurrentKey);
                        objDoc.LoadXml(strAssetCode);
                        string strValue = objDoc.SelectSingleNode("/ItemParams/ItemCode").InnerText;


                        

                        SAPbouiCOM.ComboBox objCmbPeriod = objForm.Items.Item("cmbPeriod").Specific;
                        objCmbPeriod.ComboSelectAfter += objCmbPeriod_ComboSelectAfter;


                        //SAPbouiCOM.Matrix objMatrix = objForm.Items.Item("Item_9").Specific;
                        //objMatrix.LostFocusAfter += objMatrix_LostFocusAfter;

                        getExistingInformation(objForm, strValue);

                        objForm.Visible = true;

                        //TODO liberar cache de memoria del activo fijo y del formulario
                        BYBCache.Instance.removeFromCache(T1.B1.FixedAssets.CacheItemNames.Default.fixedAssetsCurrentKey);
                        BYBCache.Instance.removeFromCache(T1.B1.FixedAssets.CacheItemNames.Default.fixedAssetsRightClickForm);


                    }
                    else if (pVal.MenuUID == B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId)
                    {
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.ActiveForm;
                        SAPbouiCOM.DBDataSource objDS = objForm.DataSources.DBDataSources.Item("@SAPXPU1");
                        objDS.InsertRecord(objDS.Size);
                        SAPbouiCOM.Matrix objMatrix = objForm.Items.Item("Item_9").Specific;
                        objMatrix.LoadFromDataSource();

                        
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objMatrix_LostFocusAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            try
            {

                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);
                    SAPbouiCOM.EditText objEditText = objForm.Items.Item("txtUPT").Specific;
                    int intUP = Convert.ToInt32(objEditText.Value);

                    SAPbouiCOM.EditText objEditTextB = objForm.Items.Item("txtABalan").Specific;
                    double intBalance = BYBHelpers.Instance.getStandarNumericValue(objEditTextB.Value);
                    //int intBalance = Convert.ToInt32(objEditTextB.Value);


                    SAPbouiCOM.Matrix objMatrix = (SAPbouiCOM.Matrix)objForm.Items.Item(B1.FixedAssets.InteractionId.Default.fixedAssetsMasterFormMatrixId).Specific;
                    SAPbouiCOM.Column objColumn = objMatrix.Columns.Item("Col_2");


                    int intValue = Convert.ToInt32(objMatrix.GetCellSpecific("Col_1", pVal.Row).Value);
                    if (intValue > 0)
                    {
                        double percent = intUP / intValue;
                        double total = Convert.ToDouble(intBalance / percent);
                        objMatrix.SetCellWithoutValidation(pVal.Row, "Col_2", Convert.ToString(total));
                    }
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            
        }

        static void objCmbPeriod_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.FixedAssetItemsService objFAService = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.FixedAssetValuesParams objFAValuesParams = null;
            SAPbouiCOM.ComboBox objCombo = null;
            SAPbouiCOM.Form objForm = null;
            SAPbobsCOM.FixedAssetEndBalance objEndBalance = null;
            
            try
            {
                objCombo = (SAPbouiCOM.ComboBox)sboObject;
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                    objFAService = objCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.FixedAssetItemsService);
                    objFAValuesParams = objFAService.GetDataInterface(SAPbobsCOM.FixedAssetItemsServiceDataInterfaces.faisFixedAssetValuesParams);
                    objFAValuesParams.FiscalYear = objCombo.Selected.Value;
                    objFAValuesParams.ItemCode = objForm.Items.Item("txtACode").Specific.Value;
                    objFAValuesParams.DepreciationArea = "IFRS";
                    objEndBalance = objFAService.GetAssetEndBalance(objFAValuesParams);
                    objForm.Items.Item("txtABalan").Specific.Value = objEndBalance.AcquisitionCost;
                }
                else
                {

                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static private void addMainMenu()
        {
            string strMenuDescription = "";
            string strMenuId = "";
            int intPosition = -1;
            try
            {
                strMenuId = B1.FixedAssets.InteractionId.Default.fixedAssetsActivateMenuId;
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = T1.B1.FixedAssets.LocalizationStrings.Default.FixedAssetsActivateMenuString;
                    strMenuId = T1.B1.FixedAssets.InteractionId.Default.fixedAssetsActivateMenuId;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    intPosition = BYBB1MainObject.Instance.B1Application.Menus.Item(InteractionId.Default.fixedAssetsActivateMenuParentId).SubMenus.Count;
                    objMenuCreationParams.Position = intPosition + 1;

                    BYBB1MainObject.Instance.B1Application.Menus.Item(InteractionId.Default.fixedAssetsActivateMenuParentId).SubMenus.AddEx(objMenuCreationParams);
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        static private void removeMainMenu()
        {
            string strMenuId = "";

            try
            {
                strMenuId = T1.B1.FixedAssets.InteractionId.Default.fixedAssetsActivateMenuId;
                if(BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    BYBB1MainObject.Instance.B1Application.Menus.RemoveEx(strMenuId);

                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        static private string localForm(string strFormId)
        {
            string strResult = "";
            
            try
            {
                if (strFormId == B1.FixedAssets.InteractionId.Default.fixedAssetsMasterFormId)
                {
                    strResult = B1.FixedAssets.Resources.FixedAssets.FAF001;
                    strResult = strResult.Replace("/strTitle/", "Unidades de Producción");
                    strResult = strResult.Replace("/strlbl001/", "Activo Fijo:");
                    strResult = strResult.Replace("/strlbl002/", "Subperíodo:");
                    strResult = strResult.Replace("/strlbl007/", "Unidades de Producción:");
                    strResult = strResult.Replace("/strlbl003/", "Período");
                    strResult = strResult.Replace("/strlbl004/", "Unidades");
                    strResult = strResult.Replace("/strlbl005/", "Valor");
                    strResult = strResult.Replace("/strlbl006/", "Transacción");
                    string strValues = getPeriods();
                    strResult = strResult.Replace("/strlbl008/", strValues);
                    
                    

                }

               
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strResult;

        }

        static private void getExistingInformation(SAPbouiCOM.Form oForm, string strFixedAssetKey)
        {
            SAPbobsCOM.GeneralService objGeneralService = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralCollectionParams objGeneralList = null;
            SAPbobsCOM.GeneralDataParams objGeneralDataParams = null;
            
            bool blFound = false;
            
            try
            {
                objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                objGeneralService = objCompanyService.GetGeneralService("SAPFXPU");
                try
                {
                    objGeneralList = objGeneralService.GetList();
                    for (int i = 0; i < objGeneralList.Count; i++)
                    {
                        objGeneralDataParams = objGeneralList.Item(i);
                        string strCode = (string)objGeneralDataParams.GetProperty("Code");
                        if (strFixedAssetKey == strCode)
                        {
                            blFound = true;
                            break;
                        }

                    }
                }
                catch (COMException comEx)
                {
                    if(comEx.ErrorCode == -2028)
                    {
                        
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        

                        SAPbouiCOM.DBDataSource objDS = oForm.DataSources.DBDataSources.Item("@SAPFXPU");
                        SAPbouiCOM.EditText objEdit = oForm.Items.Item("txtACode").Specific;
                        objEdit.Value = strFixedAssetKey;
                        SAPbouiCOM.Matrix objMatrix = oForm.Items.Item("Item_9").Specific;
                        objMatrix.AutoResizeColumns();
                        oForm.Items.Item("txtUPT").Enabled = true;
                    }
                    else
                    {
                        throw comEx;
                    }
                }
                

                if(blFound)
                {
                    SAPbouiCOM.Conditions objConditions = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    SAPbouiCOM.Condition objC= objConditions.Add();
                    objC.BracketOpenNum = 1;
                    objC.Alias = "Code";
                    objC.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    objC.CondVal = strFixedAssetKey;
                    objC.BracketCloseNum = 1;

                    SAPbouiCOM.DBDataSource objDS = oForm.DataSources.DBDataSources.Item("@SAPFXPU");
                    objDS.Query(objConditions);

                    SAPbouiCOM.DBDataSource objDS2 = oForm.DataSources.DBDataSources.Item("@SAPXPU1");
                    objDS2.Query(objConditions);


                    SAPbouiCOM.Matrix objMatrix = oForm.Items.Item("Item_9").Specific;
                    objMatrix.LoadFromDataSource();
                    objMatrix.AutoResizeColumns();

                    
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                   

                    
                    
                    
                }
                else 
                { 
                    //TODO Add new object
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        static private void addRowMenu(SAPbouiCOM.Form objForm)
        {
            string strMenuDescription = "";
            string strMenuId = "";
            try
            {
                strMenuId = B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = T1.B1.FixedAssets.LocalizationStrings.Default.fixedAssetsAddRowMenuString;
                    strMenuId = T1.B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    objForm.Menu.AddEx(objMenuCreationParams);

                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        static private void removeRowMenu()
        {
            string strMenuId = "";

            try
            {
                strMenuId = T1.B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;
                if (BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    BYBB1MainObject.Instance.B1Application.Menus.RemoveEx(strMenuId);

                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }


        static private string getPeriods()
        {

            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.PeriodCategoryParamsCollection objPeriodsCollection = null;
            SAPbobsCOM.PeriodCategory objPeriodCat = null;
            string strResult = "";


            try
            {
                objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                objPeriodsCollection = objCompanyService.GetPeriods();

                if (objPeriodsCollection != null && objPeriodsCollection.Count > 0)
                {

                    strResult += "<action type=\"add\">";
                    foreach (SAPbobsCOM.PeriodCategoryParams objPeriodParams in objPeriodsCollection)
                    {
                        objPeriodCat = objCompanyService.GetPeriod(objPeriodParams);
                        strResult += string.Format("<ValidValue value=\"{0}\" description=\"{1}\"/>", objPeriodCat.FinancialYear, objPeriodCat.FinancialYear);
                    }
                    strResult += "</action>";
                }
                
                
                
                
                
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ProductionUnits.getPeriods", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ProductionUnits.getPeriods", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strResult;

        }

        static private bool checkNoAmortization(string ObjectKey)
        {
            bool blHasNoArea = false;
            XmlDocument objDocument = new XmlDocument();
            SAPbobsCOM.Items objItem = null;
            SAPbobsCOM.DepreciationTypesService objDepreciationTypeService = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.DepreciationTypeParams objParams = null;
            SAPbobsCOM.DepreciationType objDepType = null;
            try
            {
                objDocument.LoadXml(ObjectKey);
                string strItemCode = objDocument.SelectSingleNode("/ItemParams/ItemCode").InnerText;
                objItem = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

                if(objItem.GetByKey(strItemCode))
                {
                    SAPbobsCOM.ItemsDepreciationParameters oParams = objItem.DepreciationParameters;
                    for(int i=0; i < oParams.Count; i++)
                    {
                        oParams.SetCurrentLine(i);
                        string strDeprecType = oParams.DepreciationType;
                        objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                        objDepreciationTypeService = objCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.DepreciationTypesService);
                        objParams = objDepreciationTypeService.GetDataInterface(SAPbobsCOM.DepreciationTypesServiceDataInterfaces.dtsDepreciationTypeParams);
                        objParams.Code = strDeprecType;
                        objDepType = objDepreciationTypeService.Get(objParams);
                        if(objDepType.DepreciationMethod == SAPbobsCOM.DepreciationMethodEnum.dmNoDepreciation)
                        {
                            blHasNoArea = true;
                            break;
                        }
                        
                    }

                }
                


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.checkNoAmortization", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.checkNoAmortization", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return blHasNoArea;
        }

    }
}
