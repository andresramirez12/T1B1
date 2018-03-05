using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using T1.Classes;
using System.Xml;
using System.IO;
using System.Globalization;
using System.Resources;


namespace T1.B1.FSNotes
{
    public class FSNotes
    {
    
        static private FSNotes objFSNotes = null;
        static private ModuleRegistrationClass oModuleRegistration = new Classes.ModuleRegistrationClass();
        static private XmlDocument moduleInformation = null;


        private FSNotes()
        {
            oModuleRegistration.AddModuleInformation(InteractionId.Default.configModuleCode, InteractionId.Default.configModuleName, InteractionId.Default.configModuleVersion);
            
        }


        static public void addMainMenu()
        {
            if(objFSNotes == null)
            {
                objFSNotes = new FSNotes();
            }
            
            
            string strMenuDescription = "";
            string strMenuId = "";
            try
            {
                strMenuId = B1.FSNotes.InteractionId.Default.mnuNotesLoadId;
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = T1.B1.FSNotes.LocalizationStrings.Default.NotesMainMenuString;
                    

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    BYBB1MainObject.Instance.B1Application.Menus.Item(InteractionId.Default.notesLoadMenuParentId).SubMenus.AddEx(objMenuCreationParams);
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FSNotes.addMainMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FSNotes.addMainMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        static public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            if (objFSNotes == null)
            {
                objFSNotes = new FSNotes();
            }
            
            BubbleEvent = true;
            SAPbouiCOM.Form objForm = null;
            

            try
            {
                if (!pVal.BeforeAction)
                {
                    if (pVal.MenuUID == B1.FSNotes.InteractionId.Default.mnuNotesLoadId)
                    {
                        moduleInformation = BYBCache.Instance.getFromCache(T1.Properties.Settings.Default.VersionControlCacheName);
                        
                        
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm(B1.FSNotes.InteractionId.Default.fsNotesMainFormId);
                        objFormCreationParams.FormType = InteractionId.Default.fsNotesFormType;
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);

                        objForm.Visible = true;


                    }
                    
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FSNotes.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FSNotes.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void oUpdate_PressedAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {

            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.UserDataSource oCusomtTextDS = null;
            SAPbouiCOM.UserDataSource oCusomtNoteDS = null;

            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            string strCustomValue = "";
            string strNoteNumber = "";



            try
            {
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oCusomtTextDS = oForm.DataSources.UserDataSources.Item(InteractionId.Default.formCustomTextDS);
                strCustomValue = oCusomtTextDS.ValueEx;

                

                oCusomtNoteDS = oForm.DataSources.UserDataSources.Item(InteractionId.Default.formNoteNumberDS);
                strNoteNumber = oCusomtNoteDS.ValueEx;


                

                
                oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService(InteractionId.Default.fsNotesUDOName);

                
                        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oGeneralParams.SetProperty("Code", strNoteNumber);
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                        oGeneralData.SetProperty("U_CustomText", strCustomValue);
                        oGeneralService.Update(oGeneralData);

                    
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "FSNotes.addDefaultNotesValue", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "FSNotes.addDefaultNotesValue", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void oNoteCombo_ComboSelectAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            

            SAPbouiCOM.UserDataSource oNoteDS = null;
            SAPbouiCOM.UserDataSource oCustomDS = null;
            SAPbouiCOM.UserDataSource oNoteType = null;

            
            SAPbouiCOM.Form oForm = null;
            string strValueSelected = "";
            string strCurrentValue = "";
            string strNoteType = "";




            try
            {
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oNoteDS = oForm.DataSources.UserDataSources.Item(InteractionId.Default.formNoteNumberDS);
                oCustomDS = oForm.DataSources.UserDataSources.Item(InteractionId.Default.formCustomTextDS);
                oNoteType = oForm.DataSources.UserDataSources.Item(InteractionId.Default.formTypeDS);


                strNoteType = oNoteType.ValueEx;


                strValueSelected = oNoteDS.ValueEx;
                if (strNoteType == "2")
                { 
                strCurrentValue = getCustomInformation(strValueSelected);
            }
                else
                {
                    strCurrentValue = "";
                }

                oCustomDS.ValueEx = strCurrentValue;
                



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FSNotes.oNoteCombo_ComboSelectAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FSNotes.oNoteCombo_ComboSelectAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            
        }

        static public void oGetButton_PressedAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            
            try
            {
                BYBB1MainObject.Instance.B1Application.ActivateMenuItem(InteractionId.Default.formPreviewMenuId);

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FSNotes.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FSNotes.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }


        static public void LayoutKeyEvent(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.ComboBox oCombo = null;
            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item(InteractionId.Default.formTypeCmbNote).Specific;
                string strSelectedValue = oCombo.Selected.Value;
                eventInfo.LayoutKey = strSelectedValue;

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FSNotes.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FSNotes.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }


        static public void oTypeCombo_ComboSelectAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            XmlDocument oConfig = null;
            string strSelectedValue = "";
            SAPbouiCOM.ComboBox oCombo = null;
            SAPbobsCOM.ReportTypesService rptTypeService = null;

            SAPbobsCOM.ReportTypeParams newTypeParam = null;

            SAPbobsCOM.ReportType newType = null;

            
            try
            {
                oConfig = BYBCache.Instance.getFromCache(T1.Properties.Settings.Default.VersionControlCacheName);
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oCombo = oForm.Items.Item(B1.FSNotes.InteractionId.Default.formTypeCmbId).Specific;
                if (oCombo.Selected != null)
                {
                    strSelectedValue = oCombo.Selected.Value;
                    string strReportName = InteractionId.Default[InteractionId.Default.formTypeCmbSelecterdBase + strSelectedValue].ToString();
                    if (oConfig != null)
                    {
                        string strXpath = string.Format(InteractionId.Default.configReportNodePath, InteractionId.Default.configModuleCode, strReportName);

                        XmlNode oReportNode = oConfig.SelectSingleNode(strXpath);
                        if (oReportNode != null)
                        {
                            string strReportCode = oReportNode.Attributes[InteractionId.Default.configReportCodeAttribute].Value;
                            string strReportType = oReportNode.Attributes[InteractionId.Default.configReportTypeAttribute].Value;

                            rptTypeService = BYBB1MainObject.Instance.B1Company.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
                            newTypeParam = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportTypeParams);
                            newTypeParam.TypeCode = strReportType;
                            newType = rptTypeService.GetReportType(newTypeParam);
                            newType.DefaultReportLayout = strReportCode;
                            rptTypeService.UpdateReportType(newType);




                            oForm.ReportType = strReportCode;

                        }
                    }
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FSNotes.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FSNotes.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        static private string localForm(string strFormId)
        {
            
            string strResult = "";

            try
            {
                if (strFormId == B1.FSNotes.InteractionId.Default.fsNotesMainFormId)
                {
                    strResult = B1.FSNotes.Resources.FSNotes.FSN0001;
                    strResult = strResult.Replace("/strTitle/", "Notas a los estados financieros");
                    strResult = strResult.Replace("/strlbl001/", "Nota:");
                    strResult = strResult.Replace("/strlbl002/", "Consultar:");
                    strResult = strResult.Replace("/strlbl003/", "Consultar");
                    strResult = strResult.Replace("/strlbl004/", "Personalizar");
                    strResult = strResult.Replace("/vvExample/", "Ejemplo");
                    strResult = strResult.Replace("/vvCustom/", "Personalizado");
                    strResult = string.Format(strResult, addValidNotesValues());
                    
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FSNotes.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FSNotes.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strResult;

        }

        static private string addValidNotesValues()
        {
            
            
            string strResult = "";

            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralCollectionParams oGeneralCollectionParams = null;

            
            try
            {

                
                    oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService(InteractionId.Default.fsNotesUDOName);
                    
                    oGeneralCollectionParams = oGeneralService.GetList();
                    

                    for(int i = 0; i < oGeneralCollectionParams.Count; i ++)
                {
                    oGeneralParams = oGeneralCollectionParams.Item(i);
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                    strResult += string.Format(InteractionId.Default.fsValidValuesString, oGeneralData.GetProperty("Code"), oGeneralData.GetProperty("Name"));
                }

                    return strResult;



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FSNotes.addValidNotesValues", er, 1, System.Diagnostics.EventLogEntryType.Error);
                return strResult;
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FSNotes.addValidNotesValues", er, 1, System.Diagnostics.EventLogEntryType.Error);
                return strResult;
            }
        }

        static public bool createMetaData()
        {
            bool blResult = false;
            if (objFSNotes == null)
            {
                objFSNotes = new FSNotes();
            }

            try{
            
                addDefaultNotesValue();
                addNSReports();
                blResult = true;


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

        static private void addDefaultNotesValue()
        {
            
            ResourceSet oResources = null;
            
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            
            

            try
            {
                oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService(InteractionId.Default.fsNotesUDOName);
                
                oResources = B1.FSNotes.Resources.FSNotes.ResourceManager.GetResourceSet(CultureInfo.CurrentUICulture, true, true);
                foreach(DictionaryEntry entry in oResources)
                {
                    if(entry.Key.ToString().IndexOf("Nota") == 0)
                    {
                        string[] strNoteValue = entry.Key.ToString().Split('_');
                        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oGeneralParams.SetProperty("Code", strNoteValue[1]);
                        try
                        {
                            oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                            oGeneralData.SetProperty("U_BaseText", entry.Value);
                            oGeneralService.Update(oGeneralData);
                        }
                        catch(COMException comEx)
                        {
                            if(comEx.ErrorCode == -2028)
                            {
                                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                                oGeneralData.SetProperty("Code", strNoteValue[1]);
                                oGeneralData.SetProperty("Name", "Nota " + strNoteValue[1]);
                                oGeneralData.SetProperty("U_BaseText", entry.Value);
                                oGeneralService.Add(oGeneralData);

                            }
                        }
                        
                        

                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "FSNotes.addDefaultNotesValue", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "FSNotes.addDefaultNotesValue", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static private void addNSReports()
        {

            
            XmlDocument oDocument = null;
            string strReportTypeCode ="";
            
            try
            {
                oDocument = new XmlDocument();
                oDocument.LoadXml(Resources.FSNotes.CRDefinition);
                
                XmlNodeList oCR = oDocument.SelectNodes(T1.Properties.Settings.Default.crReportTypePath);
                if(oCR != null && oCR.Count > 0)
                {
                    
                    foreach(XmlNode xnItem in oCR)
                    {
                        string strTypeName = xnItem.Attributes[InteractionId.Default.crReportTypeNameAttribute].Value;
                        string strFormId = xnItem.Attributes[InteractionId.Default.crReportFormIdAttribute].Value;
                        strReportTypeCode = oModuleRegistration.AddReportType(InteractionId.Default.configModuleCode, strTypeName, T1.Properties.Settings.Default.AddOnName, strFormId);
                        XmlNodeList oReports = xnItem.SelectNodes(T1.Properties.Settings.Default.crReportPath);
                        if(oReports != null && oReports.Count > 0)
                        {
                            foreach (XmlNode xnReport in oReports)
                            {
                                string strName = xnReport.Attributes[InteractionId.Default.crReportNameAttribute].Value;
                                string strPath = xnReport.Attributes[InteractionId.Default.crReportPathAttribute].Value;
                                string strReportCode = oModuleRegistration.AddReport(InteractionId.Default.configModuleCode, T1.Properties.Settings.Default.ReportAuthor, strName, strReportTypeCode, strPath);

                            }
                        }
                    }

                    oModuleRegistration.updateModuleRegistration(InteractionId.Default.configModuleCode);

                }
                

            }
             catch (COMException comEx)
             {
                 Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                 BYBExceptionHandling.reportException(er.Message, "FSNotes.addNSReports", er, 1, System.Diagnostics.EventLogEntryType.Error);
             }
             catch (Exception er)
             {
                 BYBExceptionHandling.reportException(er.Message, "FSNotes.addNSReports", er, 1, System.Diagnostics.EventLogEntryType.Error);
             }

        }

        static private string getCustomInformation(string strKey)
        {
            string strValue = "";

            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;


            try
            {
                oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService(InteractionId.Default.fsNotesUDOName);
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("Code", strKey);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                if(oGeneralData != null)
                {
                    strValue = oGeneralData.GetProperty("U_CustomText");
                }
                

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FSNotes.getCustomInformation", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FSNotes.getCustomInformation", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strValue;
        }

        /*
        static public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            if (objProductionUnits == null)
            {
                objProductionUnits = new FSNotes();
            }
            
            BubbleEvent = true;



        }

        static public void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            if (objProductionUnits == null)
            {
                objProductionUnits = new FSNotes();
            }

            BubbleEvent = true;

            SAPbouiCOM.Form objForm = null;

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                if (objForm.TypeEx == SAPIFRS.B1.FixedAssets.InteractionId.Default.fixedAssetsMasterFormId)
                {
                    if (objForm.BusinessObject.Key.Length > 0)
                    {
                        if (eventInfo.BeforeAction)
                        {
                            if (checkNoAmortization(objForm.BusinessObject.Key))
                            {
                                BYBCache.Instance.addToCache(SAPIFRS.B1.FixedAssets.CacheItemNames.Default.fixedAssetsCurrentKey, objForm.BusinessObject.Key, BYBCache.objCachePriority.NotRemovable);
                                BYBCache.Instance.addToCache(SAPIFRS.B1.FixedAssets.CacheItemNames.Default.fixedAssetsRightClickForm, objForm.GetAsXML(), BYBCache.objCachePriority.Default);

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
                else if (objForm.TypeEx == SAPIFRS.B1.FixedAssets.InteractionId.Default.productionUnitsFormType)
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

        

        static void objMatrix_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form objForm = null;
            try
            {
                if (pVal.ColUID == "Col_1")
                {
                    objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.EditText objEditText = objForm.Items.Item("txtUPT").Specific;
                    int intUP = Convert.ToInt32(objEditText.Value);

                    SAPbouiCOM.EditText objEditTextB = objForm.Items.Item("txtABalan").Specific;
                    int intBalance = BYBHelpers.Instance.getIntValue(objEditTextB.Value);
                    //int intBalance = Convert.ToInt32(objEditTextB.Value);


                    SAPbouiCOM.Matrix objMatrix = (SAPbouiCOM.Matrix)sboObject;
                    SAPbouiCOM.Column objColumn = objMatrix.Columns.Item("Col_2");


                    int intValue = Convert.ToInt32(objMatrix.GetCellSpecific("Col_1", pVal.Row).Value);
                    if (intValue > 0)
                    {
                        double percent = intUP / intValue;
                        int total = Convert.ToInt32(intBalance / percent);
                        objMatrix.SetCellWithoutValidation(pVal.Row, "Col_2", Convert.ToString(total));
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

       

        static private void removeMainMenu()
        {
            string strMenuId = "";

            try
            {
                strMenuId = SAPIFRS.B1.FixedAssets.InteractionId.Default.fixedAssetsActivateMenuId;
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
                    strMenuDescription = SAPIFRS.B1.FixedAssets.LocalizationStrings.Default.fixedAssetsAddRowMenuString;
                    strMenuId = SAPIFRS.B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;

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
                strMenuId = SAPIFRS.B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;
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

        static public bool createMetaData()
        {
            bool blResult = false;
            BYBB1MainObject.Instance.B1Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
            BYBB1MainObject.Instance.B1Company.XMLAsString = false;

            try
            {
                string strMetaDataXML = B1.FixedAssets.Resources.FixedAssets.MetadataCreation;
                string strXMLHeader = SAPIFRS.Properties.Resources.XMLHeader;
                XmlDocument objDocument = new XmlDocument();
                objDocument.LoadXml(strMetaDataXML);
                XmlNodeList objTables = objDocument.SelectNodes(B1.FixedAssets.InteractionId.Default.mdTablePath);
                XmlNodeList objUserFields = objDocument.SelectNodes(B1.FixedAssets.InteractionId.Default.mdUserFieldsPath);
                XmlNodeList objUDO = objDocument.SelectNodes(B1.FixedAssets.InteractionId.Default.mdUDOPath);

                if(objTables != null && objTables.Count > 0)
                {
                    foreach(XmlNode xn in objTables)
                    {
                        string strXML = strXMLHeader + xn.InnerXml;
                        using(StreamWriter sr = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)+"\\temp.xml",false))
                        {
                            sr.Write(strXML);
                        }
                        SAPbobsCOM.UserTablesMD objUMD = BYBB1MainObject.Instance.B1Company.GetBusinessObjectFromXML(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\temp.xml", 0);
                        int iResult = objUMD.Add();
                        if(iResult == 0 || iResult == -2035)
                        {
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objUMD);
                            objUMD = null;
                        }
                        else{
                            Exception er = new Exception(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());
                            BYBExceptionHandling.reportException(er.Message, "ProductionUnits.createMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
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
                            BYBExceptionHandling.reportException(er.Message, "ProductionUnits.createMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
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
                            BYBExceptionHandling.reportException(er.Message, "ProductionUnits.createMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
                        }


                    }

                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ProductionUnits.createMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ProductionUnits.createMetaData", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

            



            SAPbobsCOM.UserObjectsMD  objUserObject = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            objUserObject.GetByKey("SAPFXPU");
            string strTemp = objUserObject.GetAsXML();

            return blResult;
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
                    for(int i=0; i <= oParams.Count; i++)
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
        */
    }
}

