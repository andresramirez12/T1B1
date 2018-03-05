using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;
using System.Reflection;

namespace T1.B1.Expenses
{
    public class Expenses
    {
        static private Expenses objExpenses = null;


        #region EventManagement

        static public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form objForm = null;


            try
            {
                if (objExpenses == null)
                    objExpenses = new Expenses();

                if (!pVal.BeforeAction)
                {
                    if (pVal.MenuUID == B1.Expenses.InteractionId.Default.T1ExpensesConceptsMenuId)
                    {
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm("BYBEXPF001");
                        objFormCreationParams.FormType = "BYBEXPF001";
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);
                        objForm.Visible = true;
                    }
                    else if (pVal.MenuUID == B1.Expenses.InteractionId.Default.T1ExpensesExpenseTypeMenuId)
                    {
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm("BYBEXPF002");
                        objFormCreationParams.FormType = "BYBEXPF002";
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);
                        objForm.Visible = true;
                    }

                    else if (pVal.MenuUID == B1.Expenses.InteractionId.Default.T1ExpensesRequestMenuId)
                    {
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm("BYBEXPF003");
                        objFormCreationParams.FormType = "BYBEXPF003";
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);
                        objForm.Visible = true;
                    }
                    else if (pVal.MenuUID == B1.Expenses.InteractionId.Default.T1ExpensesReportMenuId)
                    {
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm("BYBEXPF004");
                        objFormCreationParams.FormType = "BYBEXPF004";
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);

                        #region addLandedCostsCombo

                        SAPbobsCOM.LandedCostsService oLandedCosts = null;
                        SAPbobsCOM.CompanyService oCompanyService = null;
                        SAPbobsCOM.LandedCostsParams oLandedCostsParams = null;
                        SAPbouiCOM.ComboBox oCombo = objForm.Items.Item("22").Specific;


                        oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                        oLandedCosts = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.LandedCostsService);
                        oLandedCostsParams = oLandedCosts.GetLandedCostList();

                        for (int i = 0; i < oLandedCostsParams.Count; i++)
                        {
                            SAPbobsCOM.LandedCost oCost = oLandedCosts.GetLandedCost(oLandedCostsParams.Item(i));
                            oCombo.ValidValues.Add(oCost.LandedCostNumber.ToString(), oCost.LandedCostNumber.ToString());
                        }






                        #endregion addLandedCostsCombo

                        objForm.Visible = true;
                    }
                    else if (pVal.MenuUID == "BYBEXPRC1")
                    {
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.ActiveForm;
                        SAPbouiCOM.Matrix objMatrix = objForm.Items.Item("0_U_G").Specific;

                        objMatrix.FlushToDataSource();
                        SAPbouiCOM.DBDataSource objDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP101");

                        objDS.InsertRecord(objDS.Size);
                        objMatrix.LoadFromDataSource();

                    }
                    else if (pVal.MenuUID == "BYBEXPRC2")
                    {
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.ActiveForm;
                        SAPbouiCOM.Matrix objMatrix = objForm.Items.Item("0_U_G").Specific;
                        objMatrix.FlushToDataSource();
                        SAPbouiCOM.DBDataSource objDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP101");
                        int intRow = BYBCache.Instance.getFromCache("RC" + objForm.UniqueID);

                        objDS.RemoveRecord(intRow - 1);

                        objMatrix.LoadFromDataSource();
                    }
                    else if (pVal.MenuUID == "BYBEXPRC3")
                    {
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.ActiveForm;
                        SAPbouiCOM.Matrix objMatrix = objForm.Items.Item("0_U_G").Specific;

                        objMatrix.FlushToDataSource();
                        SAPbouiCOM.DBDataSource objDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP401");

                        objDS.InsertRecord(objDS.Size);
                        objMatrix.LoadFromDataSource();

                    }
                    else if (pVal.MenuUID == "BYBEXPRC4")
                    {
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.ActiveForm;
                        SAPbouiCOM.Matrix objMatrix = objForm.Items.Item("0_U_G").Specific;
                        objMatrix.FlushToDataSource();
                        SAPbouiCOM.DBDataSource objDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP401");
                        int intRow = BYBCache.Instance.getFromCache("RC" + objForm.UniqueID);

                        objDS.RemoveRecord(intRow - 1);

                        objMatrix.LoadFromDataSource();
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

        static public void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            if (objExpenses == null)
            {
                objExpenses = new Expenses();
            }

            BubbleEvent = true;

            SAPbouiCOM.Form objForm = null;

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                if (objForm.TypeEx == "BYBEXPF001")
                {

                    if (eventInfo.BeforeAction && eventInfo.ItemUID == "0_U_G")
                    {
                        addRowMenu(objForm, eventInfo.Row);
                    }
                    else
                    {
                        removeRowMenu(objForm);
                    }

                }
                if (objForm.TypeEx == "BYBEXPF004")
                {

                    if (eventInfo.BeforeAction && eventInfo.ItemUID == "0_U_G")
                    {
                        addRowMenu1(objForm, eventInfo.Row);
                    }
                    else
                    {
                        removeRowMenu1(objForm);
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

        static public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && 
                !pVal.BeforeAction && 
                pVal.FormTypeEx == "BYBEXPF004" 
                && pVal.ItemUID == "21")
            {
                SAPbouiCOM.Form objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    B1.Expenses.Expenses.addExpenseLines(pVal);
                }
                else
                {
                    BYBB1MainObject.Instance.B1Application.MessageBox("Por favor grabe las modificaciones antes de contabilizar");
                }
            }

            if (
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK
                &&pVal.BeforeAction
                &&pVal.FormTypeEx == "BYBEXPF004"
                && pVal.ItemUID == "lbJE"
                )
            {
                BubbleEvent = false;

                SAPbouiCOM.Form objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.DBDataSource objDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                string strJE = objDS.GetValue("U_JEEntry", objDS.Offset);
                if (strJE.Trim().Length > 0)
                {
                    BYBB1MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_JournalPosting, "", strJE);
                }
                


            }




            
        }

        static public void LoadDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            #region Add Expense Request
            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.FormTypeEx == "BYBEXPF003")
            {

                if (!BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess)
                {
                    //B1.Expenses.Expenses.addExpenseRequest(BusinessObjectInfo);
                }
            }

            #endregion Add Expense Request

            #region Load Expense Window
            if(BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                && BusinessObjectInfo.FormTypeEx == "BYBEXPF004"
                && !BusinessObjectInfo.BeforeAction
                )
            {
                SAPbouiCOM.Form objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                objForm.Freeze(true);
                if (isExpensePosted(objForm))
                {
                    disableForm(objForm);
                    SAPbouiCOM.Item objItem = objForm.Items.Item("21");
                 
                    objItem.Visible = false;

                }
                else
                {
                    enableForm(objForm);
                    SAPbouiCOM.Item objItem = objForm.Items.Item("21");

                    objItem.Visible = true;
                }
                objForm.Freeze(false);
            }
            #endregion
        }


        #endregion EventManagement
        static public void addMainMenu()
        {

            if (objExpenses == null)
                objExpenses = new Expenses();

            try
            {
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(B1.Expenses.InteractionId.Default.T1ExpensesMainMenuId))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = B1.Expenses.LocalizationStrings.Default.T1ExpensesMainMenuString;
                    objMenuCreationParams.UniqueID = B1.Expenses.InteractionId.Default.T1ExpensesMainMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesMainParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesMainParentMenuId).SubMenus.AddEx(objMenuCreationParams);
                }

                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(B1.Expenses.InteractionId.Default.T1ExpensesConfigurationMenuId))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = B1.Expenses.LocalizationStrings.Default.T1ExpensesConfigurationMenuString;
                    objMenuCreationParams.UniqueID = B1.Expenses.InteractionId.Default.T1ExpensesConfigurationMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesConfigurationParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesConfigurationParentMenuId).SubMenus.AddEx(objMenuCreationParams);
                }

                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(B1.Expenses.InteractionId.Default.T1ExpensesConceptsMenuId))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = B1.Expenses.LocalizationStrings.Default.T1ExpensesConceptsMenuString;
                    objMenuCreationParams.UniqueID = B1.Expenses.InteractionId.Default.T1ExpensesConceptsMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesConceptsParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesConceptsParentMenuId).SubMenus.AddEx(objMenuCreationParams);
                }

                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(B1.Expenses.InteractionId.Default.T1ExpensesExpenseTypeMenuId))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = B1.Expenses.LocalizationStrings.Default.T1ExpensesExpenseTypeMenuString;
                    objMenuCreationParams.UniqueID = B1.Expenses.InteractionId.Default.T1ExpensesExpenseTypeMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesExpenseTypeParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesExpenseTypeParentMenuId).SubMenus.AddEx(objMenuCreationParams);
                }

                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(B1.Expenses.InteractionId.Default.T1ExpensesRequestMenuId))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = B1.Expenses.LocalizationStrings.Default.T1ExpensesRequestMenuString;
                    objMenuCreationParams.UniqueID = B1.Expenses.InteractionId.Default.T1ExpensesRequestMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesRequestParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesRequestParentMenuId).SubMenus.AddEx(objMenuCreationParams);
                }

                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(B1.Expenses.InteractionId.Default.T1ExpensesReportMenuId))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = B1.Expenses.LocalizationStrings.Default.T1ExpensesReportMenuString;
                    objMenuCreationParams.UniqueID = B1.Expenses.InteractionId.Default.T1ExpensesReportMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesReportParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesReportParentMenuId).SubMenus.AddEx(objMenuCreationParams);
                }



                #region MM

                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists("MMmnu1"))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Medios Magnéticos";
                    objMenuCreationParams.UniqueID = "MMmnu1";
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesReportParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item("BYBMS01").SubMenus.AddEx(objMenuCreationParams);
                }

                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists("MMmnu2"))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Configuración Medios Magnéticos";
                    objMenuCreationParams.UniqueID = "MMmnu2";
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesReportParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item("MMmnu1").SubMenus.AddEx(objMenuCreationParams);
                }

                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists("MMmnu3"))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Administrar Transacciones";
                    objMenuCreationParams.UniqueID ="MMmnu3";
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesReportParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item("MMmnu1").SubMenus.AddEx(objMenuCreationParams);
                }

                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists("MMmnu4"))
                {
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Asistente de Medios Magnéticos";
                    objMenuCreationParams.UniqueID = "MMmnu4";
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intCountMax = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Expenses.InteractionId.Default.T1ExpensesReportParentMenuId).SubMenus.Count;
                    objMenuCreationParams.Position = intCountMax + 1;
                    BYBB1MainObject.Instance.B1Application.Menus.Item("MMmnu1").SubMenus.AddEx(objMenuCreationParams);
                }




                #endregion

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.addMainMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.addMainMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        
        static private string localForm(string strFormId)
        {
            string strResult = "";

            if (objExpenses == null)
                objExpenses = new Expenses();

            try
            {
                if (strFormId == "BYBEXPF001")
                {
                    strResult = B1.Expenses.Resources.Expenses.BYBEXPF001;
                }
                else if (strFormId == "BYBEXPF002")
                {
                    strResult = B1.Expenses.Resources.Expenses.BYBEXPF002;
                }
                else if (strFormId == "BYBEXPF003")
                {
                    strResult = B1.Expenses.Resources.Expenses.BYBEXPF003;
                }
                else if (strFormId == "BYBEXPF004")
                {
                    strResult = B1.Expenses.Resources.Expenses.BYBEXPF004;
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strResult;

        }

        
        static private void addRowMenu(SAPbouiCOM.Form objForm, int intRowNumber)
        {
            string strMenuDescription = "";
            string strMenuId = "";
            try
            {
                strMenuId = "BYBEXPRC1";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Añadir linea";
                    //strMenuId = T1.B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Añadir Línea";
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    objForm.Menu.AddEx(objMenuCreationParams);

                }

                strMenuId = "BYBEXPRC2";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Eliminar linea";
                    //strMenuId = T1.B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Eliminar Línea";
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    objForm.Menu.AddEx(objMenuCreationParams);

                }

                BYBCache.Instance.addToCache("RC" + objForm.UniqueID, intRowNumber, BYBCache.objCachePriority.Default);

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

        static private void removeRowMenu(SAPbouiCOM.Form objForm)
        {
            string strMenuId = "";

            try
            {
                strMenuId = "BYBEXPRC1";
                if (BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    BYBB1MainObject.Instance.B1Application.Menus.RemoveEx(strMenuId);

                }
                strMenuId = "BYBEXPRC2";
                if (BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    BYBB1MainObject.Instance.B1Application.Menus.RemoveEx(strMenuId);

                }

                BYBCache.Instance.removeFromCache("RC" + objForm.UniqueID);
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

        static private void addRowMenu1(SAPbouiCOM.Form objForm, int intRowNumber)
        {
            string strMenuDescription = "";
            string strMenuId = "";
            try
            {
                strMenuId = "BYBEXPRC3";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Añadir linea";
                    //strMenuId = T1.B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Añadir Línea";
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    objForm.Menu.AddEx(objMenuCreationParams);

                }

                strMenuId = "BYBEXPRC4";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Eliminar linea";
                    //strMenuId = T1.B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Eliminar Línea";
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    objForm.Menu.AddEx(objMenuCreationParams);

                }

                BYBCache.Instance.addToCache("RC" + objForm.UniqueID, intRowNumber, BYBCache.objCachePriority.Default);

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

        static private void removeRowMenu1(SAPbouiCOM.Form objForm)
        {
            string strMenuId = "";

            try
            {
                strMenuId = "BYBEXPRC3";
                if (BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    BYBB1MainObject.Instance.B1Application.Menus.RemoveEx(strMenuId);

                }
                strMenuId = "BYBEXPRC4";
                if (BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    BYBB1MainObject.Instance.B1Application.Menus.RemoveEx(strMenuId);

                }

                BYBCache.Instance.removeFromCache("RC" + objForm.UniqueID);
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

        static public void addExpenseRequest(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {

            
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            XmlDocument oDocument = new XmlDocument();

            SAPbobsCOM.GeneralService oLegalType = null;
            SAPbobsCOM.GeneralData oLegalTypeData = null;
            SAPbobsCOM.GeneralDataParams oLegalTypeGeneralParams = null;

            double dbValue = 0;
            string strLegalType = "";


            string strCashAccount = "";
            string strControlAccount = "";
            string strBPCode = "";
            string strProject = "";

            SAPbobsCOM.Payments objPayments = null;

            try
            {
                int intPos = BusinessObjectInfo.ObjectKey.IndexOf("<Code>")+6;
                int intLength = BusinessObjectInfo.ObjectKey.IndexOf("</Code>") - intPos;
                string strCode = BusinessObjectInfo.ObjectKey.Substring(intPos, intLength);

                oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService(BusinessObjectInfo.Type);
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("Code", strCode);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                dbValue = oGeneralData.GetProperty("U_Value");
                strLegalType = oGeneralData.GetProperty("U_Type");
                strProject = oGeneralData.GetProperty("U_Project");

                oLegalType = oCompanyService.GetGeneralService("BYB_T1EXPU002");
                oLegalTypeGeneralParams = oLegalType.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oLegalTypeGeneralParams.SetProperty("Code", strLegalType);
                oLegalTypeData = oLegalType.GetByParams(oLegalTypeGeneralParams);

                strCashAccount = oLegalTypeData.GetProperty("U_CashAcct");
                strControlAccount = oLegalTypeData.GetProperty("U_FollowAcct");
                strBPCode = oLegalTypeData.GetProperty("U_CardCode");


                objPayments = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);

                
                objPayments.CardCode = strBPCode;
                objPayments.DocDate = DateTime.Now;
                objPayments.ControlAccount = strControlAccount;
                objPayments.TransferAccount = strCashAccount;
                objPayments.TransferSum = dbValue;
                objPayments.TransferDate = DateTime.Now;
                objPayments.TaxDate = DateTime.Now;
                //objPayments.ProjectCode = "";
                objPayments.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;

                if (objPayments.Add() != 0)
                {
                    string strMessage = BYBB1MainObject.Instance.B1Company.GetLastErrorDescription();
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

        static public void addExpenseLines(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Item oButton = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbouiCOM.DBDataSource oDS = null;
            string strObjectKey = "";
            string strRequestKey = "";
            string strCreditAccount = "";
            string strDefinitionKey = "";
            string strProject = "";
            string strCardCode = "";
            string strContDate = "";
            Hashtable hashWT = null;
            double dbDebitTotal = 0;


            SAPbobsCOM.GeneralService oExpensesService = null;
            SAPbobsCOM.GeneralService oExpenseRequestService = null;
            SAPbobsCOM.GeneralService oExpenseDefinitionService = null;
            SAPbobsCOM.GeneralService oConceptService = null;

            SAPbobsCOM.GeneralDataCollection oExpenseLines = null;

            SAPbobsCOM.GeneralData oExpenses = null;
            SAPbobsCOM.GeneralDataParams oExpensesParams = null;
            SAPbobsCOM.GeneralData oExpenseLineDetails = null;

            SAPbobsCOM.GeneralData oRequest = null;
            SAPbobsCOM.GeneralDataParams oRequestParams = null;

            SAPbobsCOM.GeneralData oDefinition = null;
            SAPbobsCOM.GeneralDataParams oDefinitionParams = null;

            SAPbobsCOM.GeneralData oConcept = null;
            SAPbobsCOM.GeneralDataParams oConceptParams = null;

            SAPbobsCOM.JournalEntries oJournal = null;
            SAPbobsCOM.JournalEntries_Lines oJournalLines = null;

            SAPbouiCOM.UserDataSource oDate = null;
            

            try
            {
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                hashWT = new Hashtable();
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oButton = oForm.Items.Item("1");
                    oButton.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }

                oDS = oForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                strContDate = oDS.GetValue("U_postDate", oDS.Offset);

                strObjectKey = oDS.GetValue("DocEntry", oDS.Offset);
                strRequestKey = oDS.GetValue("U_ExpenseCode", oDS.Offset);
                
                if (strContDate.Trim().Length == 0)
                {
                    BYBB1MainObject.Instance.B1Application.MessageBox("Por favor escriba la fecha de contabilización");
                }
                else
                {

                

                    oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();

                    oExpensesService = oCompanyService.GetGeneralService("BYB_T1EXPU004");

                    //oExpensesService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData.)

                    oExpensesParams = oExpensesService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oExpensesParams.SetProperty("DocEntry", strObjectKey);
                    oExpenses = oExpensesService.GetByParams(oExpensesParams);

                    //oExpenses.ToXMLString();
                    oExpenseLines = oExpenses.Child("BYB_T1EXP401");

                    oExpenseRequestService = oCompanyService.GetGeneralService("BYB_T1EXPU003");
                    oRequestParams = oExpenseRequestService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                    oRequestParams.SetProperty("Code", strRequestKey);
                    oRequest = oExpenseRequestService.GetByParams(oRequestParams);
                    strDefinitionKey = oRequest.GetProperty("U_Type");
                    strProject = oRequest.GetProperty("U_Project");

                    oExpenseDefinitionService = oCompanyService.GetGeneralService("BYB_T1EXPU002");
                    oDefinitionParams = oExpenseDefinitionService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oDefinitionParams.SetProperty("Code", strDefinitionKey);
                    oDefinition = oExpenseDefinitionService.GetByParams(oDefinitionParams);
                    strCreditAccount = oDefinition.GetProperty("U_FollowAcct");
                    strCardCode = oDefinition.GetProperty("U_CardCode");

                    oJournal = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    oJournal.ReferenceDate = BYBHelpers.Instance.getDateTimeFormString(strContDate);
                    oJournal.TaxDate = BYBHelpers.Instance.getDateTimeFormString(strContDate);
                    oJournal.DueDate = BYBHelpers.Instance.getDateTimeFormString(strContDate);
                    oJournal.Reference = strObjectKey;
                    oJournal.Memo = "Contabilizacion de la legalizacion No. " + strObjectKey;
                    if(T1.B1.Expenses.InteractionId.Default.T1ExpensesTransactionCode.Trim().Length > 0)
                    {
                        oJournal.TransactionCode = T1.B1.Expenses.InteractionId.Default.T1ExpensesTransactionCode.Trim();
                    }
                    oJournal.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES;
                    oJournal.AutomaticWT = SAPbobsCOM.BoYesNoEnum.tYES;
                    if(strProject.Trim().Length > 0)
                    {
                        oJournal.ProjectCode = strProject;
                    }
                    oJournalLines = oJournal.Lines;

                    if (strCardCode.Trim().Length > 0)
                    {
                        oJournalLines.ShortName = strCardCode;
                        oJournalLines.ControlAccount = strCreditAccount;
                    }
                    else
                    {
                        oJournalLines.AccountCode = strCreditAccount;
                    }
                    oJournalLines.Credit = dbDebitTotal;


                    for (int i = 0; i < oExpenseLines.Count; i++)
                    {
                        oExpenseLineDetails = oExpenseLines.Item(i);
                        double dbValue = oExpenseLineDetails.GetProperty("U_Value");
                        string strConceptKey = oExpenseLineDetails.GetProperty("U_Concept");
                        string strThirdParty = oExpenseLineDetails.GetProperty("U_ThirdParty");
                        string strLineProject = oExpenseLineDetails.GetProperty("U_Project");


                        string strDim1 = oExpenseLineDetails.GetProperty("U_DIM1");
                        string strDim2 = oExpenseLineDetails.GetProperty("U_DIM2");
                        string strDim3 = oExpenseLineDetails.GetProperty("U_DIM3");
                        string strDim4 = oExpenseLineDetails.GetProperty("U_DIM4");

                        if (strLineProject.Trim().Length == 0)
                        {
                            strLineProject = strProject;
                        }
                        if (strConceptKey.Trim().Length > 0)
                        {
                            oConceptService = oCompanyService.GetGeneralService("BYB_T1EXPU001");
                            oConceptParams = oConceptService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            oConceptParams.SetProperty("Code", strConceptKey);
                            oConcept = oConceptService.GetByParams(oConceptParams);

                            string strDebitAccount = oConcept.GetProperty("U_Account");
                            string strTaxCode = oConcept.GetProperty("U_TaxCode");
                            string strWTax = oConcept.GetProperty("U_WTax");


                            oJournalLines.Add();
                            oJournalLines.SetCurrentLine(i + 1);
                            if (strLineProject.Length > 0)
                            {
                                oJournalLines.ProjectCode = strLineProject;
                            }


                            if (strDim1.Length > 0)
                            {
                                oJournalLines.CostingCode = strDim1;
                            }
                            if (strDim2.Length > 0)
                            {
                                oJournalLines.CostingCode2 = strDim2;
                            }
                            if (strDim3.Length > 0)
                            {
                                oJournalLines.CostingCode3 = strDim3;
                            }
                            if (strDim4.Length > 0)
                            {
                                oJournalLines.CostingCode4 = strDim4;
                            }


                            //if (strThirdParty.Trim().Length > 0)
                            //{
                            //  oJournalLines.ShortName = strThirdParty;
                            //}
                            //else { 
                            oJournalLines.AccountCode = strDebitAccount;
                            //}
                            if (strTaxCode.Length > 0)
                            {
                                SAPbobsCOM.SalesTaxCodes oCode = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesTaxCodes);
                                if (oCode.GetByKey(strTaxCode))
                                {
                                    double dblRate = oCode.Rate;
                                    oJournalLines.TaxCode = strTaxCode;
                                    oJournalLines.TaxPostAccount = SAPbobsCOM.BoTaxPostAccEnum.tpa_PurchaseTaxAccount;
                                    dbDebitTotal += dbValue * (dblRate / 100);
                                }
                            }
                            if (strWTax.Length > 0)
                            {
                                ///TODO add WT logic for lines

                            }
                            oJournalLines.Debit = dbValue;
                            dbDebitTotal += dbValue;
                        }
                    }

                    oJournalLines.SetCurrentLine(0);
                    oJournalLines.Credit = dbDebitTotal;

                    if (!BYBB1MainObject.Instance.B1Company.InTransaction)
                    {

                        BYBB1MainObject.Instance.B1Company.StartTransaction();
                        if (oJournal.Add() != 0)
                        {
                            BYBB1MainObject.Instance.B1Application.MessageBox(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());

                        }
                        else
                        {
                            oExpenses.SetProperty("U_isPosted", "Y");
                            oExpenses.SetProperty("U_postDate", BYBHelpers.Instance.getDateTimeFormString(strContDate));
                            oExpenses.SetProperty("U_JEEntry", Convert.ToInt32(BYBB1MainObject.Instance.B1Company.GetNewObjectKey()));

                            oExpensesService.Update(oExpenses);
                            oForm.ActiveItem = "Item_2";
                            disableForm(oForm);
                            BYBB1MainObject.Instance.B1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            BYBB1MainObject.Instance.B1Application.MessageBox("La contabilización se llevo a cabo con éxito: Número de Asiento " + BYBB1MainObject.Instance.B1Company.GetNewObjectKey());

                        }

                    }
                    else
                    {
                        BYBB1MainObject.Instance.B1Application.MessageBox("La base de datos se encuentra bloqueada por otro proceso. Intente en unos instantes.");
                    }
                }
                

            }
            
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
                if(BYBB1MainObject.Instance.B1Company.InTransaction)
                { BYBB1MainObject.Instance.B1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
                if (BYBB1MainObject.Instance.B1Company.InTransaction)
                { BYBB1MainObject.Instance.B1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
            }

            
        }

        static private bool isExpensePosted(SAPbouiCOM.Form oForm)
        {
            bool blReturn = false;

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbouiCOM.DBDataSource oDS = null;

            SAPbobsCOM.GeneralService oExpensesService = null;
            SAPbobsCOM.GeneralData oExpenses = null;
            SAPbobsCOM.GeneralDataParams oExpensesParams = null;
            
            string strObjectKey = "";
            try
            {

                oDS = oForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");


                strObjectKey = oDS.GetValue("DocEntry", oDS.Offset);


                oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();

                oExpensesService = oCompanyService.GetGeneralService("BYB_T1EXPU004");

                oExpensesParams = oExpensesService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oExpensesParams.SetProperty("DocEntry", strObjectKey);
                oExpenses = oExpensesService.GetByParams(oExpensesParams);
                if(oExpenses != null)
                {
                    string strIsPosted = oExpenses.GetProperty("U_isPosted");
                    if(strIsPosted.Trim() == "Y")
                    {
                        blReturn = true;
                    }
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
            return blReturn;

        }

        static private void disableForm(SAPbouiCOM.Form oForm)
        {
            

            
            try
            {
                for(int i=0; i < oForm.Items.Count; i++)
                {
                    SAPbouiCOM.Item objItem = oForm.Items.Item(i);
                    if(objItem.UniqueID != "1" && objItem.UniqueID != "2" && objItem.UniqueID != "txtJE")
                    {
                        objItem.Enabled = false;
                    }

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

        static private void enableForm(SAPbouiCOM.Form oForm)
        {



            try
            {
                for (int i = 0; i < oForm.Items.Count; i++)
                {
                    SAPbouiCOM.Item objItem = oForm.Items.Item(i);
                    if (objItem.UniqueID != "1" && objItem.UniqueID != "2" && objItem.UniqueID!= "txtJE")
                    {
                        
                        objItem.Enabled = true;
                    }

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



    }
}
