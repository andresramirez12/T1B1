using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;

namespace T1.B1.BPImpairment
{
    public class BPFiltering
    {
        static private BPFiltering objBPFiltering = null;

        private BPFiltering()
        {

        }

        static public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form objForm = null;
            
            SAPbouiCOM.ChooseFromList oChooseFromList = null;
            try
            {
                if (objBPFiltering == null)
                    objBPFiltering = new BPFiltering();
                
                if (!pVal.BeforeAction)
                {
                    if (pVal.MenuUID == InteractionId.Default.mnuBPFilteringId)
                    {
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm(InteractionId.Default.frmBPFilteringFormId);
                        objFormCreationParams.FormType = InteractionId.Default.frmBPFilteringFormType;
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);

                        oChooseFromList = objForm.ChooseFromLists.Item("CFLBPF");
                        setConditions(oChooseFromList,true);
                        oChooseFromList = objForm.ChooseFromLists.Item("CFLBPT");
                        setConditions(oChooseFromList,true);

                        

                        objForm.Visible = true;
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

        static public void objButton_ClickAfter(string ItemUID, string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.UserDataSource objDS = null;
            string strItemValidation = "";
            string strGroupValidation = "";
            string strWarehouseValidation = "";
            SAPbouiCOM.ChooseFromList oChooseFromList = null;

            try
            {
                if (objBPFiltering == null)
                    objBPFiltering = new BPFiltering();
                
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oChooseFromList = objForm.ChooseFromLists.Item("CFLBPF");
                setConditions(oChooseFromList, true);
                oChooseFromList = objForm.ChooseFromLists.Item("CFLBPT");
                setConditions(oChooseFromList, true);

                objDS = objForm.DataSources.UserDataSources.Item("UD_BPF");
                BYBCache.Instance.addToCache(B1.BPImpairment.CacheItemNames.Default.strBPFrom, objDS.ValueEx,BYBCache.objCachePriority.Default);
                strItemValidation += objDS.ValueEx;
                objDS = objForm.DataSources.UserDataSources.Item("UD_BPT");
                BYBCache.Instance.addToCache(B1.BPImpairment.CacheItemNames.Default.strIBPTo, objDS.ValueEx, BYBCache.objCachePriority.Default);
                strItemValidation += objDS.ValueEx;

                objDS = objForm.DataSources.UserDataSources.Item("UD_BPG");
                BYBCache.Instance.addToCache(B1.BPImpairment.CacheItemNames.Default.strBPGroup, objDS.ValueEx, BYBCache.objCachePriority.Default);
                strGroupValidation += objDS.ValueEx;

                objDS = objForm.DataSources.UserDataSources.Item("UD_CPP");
                BYBCache.Instance.addToCache(B1.BPImpairment.CacheItemNames.Default.strBPType, objDS.ValueEx, BYBCache.objCachePriority.Default);
                
                if (strItemValidation.Length == 0 && strGroupValidation.Length == 0 && strWarehouseValidation.Length == 0)
                {
                    BYBB1MainObject.Instance.B1Application.MessageBox(InventoryImpairment.MessageStrings.Default.selectAnyValidationMessage);
                }
                else
                {
                    objForm.Close();
                    BPImpairment.openForm();
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.objButton_ClickAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.objButton_ClickAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objEditBPF_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objBPFiltering == null)
                objBPFiltering = new BPFiltering();

            try
            {
                
                
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);
                
                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item("Item_1").Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);
                        
                        if (SelectedData != null)
                        {
                            objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                        }
                        
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.objEditBPF_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.objEditBPF_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objEditBPT_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objBPFiltering == null)
                objBPFiltering = new BPFiltering();

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item("Item_4").Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.objEditBPT_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.objEditBPT_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objEditBPG_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objBPFiltering == null)
                objBPFiltering = new BPFiltering();

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item("Item_24").Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.objEditBPF_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.objEditBPF_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void addMainMenu()
        {
            string strMenuDescription = "";
            string strMenuId = "";

            if (objBPFiltering == null)
                objBPFiltering = new BPFiltering();

            try
            {
                

                strMenuId = B1.BPImpairment.InteractionId.Default.mnuBPFilteringId;
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = LocalizationStrings.Default.mnuItemsFilteringString;
                    strMenuId = InteractionId.Default.mnuBPFilteringId;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    BYBB1MainObject.Instance.B1Application.Menus.Item(InteractionId.Default.mnuBPFilteringParent).SubMenus.AddEx(objMenuCreationParams);
                }

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

            if (objBPFiltering == null)
                objBPFiltering = new BPFiltering();

            try
            {
                

                if (strFormId == B1.BPImpairment.InteractionId.Default.frmBPFilteringFormId)
                {
                    strResult = B1.BPImpairment.Resources.BPImpairment.DBP001;
                    
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

        static private void setConditions(SAPbouiCOM.ChooseFromList oCFL, bool setConditions)
        {
            SAPbouiCOM.Conditions oConditions = null;
            SAPbouiCOM.Condition oCond = null;
            try
            {
                if (setConditions)
                {

                    oConditions = oCFL.GetConditions();
                    oCond = oConditions.Add();
                    oCond.Alias = "CardType";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                    oCond.CondVal = "L";
                    oCFL.SetConditions(oConditions);
                }
                else
                {
                    oCFL.SetConditions(null);
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.setConditions", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.setConditions", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }
    }
}
