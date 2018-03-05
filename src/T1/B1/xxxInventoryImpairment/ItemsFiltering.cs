using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;

namespace T1.B1.InventoryImpairment
{
    public class ItemsFiltering
    {
        static private ItemsFiltering objItemsFiltering = null;

        private ItemsFiltering()
        {

        }

        static public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form objForm = null;
            
            
            try
            {
                if (objItemsFiltering == null)
                    objItemsFiltering = new ItemsFiltering();
                
                if (!pVal.BeforeAction)
                {
                    if (pVal.MenuUID == B1.InventoryImpairment.InteractionId.Default.mnuItemsFilteringId)
                    {
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm(B1.InventoryImpairment.InteractionId.Default.frmItemFilteringFormId);
                        objFormCreationParams.FormType = InteractionId.Default.frmItemFilteringFormType;
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);

                        



                                                
                        objForm.Visible = true;
                    }
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ItemsFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ItemsFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objButton_ClickAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.UserDataSource objDS = null;
            string strItemValidation = "";
            string strGroupValidation = "";
            string strWarehouseValidation = "";

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                objDS = objForm.DataSources.UserDataSources.Item("UD_ITF");
                BYBCache.Instance.addToCache(CacheItemNames.Default.strItemFrom, objDS.ValueEx,BYBCache.objCachePriority.Default);
                strItemValidation += objDS.ValueEx;
                objDS = objForm.DataSources.UserDataSources.Item("UD_ITT");
                BYBCache.Instance.addToCache(CacheItemNames.Default.strItemTo, objDS.ValueEx, BYBCache.objCachePriority.Default);
                strItemValidation += objDS.ValueEx;

                objDS = objForm.DataSources.UserDataSources.Item("UD_ITG");
                BYBCache.Instance.addToCache(CacheItemNames.Default.strItemGroup, objDS.ValueEx, BYBCache.objCachePriority.Default);
                strGroupValidation += objDS.ValueEx;


                objDS = objForm.DataSources.UserDataSources.Item("UD_WHF");
                BYBCache.Instance.addToCache(CacheItemNames.Default.strWarehouseCodeFrom, objDS.ValueEx, BYBCache.objCachePriority.Default);
                strWarehouseValidation += objDS.ValueEx;
                objDS = objForm.DataSources.UserDataSources.Item("UD_WHT");
                BYBCache.Instance.addToCache(CacheItemNames.Default.strWarehoseCodeTo, objDS.ValueEx, BYBCache.objCachePriority.Default);
                strWarehouseValidation += objDS.ValueEx;

                if (strItemValidation.Length == 0 && strGroupValidation.Length == 0 && strWarehouseValidation.Length == 0)
                {
                    BYBB1MainObject.Instance.B1Application.MessageBox(InventoryImpairment.MessageStrings.Default.selectAnyValidationMessage);
                }
                else
                {
                    objForm.Close();
                    ItemImpairment.openForm();
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

        

        static public void addMainMenu()
        {
            string strMenuDescription = "";
            string strMenuId = "";
            try
            {
                if (objItemsFiltering == null)
                    objItemsFiltering = new ItemsFiltering();

                strMenuId = B1.InventoryImpairment.InteractionId.Default.mnuItemsFilteringId;
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = LocalizationStrings.Default.mnuItemsFilteringString;
                    strMenuId = InteractionId.Default.mnuItemsFilteringId;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    BYBB1MainObject.Instance.B1Application.Menus.Item(InteractionId.Default.mnuItemsFilteringParent).SubMenus.AddEx(objMenuCreationParams);
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ItemsFiltering.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ItemsFiltering.addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        static private string localForm(string strFormId)
        {
            string strResult = "";
            try
            {
                if (objItemsFiltering == null)
                    objItemsFiltering = new ItemsFiltering();

                if (strFormId == B1.InventoryImpairment.InteractionId.Default.frmItemFilteringFormId)
                {
                    strResult = B1.InventoryImpairment.Resources.InventoryImpairment.DIN001;
                    
                }

                else if (strFormId == B1.InventoryImpairment.InteractionId.Default.frmItemResultFormId)
                {
                    strResult = B1.InventoryImpairment.Resources.InventoryImpairment.DIN002;

                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ItemsFiltering.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ItemsFiltering.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strResult;

        }

        static public void objEditITF_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objItemsFiltering == null)
                objItemsFiltering = new ItemsFiltering();

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item(B1.InventoryImpairment.InteractionId.Default.frmEditItemFromId).Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditITF_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditITF_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objEditITT_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objItemsFiltering == null)
                objItemsFiltering = new ItemsFiltering();

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item(B1.InventoryImpairment.InteractionId.Default.frmEditItemToId).Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditITT_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditITT_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objEditITG_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objItemsFiltering == null)
                objItemsFiltering = new ItemsFiltering();

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item(B1.InventoryImpairment.InteractionId.Default.frmEditItemGroupId).Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditITG_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditITG_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objEditWF_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objItemsFiltering == null)
                objItemsFiltering = new ItemsFiltering();

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item(B1.InventoryImpairment.InteractionId.Default.frmEditWarehouseFromId).Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditWF_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditWF_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objEditWT_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objItemsFiltering == null)
                objItemsFiltering = new ItemsFiltering();

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item(B1.InventoryImpairment.InteractionId.Default.frmEditWarehouseToId).Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditWT_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ItemsFiltering.objEditWT_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

    }
}
