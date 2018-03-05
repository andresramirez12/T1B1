using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;
using System.Globalization;

namespace T1.B1.InventoryImpairment
{
    public class ItemImpairment
    {
        static private ItemImpairment objItemsImpairement = null;

        private ItemImpairment()
        {

        }
        
        public static void openForm()
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.ChooseFromList oChooseFromList = null;
            SAPbouiCOM.FormCreationParams objFormCreationParams = null;

            try
            {
                if (objItemsImpairement == null)
                {
                    objItemsImpairement = new ItemImpairment();
                }
                
                objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objFormCreationParams.XmlData = localForm(B1.InventoryImpairment.InteractionId.Default.frmItemResultFormId);
                objFormCreationParams.FormType = InteractionId.Default.frmItemResultFormType;
                objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);

                SAPbouiCOM.ComboBox objCombo = objForm.Items.Item("Item_19").Specific;
                TransactionCodes.TransactionCodes.fillValidValuesFromDB(ref objCombo);
                loadDataGrid(ref objForm);

                oChooseFromList = objForm.ChooseFromLists.Item("CFL_INVA");
                setConditions(oChooseFromList, true);
                oChooseFromList = objForm.ChooseFromLists.Item("CFL_DETA");
                setConditions(oChooseFromList, true);


               
                

                objForm.Visible = true;



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

        static public void objButtonConytab_ClickAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            XmlDocument objDocument = null;
            XmlNodeList objNodeList = null;
            double dbTotalSum = 0;
            SAPbouiCOM.DataTable objDataTable = null;
            SAPbouiCOM.Form oForm;
            string strDataTableXML = "";

            SAPbobsCOM.JournalEntries objJournalEntry = null;
            string strInventAccount = "";
            string strImpairmentAccount = "";
            string strReference1 = "";
            string strReference2 = "";
            string strReference3 = "";
            string strTransactionCode = "";
            string  strIsRecover = "";
            SAPbouiCOM.ChooseFromList oChooseFromList = null;

            SAPbouiCOM.UserDataSource oUsrDS = null;
            SAPbouiCOM.DBDataSource oDBDS = null;

            int intRetCode = -1;

            

            try
            {
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                oUsrDS = oForm.DataSources.UserDataSources.Item("UD_INVA");
                strInventAccount = oUsrDS.ValueEx;
                oUsrDS = oForm.DataSources.UserDataSources.Item("UD_DETA");
                strImpairmentAccount = oUsrDS.ValueEx;
                oUsrDS = oForm.DataSources.UserDataSources.Item("UD_REF1");
                strReference1 = oUsrDS.ValueEx;
                oUsrDS = oForm.DataSources.UserDataSources.Item("UD_REF2");
                strReference2 = oUsrDS.ValueEx;
                oUsrDS = oForm.DataSources.UserDataSources.Item("UD_REF3");
                strReference3 = oUsrDS.ValueEx;
                oUsrDS = oForm.DataSources.UserDataSources.Item("UD_REC");
                strIsRecover = oUsrDS.ValueEx;
                oDBDS = oForm.DataSources.DBDataSources.Item("OTRC");
                strTransactionCode = oDBDS.GetValue(0,0);

                if (strTransactionCode.Length > 0 && strImpairmentAccount.Length > 0 && strImpairmentAccount.Length > 0)
                {

                    oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("SAPINIM");
                    oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                    

                    oChildren = oGeneralData.Child("SAPNIM1");
                    objDocument = new XmlDocument();

                    objDataTable = oForm.DataSources.DataTables.Item("DT_ITR");
                    strDataTableXML = objDataTable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly);
                    objDocument.LoadXml(strDataTableXML);
                    objNodeList = objDocument.SelectNodes("/DataTable/Rows/Row/Cells[./Cell[8]/ColumnUid/text() = 'Checked' and ./Cell[8]/Value/text() = 'Y']");
                    if (objNodeList != null)
                    {
                        foreach (XmlNode xn in objNodeList)
                        {
                            double dbImp = BYBHelpers.Instance.getStandarNumericValue(xn.SelectSingleNode("./Cell[12]/Value").InnerText);
                            
                            string strItemCode = xn.SelectSingleNode("./Cell[1]/Value").InnerText;
                            string strWhsCode = xn.SelectSingleNode("./Cell[6]/Value").InnerText;
                            

                            oChild = oChildren.Add();

                            oChild.SetProperty("U_ItemCode", strItemCode);
                            oChild.SetProperty("U_WhsCode", strWhsCode);
                            oChild.SetProperty("U_Impair", dbImp);

                            dbTotalSum += dbImp;


                        }

                    }

                    if(dbTotalSum > 0)
                    {
                        objJournalEntry = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        objJournalEntry.Reference = strReference1;
                        objJournalEntry.Reference2 = strReference2;
                        objJournalEntry.Reference3 = strReference3;
                        objJournalEntry.ReferenceDate = DateTime.Now;
                        objJournalEntry.TransactionCode = strTransactionCode;
                        objJournalEntry.Lines.AccountCode = strInventAccount;
                        if(strIsRecover == "Y")
                        {
                            objJournalEntry.Lines.Debit = dbTotalSum;
                        }
                        else
                        {
                            objJournalEntry.Lines.Credit = dbTotalSum;
                        }
                        
                        objJournalEntry.Lines.Add();
                        objJournalEntry.Lines.AccountCode = strImpairmentAccount;

                        if (strIsRecover == "Y")
                        {
                            objJournalEntry.Lines.Credit = dbTotalSum;
                        }
                        else
                        {
                            objJournalEntry.Lines.Debit = dbTotalSum;
                        }

                        intRetCode = objJournalEntry.Add();
                        if(intRetCode != 0)
                        {
                            BYBB1MainObject.Instance.B1Application.MessageBox(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());
                        }
                        else
                        {
                            oGeneralData.SetProperty("U_JETransId", Convert.ToInt32(BYBB1MainObject.Instance.B1Company.GetNewObjectKey()));
                            oGeneralData.SetProperty("U_Total", dbTotalSum);
                            oGeneralParams = oGeneralService.Add(oGeneralData);


                            int intNumber = oGeneralParams.GetProperty("DocEntry");
                            if (intNumber > 0)
                            {
                                BYBB1MainObject.Instance.B1Application.MessageBox(MessageStrings.Default.impairmentDoneMessage + "Asiento número: " + BYBB1MainObject.Instance.B1Company.GetNewObjectKey());
                                oChooseFromList = oForm.ChooseFromLists.Item("CFL_INVA");
                                setConditions(oChooseFromList, false);
                                oChooseFromList = oForm.ChooseFromLists.Item("CFL_DETA");
                                setConditions(oChooseFromList, false);

                            }
                            
                            else
                            {
                                BYBB1MainObject.Instance.B1Application.MessageBox("Se produjo un error al registrar el detalle de losartículos. El deterioro se contabilizó con el asiento número: " + BYBB1MainObject.Instance.B1Company.GetNewObjectKey());
                            }

                            
                            oForm.Close();


                        }



                    }
                    else
                    {
                        BYBB1MainObject.Instance.B1Application.MessageBox(MessageStrings.Default.selectTransactionMessage);
                    }
                }
                else
                {
                    BYBB1MainObject.Instance.B1Application.MessageBox(MessageStrings.Default.selectTransactionMessage);
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

        static public void objButton_ClickAfter(string FormUId, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable objDataTable = null;
            SAPbouiCOM.UserDataSource objUserDs = null;
            SAPbouiCOM.Form oForm;
            string strDataTableXML = "";
            XmlDocument objDocument = null;
            XmlNodeList objNodeList = null;
            double dbTotalSum = 0;
            try
            {
                objDocument = new XmlDocument();
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objDataTable = oForm.DataSources.DataTables.Item("DT_ITR");
                strDataTableXML = objDataTable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly);
                objDocument.LoadXml(strDataTableXML);
                objNodeList = objDocument.SelectNodes("/DataTable/Rows/Row/Cells[./Cell[8]/ColumnUid/text() = 'Checked' and ./Cell[8]/Value/text() = 'Y']");
                if(objNodeList != null)
                {
                    foreach(XmlNode xn in objNodeList)
                    {
                        //double dbTotal = Convert.ToDouble(xn.SelectSingleNode("./Cell[12]/Value").InnerText);
                        double dbTotal = BYBHelpers.Instance.getStandarNumericValue(xn.SelectSingleNode("./Cell[12]/Value").InnerText);

                        dbTotalSum += dbTotal;


                    }

                }
                objUserDs = oForm.DataSources.UserDataSources.Item("UD_SUM");

                CultureInfo objTest = new CultureInfo("en-US");
                objTest.NumberFormat.NumberDecimalSeparator = ".";
                objTest.NumberFormat.NumberGroupSeparator = ",";

                objUserDs.ValueEx = Convert.ToString(dbTotalSum, objTest);
                



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ItemImpairment.objButton_ClickAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ItemImpairment.objButton_ClickAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static private string localForm(string strFormId)
        {
            string strResult = "";
            try
            {
                if (strFormId == B1.InventoryImpairment.InteractionId.Default.frmItemResultFormId)
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

        static private void loadDataGrid(ref SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DataTable objDataTable = null;
            string strItemFrom ="";
            string strItemTo = "";
            string strGroupCode = "";
            string strWareHouseFrom = "";
            string strWareHouseTo = "";
            string strComparisonItem = "";
            string strComparisonWH = "";
            string strFilter = "";
            string strQuery = "";
            try
            {
                objDataTable = oForm.DataSources.DataTables.Item("DT_ITR");
                strItemFrom = BYBCache.Instance.getFromCache(CacheItemNames.Default.strItemFrom);
                strItemTo = BYBCache.Instance.getFromCache(CacheItemNames.Default.strItemTo);
                strGroupCode = BYBCache.Instance.getFromCache(CacheItemNames.Default.strItemGroup);
                strWareHouseFrom = BYBCache.Instance.getFromCache(CacheItemNames.Default.strWarehouseCodeFrom);
                strWareHouseTo = BYBCache.Instance.getFromCache(CacheItemNames.Default.strWarehoseCodeTo);

                if (strItemFrom.Length >0 && strItemTo.Length >0)
                {
                    strComparisonItem = ">";
                    strFilter = InventoryImpairment.Resources.sqlQueries.sql0002;
                    strFilter = string.Format(strFilter, strComparisonItem, strItemFrom);
                    strFilter += " " + InventoryImpairment.Resources.sqlQueries.sql0003;
                    strFilter = string.Format(strFilter, strItemTo);
                }
                else
                {
                    if (strItemFrom.Length > 0)
                    {
                        strFilter = InventoryImpairment.Resources.sqlQueries.sql0002;
                        strFilter = string.Format(strFilter, strComparisonItem, strItemFrom);

                    }
                    else if (strItemTo.Length > 0)
                    {
                        strFilter = InventoryImpairment.Resources.sqlQueries.sql0002;
                        strFilter = string.Format(strFilter, strComparisonItem, strItemTo);
                    }

                }

                if (strGroupCode.Length > 0)
                {
                    strFilter += " " + InventoryImpairment.Resources.sqlQueries.sql0004;
                    strFilter = string.Format(strFilter, strGroupCode);

                }



                if (strWareHouseFrom.Length > 0 && strWareHouseTo.Length > 0)
                {
                    strComparisonWH = ">";
                    strFilter += " " + InventoryImpairment.Resources.sqlQueries.sql0005;
                    strFilter = string.Format(strFilter, strComparisonWH, strWareHouseFrom);
                    strFilter += " " + InventoryImpairment.Resources.sqlQueries.sql0006;
                    strFilter = string.Format(strFilter, strWareHouseTo);
                }
                else
                {
                    if (strWareHouseFrom.Length > 0)
                    {
                        strFilter += " " + InventoryImpairment.Resources.sqlQueries.sql0005;
                        strFilter = string.Format(strFilter, strComparisonWH, strWareHouseFrom);

                    }
                    else if (strWareHouseTo.Length > 0)
                    {
                        strFilter += " " +  InventoryImpairment.Resources.sqlQueries.sql0005;
                        strFilter = string.Format(strFilter, strComparisonWH, strWareHouseTo);
                    }

                }

                
                //TODO Add Analytics Button - Check Query
                //string strPath = AppDomain.CurrentDomain.BaseDirectory + @"localResources\chart_bar.png";
                strQuery = InventoryImpairment.Resources.sqlQueries.sql0001 + " " + strFilter;

                objDataTable.ExecuteQuery(strQuery);

                SAPbouiCOM.EditTextColumn oEditTExt = null;
                SAPbouiCOM.GridColumn objGridColumn = null;

                

                SAPbouiCOM.Grid objGrid = oForm.Items.Item("Item_15").Specific;

                objGrid.CollapseLevel = 1;


                objGrid.Columns.Item(7).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;


                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;
                objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                oEditTExt.LinkedObjectType = "4";


                objGridColumn = objGrid.Columns.Item(2);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;


                
                objGridColumn = objGrid.Columns.Item(3);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(4);
                objGridColumn.Editable = false;
                


                

                objGridColumn = objGrid.Columns.Item(5);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;
                objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                oEditTExt.LinkedObjectType = "64";

                objGridColumn = objGrid.Columns.Item(6);
                objGridColumn.Editable = false;


                

                objGridColumn = objGrid.Columns.Item(8);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(9);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;


                objGridColumn = objGrid.Columns.Item(10);
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(11);
                objGridColumn.RightJustified = true;


                

                
                

                objGrid.AutoResizeColumns();

                
                
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

        static public void objEditInvAcc_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objItemsImpairement == null)
                objItemsImpairement = new ItemImpairment();

            try
            {


                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item(B1.InventoryImpairment.InteractionId.Default.frmCmbInvAcctId).Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ItemImpairment.objEditInvAcc_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ItemImpairment.objEditInvAcc_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objEditInvDet_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objItemsImpairement == null)
                objItemsImpairement = new ItemImpairment();

            try
            {


                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item(B1.InventoryImpairment.InteractionId.Default.frmCmbImpAccId).Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ItemImpairment.objEditInvDet_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ItemImpairment.objEditInvDet_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objGridColumn_LostFocusAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            double dbPrice = 0;
            double dbStock = 0;
            double dbImpair = 0;
            double dbTotal = 0;
            
            SAPbouiCOM.DataTable objDataTable = null;
            SAPbouiCOM.Grid objGrid = null;
            SAPbouiCOM.Form objForm = null;

            int intDataTableIndex = -1;
            double dbPV = 0;
            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);
                objForm.Freeze(true);
                objGrid = objForm.Items.Item(pVal.ItemUID).Specific;
                objDataTable = objGrid.DataTable;
                intDataTableIndex = objGrid.GetDataTableRowIndex(pVal.Row);
                dbPrice = objDataTable.GetValue(8, intDataTableIndex);
                dbStock = objDataTable.GetValue(9, intDataTableIndex);
                dbImpair = objDataTable.GetValue(10, intDataTableIndex);
                dbTotal = dbStock * dbImpair;
                objDataTable.SetValue(11, intDataTableIndex, dbTotal);
                //objDataTable.SetValue(12, intDataTableIndex, dbDue - dbPV);
                //objDataTable.SetValue(14, intDataTableIndex, (dbDue - dbPV) * (dbPercent / 100));



                objForm.Freeze(false);

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "InventoryImpairment.objGridColumn_LostFocusAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "InventoryImpairment.objGridColumn_LostFocusAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            finally
            {
                objForm.Freeze(false);
            }
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
                    oCond.Alias = "Postable";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = "Y";
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
                BYBExceptionHandling.reportException(er.Message, "ItemImpairment.setConditions", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ItemImpairment.setConditions", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

    }
}
