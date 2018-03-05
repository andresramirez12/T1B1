using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;
using System.Globalization;

namespace T1.B1.BPImpairment
{
    public class BPImpairment
    {
        static private BPImpairment objBPImpairement = null;

        private BPImpairment()
        {

        }
        
        public static void openForm()
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.ChooseFromList oChooseFromList = null;
            SAPbouiCOM.FormCreationParams objFormCreationParams = null;


            try
            {
                if (objBPImpairement == null)
                    objBPImpairement = new BPImpairment();
                
                
                objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objFormCreationParams.XmlData = localForm(B1.BPImpairment.InteractionId.Default.frmBPResultFormId);
                objFormCreationParams.FormType = InteractionId.Default.frmBPResultFormType;
                objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);

                SAPbouiCOM.ComboBox objCombo = objForm.Items.Item("Item_19").Specific;
                TransactionCodes.TransactionCodes.fillValidValuesFromDB(ref objCombo);
                loadDataGrid(ref objForm);

                
                

                oChooseFromList = objForm.ChooseFromLists.Item("CFL_BPA");
                setConditions(oChooseFromList, true);
                oChooseFromList = objForm.ChooseFromLists.Item("CFL_DETA");
                setConditions(oChooseFromList, true);

                objForm.Visible = true;



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "BPImpairment.openForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.openForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }


        static public void objEditBPAcc_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objBPImpairement == null)
                objBPImpairement = new BPImpairment();

            try
            {


                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item("Item_10").Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPImpairment.objEditBPAcc_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPImpairment.objEditBPAcc_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objEditBPDet_ChooseFromListAfter(SAPbouiCOM.DataTable SelectedData, string FormUID)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.UserDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            if (objBPImpairement == null)
                objBPImpairement = new BPImpairment();

            try
            {


                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);

                oEdit = (SAPbouiCOM.EditText)objForm.Items.Item("Item_11").Specific;
                objDS = objForm.DataSources.UserDataSources.Item(oEdit.DataBind.Alias);

                if (SelectedData != null)
                {
                    objDS.Value = Convert.ChangeType(SelectedData.GetValue(0, 0), Type.GetType("System.String"));
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPImpairment.objEditBPDet_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPImpairment.objEditBPDet_ChooseFromListAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objButton_PressedAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            XmlDocument objDocument = null;
            XmlNodeList objNodeList = null;
            SAPbouiCOM.ChooseFromList oChooseFromList = null;
            
            SAPbouiCOM.DataTable objDataTable = null;
            SAPbouiCOM.Form oForm;
            string strDataTableXML = "";

            SAPbobsCOM.JournalEntries objJournalEntry = null;
            string strAccount = "";
            string strImpairmentAccount = "";
            string strReference1 = "";
            string strReference2 = "";
            string strReference3 = "";
            string strTransactionCode = "";
            string strIsRecover = "";
            double dbTotal = 0;

            SAPbouiCOM.UserDataSource oUsrDS = null;
            SAPbouiCOM.DBDataSource oDBDS = null;

            int intRetCode = -1;

            try
            {
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                oUsrDS = oForm.DataSources.UserDataSources.Item("UD_BPA");
                strAccount = oUsrDS.ValueEx;
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
                strTransactionCode = oDBDS.GetValue(0, 0);
                oUsrDS = oForm.DataSources.UserDataSources.Item("UD_SUM");
                dbTotal = BYBHelpers.Instance.getStandarNumericValue(oUsrDS.ValueEx);

                if (strTransactionCode.Length > 0 && strAccount.Length > 0 && strImpairmentAccount.Length > 0)
                {

                    oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("SAPBPIM");
                    oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));


                    oChildren = oGeneralData.Child("SAPPIM1");
                    objDocument = new XmlDocument();

                    objDataTable = oForm.DataSources.DataTables.Item("DT_BPI");
                    strDataTableXML = objDataTable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly);
                    objDocument.LoadXml(strDataTableXML);
                    objNodeList = objDocument.SelectNodes("/DataTable/Rows/Row/Cells[./Cell[7]/ColumnUid/text() = 'Checked' and ./Cell[7]/Value/text() = 'Y']");
                    if (objNodeList != null)
                    {
                        foreach (XmlNode xn in objNodeList)
                        {
                            string strCardCode = xn.SelectSingleNode("./Cell[1]/Value").InnerText;
                            string strDocument = xn.SelectSingleNode("./Cell[3]/Value").InnerText;
                            double dbImp = BYBHelpers.Instance.getStandarNumericValue(xn.SelectSingleNode("./Cell[15]/Value").InnerText);



                            oChild = oChildren.Add();

                            oChild.SetProperty("U_CardCode", strCardCode);
                            oChild.SetProperty("U_DocEntry", strDocument);
                            oChild.SetProperty("U_Impair", dbImp);

                            


                        }

                    }

                    if (dbTotal > 0)
                    {
                        objJournalEntry = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        objJournalEntry.Reference = strReference1;
                        objJournalEntry.Reference2 = strReference2;
                        objJournalEntry.Reference3 = strReference3;
                        objJournalEntry.ReferenceDate = DateTime.Now;
                        objJournalEntry.TransactionCode = strTransactionCode;
                        objJournalEntry.Lines.AccountCode = strAccount;
                        if (strIsRecover == "Y")
                        {
                            objJournalEntry.Lines.Debit = dbTotal;
                        }
                        else
                        {
                            objJournalEntry.Lines.Credit = dbTotal;
                        }

                        objJournalEntry.Lines.Add();
                        objJournalEntry.Lines.AccountCode = strImpairmentAccount;

                        if (strIsRecover == "Y")
                        {
                            objJournalEntry.Lines.Credit = dbTotal;
                        }
                        else
                        {
                            objJournalEntry.Lines.Debit = dbTotal;
                        }

                        intRetCode = objJournalEntry.Add();
                        if (intRetCode != 0)
                        {
                            BYBB1MainObject.Instance.B1Application.MessageBox(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription());
                        }
                        else
                        {
                            oGeneralData.SetProperty("U_JETransId", Convert.ToInt32(BYBB1MainObject.Instance.B1Company.GetNewObjectKey()));
                            oGeneralData.SetProperty("U_Total", dbTotal);
                            oGeneralParams = oGeneralService.Add(oGeneralData);
                            int intNumber = oGeneralParams.GetProperty("DocEntry");
                            if (intNumber > 0)
                            {
                                BYBB1MainObject.Instance.B1Application.MessageBox(MessageStrings.Default.impairmentDoneMessage + "Asiento número: " + BYBB1MainObject.Instance.B1Company.GetNewObjectKey());
                                oChooseFromList = oForm.ChooseFromLists.Item("CFL_BPA");
                                setConditions(oChooseFromList, false);
                                oChooseFromList = oForm.ChooseFromLists.Item("CFL_DETA");
                                setConditions(oChooseFromList, false);
                                
                            }
                            else
                            {
                                BYBB1MainObject.Instance.B1Application.MessageBox("Se produjo un error al registrar el detalle de los socios de negocio. El deterioro se contabilizó con el asiento número: " + BYBB1MainObject.Instance.B1Company.GetNewObjectKey());

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
                BYBExceptionHandling.reportException("", "BPImpairment.objButton_PressedAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.objButton_PressedAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void objButton_ClickAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
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
                objDataTable = oForm.DataSources.DataTables.Item("DT_BPI");
                strDataTableXML = objDataTable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly);
                objDocument.LoadXml(strDataTableXML);
                objNodeList = objDocument.SelectNodes("/DataTable/Rows/Row/Cells[./Cell[7]/ColumnUid/text() = 'Checked' and ./Cell[7]/Value/text() = 'Y']");
                if(objNodeList != null)
                {
                    foreach(XmlNode xn in objNodeList)
                    {
                        

                        double dbNewPrice = BYBHelpers.Instance.getStandarNumericValue(xn.SelectSingleNode("./Cell[15]/Value").InnerText); // (totalDebt * dbIndex / 365) * totalDays;
                        //double dbNewPrice = Convert.ToDouble(xn.SelectSingleNode("./Cell[15]/Value").InnerText);
                        
                        dbTotalSum += dbNewPrice;


                    }

                }

                //objDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly, objDocument.OuterXml);

                objUserDs = oForm.DataSources.UserDataSources.Item("UD_SUM");

                CultureInfo objTest = new CultureInfo("en-US");
                objTest.NumberFormat.NumberDecimalSeparator = ".";
                objTest.NumberFormat.NumberGroupSeparator = ",";



                objUserDs.ValueEx = Convert.ToString(dbTotalSum, objTest);
                



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "BPImpairment.objButton_ClickAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.objButton_ClickAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static private string localForm(string strFormId)
        {
            string strResult = "";
            if (objBPImpairement == null)
                objBPImpairement = new BPImpairment();

            try
            {
                if (strFormId == InteractionId.Default.frmBPResultFormId)
                {
                    strResult = Resources.BPImpairment.DBP002;

                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "BPImpairment.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strResult;

        }

        static private void loadDataGrid(ref SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DataTable objDataTable = null;
            string strBPFrom ="";
            string strBPTo = "";
            string strGroupCode = "";
            string strBPType = "";
            string strFilter = "";
            string strQuery = "";
            string strComparisonItem = "";

            if (objBPImpairement == null)
                objBPImpairement = new BPImpairment();

            try
            {
                objDataTable = oForm.DataSources.DataTables.Item("DT_BPI");
                strBPFrom = BYBCache.Instance.getFromCache(CacheItemNames.Default.strBPFrom);
                strBPTo = BYBCache.Instance.getFromCache(CacheItemNames.Default.strIBPTo);
                strGroupCode = BYBCache.Instance.getFromCache(CacheItemNames.Default.strBPGroup);
                strBPType = BYBCache.Instance.getFromCache(CacheItemNames.Default.strBPType);

                if (strBPType == "Y")
                {
                    strBPType = Resources.sqlQueries.BPTypePUR;
                }
                else
                {
                    strBPType = Resources.sqlQueries.BPTypeINV;
                }


                if (strBPFrom.Length > 0 && strBPTo.Length > 0)
                {
                    strComparisonItem = ">";
                    strFilter = Resources.sqlQueries.sql0002;
                    strFilter = string.Format(strFilter, strComparisonItem, strBPFrom);
                    strFilter += " " + Resources.sqlQueries.sql0003;
                    strFilter = string.Format(strFilter, strBPTo);
                }
                else
                {
                    if (strBPFrom.Length > 0)
                    {
                        strFilter = Resources.sqlQueries.sql0002;
                        strFilter = string.Format(strFilter, strComparisonItem, strBPFrom);

                    }
                    else if (strBPTo.Length > 0)
                    {
                        strFilter = Resources.sqlQueries.sql0002;
                        strFilter = string.Format(strFilter, strComparisonItem, strBPTo);
                    }

                }

                if (strGroupCode.Length > 0)
                {
                    strFilter += " " + Resources.sqlQueries.sql0004;
                    strFilter = string.Format(strFilter, strGroupCode);

                }

                strQuery = string.Format(Resources.sqlQueries.sql0001, strBPType) + " " + strFilter;

                objDataTable.ExecuteQuery(strQuery);

                SAPbouiCOM.EditTextColumn oEditTExt = null;
                SAPbouiCOM.GridColumn objGridColumn = null;
                
                SAPbouiCOM.Grid objGrid = oForm.Items.Item("Item_15").Specific;

                //objGrid.Columns.Item(11).Type = SAPbouiCOM.BoGridColumnType.gct_Picture;
                //objGrid.Columns.Item(11).Width = 20;

                objGrid.Columns.Item(6).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

                objGrid.CollapseLevel = 1;


                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;


                objGridColumn = objGrid.Columns.Item(2);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;


                objGrid.Columns.Item(2).Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oEditTExt = (SAPbouiCOM.EditTextColumn)objGrid.Columns.Item(2);

                if (strBPType == Resources.sqlQueries.BPTypePUR)
                    oEditTExt.LinkedObjectType = InteractionId.Default.PurchaseObjectType;
                else
                    oEditTExt.LinkedObjectType = InteractionId.Default.InvoiceObjectType;


                objGridColumn = objGrid.Columns.Item(3);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;


                objGridColumn = objGrid.Columns.Item(4);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(5);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(7);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(8);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(9);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(10);
                objGridColumn.RightJustified = true;


                SAPbouiCOM.CommonSetting setting = objGrid.CommonSetting;

                int greenColor = Color.Green.R | (Color.Green.G << 8) | (Color.Green.B << 16);
                int yellowColor = Color.Yellow.R | (Color.Yellow.G << 8) | (Color.Yellow.B << 16);
                int redColor = Color.Red.R | (Color.Red.G << 8) | (Color.Red.B << 16);

                for (int i = 1; i <= objGrid.Rows.Count; i++)
                {
                    int intDataTableIndex = objGrid.GetDataTableRowIndex(i - 1);
                    if (intDataTableIndex >= 0)
                    {
                        int intAgeValue = objDataTable.GetValue(5, intDataTableIndex);
                        if (intAgeValue <= 30)
                            setting.SetCellBackColor(i, 6, greenColor);
                        else if (intAgeValue > 30 && intAgeValue <= 60)
                            setting.SetCellBackColor(i, 6, yellowColor);
                        else
                            setting.SetCellBackColor(i, 6, redColor);
                    }
                }


                objGridColumn = objGrid.Columns.Item(11);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(12);
                objGridColumn.Editable = true;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(13);
                objGridColumn.Editable = true;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(14);
                objGridColumn.Editable = true;
                objGridColumn.RightJustified = true;


                objGrid.AutoResizeColumns();

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPImpairment.loadDataGrid", er, 1, System.Diagnostics.EventLogEntryType.Error);
                
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPImpairment.loadDataGrid", er, 1, System.Diagnostics.EventLogEntryType.Error);
                
            }
        }

        static public void objGridColumn_LostFocusAfter(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            double dbDue = 0;
            double dbIndex = 0;
            double intAge = 0;
            double dbPercent = 0;
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
                dbDue = objDataTable.GetValue(9, intDataTableIndex);
                dbIndex = objDataTable.GetValue(10, intDataTableIndex);
                dbPercent = objDataTable.GetValue(13, intDataTableIndex);
                intAge = objDataTable.GetValue(5, intDataTableIndex);
                dbPV = dbDue / (Math.Pow((1 + ((dbIndex / 100)/365)),intAge));
                objDataTable.SetValue(11, intDataTableIndex, dbPV);
                objDataTable.SetValue(12, intDataTableIndex, dbDue - dbPV);
                objDataTable.SetValue(14, intDataTableIndex, (dbDue - dbPV) * (dbPercent/100));

                

                objForm.Freeze(false);

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "BPImpairment.objGridColumn_LostFocusAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "BPImpairment.objGridColumn_LostFocusAfter", er, 1, System.Diagnostics.EventLogEntryType.Error);
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
                BYBExceptionHandling.reportException(er.Message, "BPImpairment.setConditions", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPImpairment.setConditions", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        


    }
}
