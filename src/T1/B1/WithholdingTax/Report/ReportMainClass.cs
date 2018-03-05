using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;
using System.Collections;
using System.Globalization;
using System.Windows.Forms;

namespace T1.B1.WithholdingTax.Report
{
    public class ReportMainClass
    {
        static private ReportMainClass objReportMainClass = null;
        static private string InternalClassID = B1.WithholdingTax.Report.InteractionId.Default.InternalClassId;

        private ReportMainClass()
        {

        }

        #region Events

        static public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form objForm = null;


            try
            {
                if (objReportMainClass == null)
                    objReportMainClass = new ReportMainClass();

                if (!pVal.BeforeAction)
                {
                    if (pVal.MenuUID == B1.WithholdingTax.Report.InteractionId.Default.mnuWHTReportId)
                    {
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm(B1.WithholdingTax.Report.InteractionId.Default.frmWHTReportFormTypeEx);
                        objFormCreationParams.FormType = B1.WithholdingTax.Report.InteractionId.Default.frmWHTReportFormTypeEx;
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);
                        objForm.Visible = true;
                    }
                    
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, InternalClassID + "MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, InternalClassID + "MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static public void ItemEvent(ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form objForm = null;


            try
            {
                #region ChooseFromList
                if (
                    pVal.FormTypeEx == B1.WithholdingTax.Report.InteractionId.Default.frmWHTReportFormTypeEx
                    
                    && pVal.BeforeAction == false
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    )
                {
                    SAPbouiCOM.ChooseFromListEvent newPval = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    if(pVal.ItemUID == "tWTFrom")
                    {
                        setDSValue("U_WFrom", newPval);
                    }
                    else if (pVal.ItemUID == "tWTTo")
                    {
                        setDSValue("U_WTo", newPval);
                    }
                    else if (pVal.ItemUID == "tTPFrom")
                    {
                        setDSValue("U_TFrom", newPval);
                    }
                    else if (pVal.ItemUID == "tTPTo")
                    {
                        setDSValue("U_TTo", newPval);
                    }
                }
                #endregion ChooseFromList

                if(
                    pVal.FormTypeEx == B1.WithholdingTax.Report.InteractionId.Default.frmWHTReportFormTypeEx
                    && pVal.ItemUID == "Item_15"
                    && !pVal.BeforeAction
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    )
                {


                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, InternalClassID + "ItemEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, InternalClassID + "ItemEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static private void setDSValue(string strDS, SAPbouiCOM.ChooseFromListEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.UserDataSource oDS = null;
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.EditText oEdit = null;
            string strValue = "";

            try
            {
                oDT = pVal.SelectedObjects;
                if (oDT != null)
                {
                    objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                    oEdit = objForm.Items.Item(pVal.ItemUID).Specific;
                    strValue = oEdit.ChooseFromListAlias;
                    oDS = objForm.DataSources.UserDataSources.Item(strDS);
                    oDS.ValueEx = oDT.GetValue(strValue, 0);
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, InternalClassID + "setDSValue", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, InternalClassID + "setDSValue", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        

        #endregion Events



        static public void addMenu()
        {
            string strMenuDescription = "";
            string strMenuId = "";

            if (objReportMainClass == null)
                objReportMainClass = new ReportMainClass();

            try
            {
                #region Withholding Tax Report
                strMenuId = B1.WithholdingTax.Report.InteractionId.Default.mnuWHTReportId;
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = B1.WithholdingTax.Report.MessageStrings.Default.mnuWHTReportDescription;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intTotal = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.WithholdingTax.InteractionId.Default.mnuWHTReportId).SubMenus.Count + 1;
                    objMenuCreationParams.Position = intTotal;


                    BYBB1MainObject.Instance.B1Application.Menus.Item(B1.WithholdingTax.InteractionId.Default.mnuWHTReportId).SubMenus.AddEx(objMenuCreationParams);
                }
                #endregion Withholding Tax Report
                

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, InternalClassID + "addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, InternalClassID + "addMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        static private string localForm(string strFormId)
        {
            string strResult = "";


            try
            {
                if (strFormId == B1.WithholdingTax.Report.InteractionId.Default.frmWHTReportFormTypeEx)
                {
                    strResult = B1.WithholdingTax.Report.Resources.ReportResources.WHTFRM007;

                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, InternalClassID + "localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, InternalClassID +  "localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strResult;

        }



        /*

        static public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;

            try
            {
                if (objReportMainClass == null)
                    objReportMainClass = new WithholdingTax();

                #region getBPWT in Memory
                if (pVal.FormTypeEx == "141" && pVal.ItemUID == "4" && !pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                {
                    bool blDisabled = BYBCache.Instance.getFromCache("Disable_" + pVal.FormUID) != null ? BYBCache.Instance.getFromCache("Disable_" + pVal.FormUID) : false;
                    if (!blDisabled)
                    {
                        getWTCodesForBP(pVal);
                    }
                }
                if (pVal.FormTypeEx == "133" && pVal.ItemUID == "4" && !pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                {
                    bool blDisabled = BYBCache.Instance.getFromCache("Disable_" + pVal.FormUID) != null ? BYBCache.Instance.getFromCache("Disable_" + pVal.FormUID) : false;
                    if (!blDisabled)
                    {
                        getWTCodesForBP(pVal);
                    }
                }

                #endregion getBPWT in Memory

                #region Check WithHoldingTax Utility

                if (pVal.FormTypeEx == "BYB_WHTFRM006")
                {
                    if (pVal.BeforeAction)
                    {

                    }
                    else
                    {

                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                        {
                            if (pVal.ItemUID == "btSearch")
                            {
                                getAllNotRegisteredDocuments(pVal);
                            }
                            else if (pVal.ItemUID == "btFix")
                            {
                                fixAllNotRegisteredDocuments(pVal);
                            }


                        }
                    }

                }



                #endregion

                #region SelfWithholdingTax Definition Window

                if (pVal.FormTypeEx == "BYB_WHTFRM003")
                {
                    if (pVal.BeforeAction)
                    {

                    }
                    else
                    {
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                        {
                            if (pVal.ItemUID == "0_U_G")
                            {
                                setBPNameAfterCFL((SAPbouiCOM.ChooseFromListEvent)pVal);
                            }


                        }
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                        {
                            if (pVal.ItemUID == "btAddAll")
                            {
                                addAllBPs(pVal);
                            }
                        }
                    }

                }


                #endregion SelfWithholdingTax Definition Window


                if (pVal.FormTypeEx == "60504" && pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {

                    bool blDisabled = BYBCache.Instance.getFromCache("Disable_" + pVal.FormUID) != null ? BYBCache.Instance.getFromCache("Disable_" + pVal.FormUID) : false;
                    if (!blDisabled)
                    {
                        SAPbouiCOM.Form oForm = BYBB1MainObject.Instance.B1Application.Forms.ActiveForm;
                        BYBCache.Instance.addToCache("LastActiveForm", oForm.UniqueID, BYBCache.objCachePriority.Default);
                        setBPWT(oForm, pVal);
                        initWTForm(pVal);
                    }

                }

                if (pVal.FormTypeEx == "60504" && pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "1")
                {
                    if (!pVal.InnerEvent)
                    {
                        SAPbouiCOM.Form objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            int intRes = BYBB1MainObject.Instance.B1Application.MessageBox("La modificación manual de las retenciones deshabilitará termporalmente el cálculo automatico por articulo para este documento. Desea Continuar? ", 2, "Sí", "No");
                            if (intRes == 2)
                            {
                                BubbleEvent = false;
                            }
                            else
                            {
                                string strLast = BYBCache.Instance.getFromCache("LastActiveForm");
                                BYBCache.Instance.addToCache("Disable_" + strLast, true, BYBCache.objCachePriority.Default);
                            }
                        }
                    }
                }

                if (pVal.FormTypeEx == "141" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
                {
                    BYBCache.Instance.removeFromCache("Disable_" + pVal.FormUID);
                }




                if ((
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT ||
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                ) && pVal.FormTypeEx == "141" && pVal.ItemUID == "38")
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ColUID == "1" ||
                            pVal.ColUID == "11" ||
                            pVal.ColUID == "14" ||
                            pVal.ColUID == "15" ||
                            pVal.ColUID == "174")
                        {

                            if (BYBB1MainObject.Instance.B1Application.Menus.Exists("5897"))
                            {
                                bool blDisabled = BYBCache.Instance.getFromCache("Disable_" + pVal.FormUID) != null ? BYBCache.Instance.getFromCache("Disable_" + pVal.FormUID) : false;
                                if (!blDisabled)
                                {
                                    BYBCache.Instance.addToCache("WTBeenModified", true, BYBCache.objCachePriority.Default);
                                    BYBCache.Instance.addToCache("WTAddOnGenerated", true, BYBCache.objCachePriority.Default);
                                    BYBB1MainObject.Instance.B1Application.ActivateMenuItem("5897");
                                }
                            }
                        }
                    }
                }

                #region Add Missing Withholding Tax
                if (pVal.FormTypeEx == "BYB_WHTFRM004")
                {
                    if (pVal.BeforeAction)
                    {

                    }
                    else
                    {
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                        {
                            if (pVal.ItemUID == "btnCalc")
                            {
                                addMissingSelfWithHolding(pVal);
                            }
                        }
                        else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
                        {
                            if (pVal.ItemUID == "grdSWT")
                            {
                                selectAllPendingDocuments(pVal);
                            }
                        }
                    }
                }


                #endregion Add Missing Withholding Tax
                if (pVal.FormTypeEx == "BYB_WHTFRM005" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction && pVal.ItemUID == "btnCalc")
                {
                    addTaxandWTAdjust(pVal);
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

        static private void addTaxandWTAdjust(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.UserDataSource oDocEntry = null;
            SAPbouiCOM.UserDataSource oCardCode = null;
            SAPbouiCOM.UserDataSource oDocDate = null;
            SAPbouiCOM.UserDataSource oDocType = null;

            XmlDocument oXML = new XmlDocument();
            SAPbouiCOM.DataTable dtNew = null;
            int intDocEntry = 0;
            string strCardCode = "";
            string strMessage = "";
            SAPbobsCOM.JournalEntries oJournal = null;

            SAPbobsCOM.WithholdingTaxCodes oWT = null;
            SAPbobsCOM.SalesTaxCodes oTax = null;

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oDocEntry = objForm.DataSources.UserDataSources.Item("oDocEntry");
                oCardCode = objForm.DataSources.UserDataSources.Item("oCardCode");
                oDocDate = objForm.DataSources.UserDataSources.Item("oDocDate");
                oDocType = objForm.DataSources.UserDataSources.Item("oDocType");
                intDocEntry = Convert.ToInt32(oDocEntry.ValueEx);
                strCardCode = oCardCode.ValueEx;
                dtNew = objForm.DataSources.DataTables.Item("dtNew");
                oXML.LoadXml(dtNew.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));
                XmlNodeList oNodes = oXML.SelectNodes("DataTable/Rows/Row");
                if (oNodes != null && oNodes.Count > 0)
                {
                    foreach (XmlNode xn in oNodes)
                    {
                        string strOperation = xn.SelectSingleNode("Cells/Cell[1]/Value").InnerText;
                        string strWTCode = xn.SelectSingleNode("Cells/Cell[3]/Value").InnerText;
                        string strTaxCode = xn.SelectSingleNode("Cells/Cell[4]/Value").InnerText;
                        double dbBase = BYBHelpers.Instance.getStandarNumericValue(xn.SelectSingleNode("Cells/Cell[5]/Value").InnerText);
                        double dbValue = BYBHelpers.Instance.getStandarNumericValue(xn.SelectSingleNode("Cells/Cell[6]/Value").InnerText);

                        if (strOperation.Trim().Length > 0 &&
                            (strWTCode.Trim().Length > 0 || strTaxCode.Trim().Length > 0) &&
                            dbBase > 0 &&
                            dbValue > 0
                            )
                        {


                            oJournal = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                            oJournal.Memo = "Ajuste Retenciones e Impuestos";



                            oJournal.TaxDate = new DateTime(Convert.ToInt32(oDocDate.ValueEx.Substring(0,4)), Convert.ToInt32(oDocDate.ValueEx.Substring(4, 2)), Convert.ToInt32(oDocDate.ValueEx.Substring(6, 2)));// Convert.ToDateTime(    );
                            oJournal.ReferenceDate = new DateTime(Convert.ToInt32(oDocDate.ValueEx.Substring(0, 4)), Convert.ToInt32(oDocDate.ValueEx.Substring(4, 2)), Convert.ToInt32(oDocDate.ValueEx.Substring(6, 2)));// Convert.ToDateTime(    );
                            oJournal.DueDate = new DateTime(Convert.ToInt32(oDocDate.ValueEx.Substring(0, 4)), Convert.ToInt32(oDocDate.ValueEx.Substring(4, 2)), Convert.ToInt32(oDocDate.ValueEx.Substring(6, 2)));// Convert.ToDateTime(    );

                            oJournal.Reference = Convert.ToString(oDocEntry.ValueEx);
                            oJournal.Reference3 = Convert.ToString(oDocEntry.ValueEx);

                            #region WithHolding
                            if (strWTCode.Trim().Length > 0)
                            {
                                oWT = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                                if (oWT.GetByKey(strWTCode))
                                {
                                    oJournal.Lines.Credit = dbValue;
                                    if (strOperation.Trim() == "ATO")
                                    {
                                        oJournal.Lines.AccountCode = oWT.Account;
                                    }
                                    else if (strOperation.Trim() == "AJU")
                                    {
                                        oJournal.Lines.ShortName = strCardCode;
                                    }
                                    oJournal.Lines.Add();
                                    oJournal.Lines.SetCurrentLine(1);
                                    oJournal.Lines.Debit = dbValue;
                                    if (strOperation.Trim() == "ATO")
                                    {
                                        oJournal.Lines.ShortName = strCardCode;
                                    }
                                    else if (strOperation.Trim() == "AJU")
                                    {
                                        oJournal.Lines.AccountCode = oWT.Account;
                                    }



                                }
                                else
                                {
                                    strMessage = "No se encontró el código de retenciones seleccionado.";

                                }

                                int intResult = oJournal.Add();
                                if (intResult != 0)
                                {
                                    string strMsg = BYBB1MainObject.Instance.B1Company.GetLastErrorDescription();
                                    Exception er = new Exception(Convert.ToString("COM Error::" + Convert.ToString(BYBB1MainObject.Instance.B1Company.GetLastErrorCode()) + "::" + Convert.ToString(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription())));
                                    BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
                                }
                                else
                                {
                                    string strLastKey = BYBB1MainObject.Instance.B1Company.GetNewObjectKey();

                                    SAPbobsCOM.GeneralService oGeneralService = null;
                                    SAPbobsCOM.GeneralData oGeneralData = null;
                                    SAPbobsCOM.GeneralData oChild = null;
                                    SAPbobsCOM.GeneralDataCollection oChildren = null;
                                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                                    SAPbobsCOM.CompanyService oCompanyService = null;

                                    SAPbobsCOM.Recordset objRecordSet = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string strSQL = B1.WithholdingTax.Resources.dbQueries.sqlGetUDODocEntry;
                                    strSQL = strSQL.Replace("[--CardCode--]", strCardCode);
                                    strSQL = strSQL.Replace("[--DocEntry--]", oDocEntry.ValueEx);

                                    objRecordSet.DoQuery(strSQL);

                                    if (objRecordSet.RecordCount == 1)
                                    {
                                        int intUDODocEntry = objRecordSet.Fields.Item(0).Value;


                                        oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                                        oGeneralService = oCompanyService.GetGeneralService("BYB_T1WHTMOV");

                                        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                        oGeneralParams.SetProperty("DocEntry", intUDODocEntry);


                                        oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                                        oChildren = oGeneralData.Child("BYB_T1WHT401");

                                        #region Build Lines

                                        oChild = oChildren.Add();
                                        oChild.SetProperty("U_Type", "W");
                                        oChild.SetProperty("U_Operation", strOperation);
                                        oChild.SetProperty("U_Source", "D");
                                        oChild.SetProperty("U_Code", strWTCode);
                                        oChild.SetProperty("U_Percent", oWT.Lines.Rate);
                                        oChild.SetProperty("U_BaseAmnt", dbBase);
                                        oChild.SetProperty("U_Value", dbValue);
                                        oChild.SetProperty("U_JE", strLastKey);
                                        #endregion Build Lines


                                        oGeneralService.Update(oGeneralData);
                                        BYBB1MainObject.Instance.B1Application.MessageBox("El ajuste se contabilizó con éxito.");

                                        objForm.Close();

                                    }

                                }



                            }
                            #endregion WithHolding

                            #region TAX
                            if (strTaxCode.Trim().Length > 0)
                            {
                                oTax = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesTaxCodes);
                                if (oTax.GetByKey(strTaxCode))
                                {

                                    SAPbobsCOM.SalesTaxAuthorities oSTAuth = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesTaxAuthorities);
                                    oSTAuth.GetByKey(oTax.Lines.STACode, oTax.Lines.STAType);

                                    string strAccount = oSTAuth.AOrRTaxAccount;

                                    oJournal.Lines.Credit = dbValue;
                                    if (strOperation.Trim() == "ATO")
                                    {
                                        oJournal.Lines.ShortName = strCardCode;
                                    }
                                    else if (strOperation.Trim() == "AJU")
                                    {
                                        oJournal.Lines.AccountCode = strAccount;
                                    }

                                    oJournal.Lines.Add();
                                    oJournal.Lines.SetCurrentLine(1);
                                    oJournal.Lines.Debit = dbValue;
                                    if (strOperation.Trim() == "ATO")
                                    {
                                        oJournal.Lines.AccountCode = strAccount;
                                    }
                                    else if (strOperation.Trim() == "AJU")
                                    {
                                        oJournal.Lines.ShortName = strCardCode;
                                    }
                                }
                                else
                                {
                                    strMessage = "No se encontró el código de impuestos seleccionado.";

                                }

                                int intResult = oJournal.Add();
                                if (intResult != 0)
                                {
                                    string strMsg = BYBB1MainObject.Instance.B1Company.GetLastErrorDescription();
                                    Exception er = new Exception(Convert.ToString("COM Error::" + Convert.ToString(BYBB1MainObject.Instance.B1Company.GetLastErrorCode()) + "::" + Convert.ToString(BYBB1MainObject.Instance.B1Company.GetLastErrorDescription())));
                                    BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
                                }
                                else
                                {
                                    string strLastKey = BYBB1MainObject.Instance.B1Company.GetNewObjectKey();

                                    SAPbobsCOM.GeneralService oGeneralService = null;
                                    SAPbobsCOM.GeneralData oGeneralData = null;
                                    SAPbobsCOM.GeneralData oChild = null;
                                    SAPbobsCOM.GeneralDataCollection oChildren = null;
                                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                                    SAPbobsCOM.CompanyService oCompanyService = null;

                                    SAPbobsCOM.Recordset objRecordSet = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string strSQL = B1.WithholdingTax.Resources.dbQueries.sqlGetUDODocEntry;
                                    strSQL = strSQL.Replace("[--CardCode--]", strCardCode);
                                    strSQL = strSQL.Replace("[--DocEntry--]", oDocEntry.ValueEx);

                                    objRecordSet.DoQuery(strSQL);

                                    if (objRecordSet.RecordCount == 1)
                                    {
                                        int intUDODocEntry = objRecordSet.Fields.Item(0).Value;


                                        oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                                        oGeneralService = oCompanyService.GetGeneralService("BYB_T1WHTMOV");

                                        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                        oGeneralParams.SetProperty("DocEntry", intUDODocEntry);


                                        oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                                        oChildren = oGeneralData.Child("BYB_T1WHT401");

                                        #region Build Lines

                                        oChild = oChildren.Add();
                                        oChild.SetProperty("U_Type", "T");
                                        oChild.SetProperty("U_Operation", strOperation);
                                        oChild.SetProperty("U_Source", "D");
                                        oChild.SetProperty("U_Code", strTaxCode);
                                        oChild.SetProperty("U_Percent", oTax.Rate);
                                        oChild.SetProperty("U_BaseAmnt", dbBase);
                                        oChild.SetProperty("U_Value", dbValue);
                                        oChild.SetProperty("U_JE", strLastKey);
                                        #endregion Build Lines


                                        oGeneralService.Update(oGeneralData);
                                        BYBB1MainObject.Instance.B1Application.MessageBox("El ajuste se contabilizó con éxito.");

                                        objForm.Close();

                                    }

                                }



                            }
                            #endregion TAX








                        }
                    }

                }
                else
                {
                    strMessage = "No se han encontrado operaciones de ajuste para procesar.";
                }
                BYBB1MainObject.Instance.B1Application.MessageBox(strMessage);





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

        static private void setBPWT(SAPbouiCOM.Form objForm, SAPbouiCOM.ItemEvent pVal)
        {
            XmlNode ndBPWT = null;
            SAPbouiCOM.Matrix oMatrix = null;

            XmlDocument xmlDoc = new XmlDocument();
            SAPbouiCOM.Form wtForm = null;

            try
            {
                bool blWTSource = BYBCache.Instance.getFromCache("WTAddOnGenerated") != null ? BYBCache.Instance.getFromCache("WTAddOnGenerated") : false;
                BYBCache.Instance.removeFromCache("WTAddOnGenerated");

                bool blWTReset = BYBCache.Instance.getFromCache("WTReset") != null ? BYBCache.Instance.getFromCache("WTReset") : false;
                BYBCache.Instance.removeFromCache("WTReset");


                wtForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                oMatrix = wtForm.Items.Item("6").Specific;

                
                SAPbouiCOM.EditText oEdit = oMatrix.GetCellSpecific("1", 1);
                string strFirstValue = oEdit.Value.Trim();
                if (strFirstValue.Length > 0)
                {






                    bool blBeenModified = BYBCache.Instance.getFromCache("WTBeenModified") != null ? BYBCache.Instance.getFromCache("WTBeenModified") : false;
                    BYBCache.Instance.removeFromCache("WTBeenModified");
                    if (!blBeenModified && blWTSource)
                    {
                        wtForm.Close();
                    }
                    else if (blBeenModified && blWTSource)
                    {
                        WTCalculation(objForm);
                        Hashtable finalHash = BYBCache.Instance.getFromCache("WTCalc" + objForm.UniqueID) != null ? BYBCache.Instance.getFromCache("WTCalc" + objForm.UniqueID) : new Hashtable();
                        Hashtable hashMatrixFastLocate = BYBCache.Instance.getFromCache("WTMatrixPos" + objForm.UniqueID) != null ? BYBCache.Instance.getFromCache("WTMatrixPos" + objForm.UniqueID) : new Hashtable();
                        foreach (string strKey in hashMatrixFastLocate.Keys)
                        {
                            int intPos = (int)hashMatrixFastLocate[strKey];
                            double dbWT = 0;
                            if (finalHash.ContainsKey(strKey))
                            {
                                ArrayList arr = (ArrayList)finalHash[strKey];

                                dbWT = (double)arr[3];

                            }
                            SAPbouiCOM.EditText oEditVal = oMatrix.GetCellSpecific("7", intPos);

                            //TODO Include formater provider and change array object to class
                            oEditVal.Value = Convert.ToString(dbWT);
                        }
                        if (wtForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            SAPbouiCOM.Item oItem = wtForm.Items.Item("1");
                            oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        }

                        wtForm.Close();






                    }

                }
                else
                {
                    ndBPWT = BYBCache.Instance.getFromCache("BPWT" + objForm.UniqueID);
                    if (ndBPWT != null)
                    {
                        XmlNodeList oNodeList = ndBPWT.SelectNodes("row");
                        if (oNodeList != null)
                        {
                            int intWTCount = oNodeList.Count;
                            if (intWTCount > 0)
                            {
                                oMatrix.AddRow(intWTCount);
                                int intCount = 1;
                                Hashtable hashQuickPosition = new Hashtable();
                                foreach (XmlNode xn in oNodeList)
                                {
                                    XmlElement xm = (XmlElement)xn.SelectSingleNode("WTCode");

                                    SAPbouiCOM.EditText oEditVal = oMatrix.GetCellSpecific("1", intCount);
                                    oEditVal.Value = xm.InnerText;
                                    hashQuickPosition.Add(xm.InnerText, intCount);
                                    intCount++;


                                }
                                BYBCache.Instance.addToCache("WTMatrixPos" + objForm.UniqueID, hashQuickPosition, BYBCache.objCachePriority.Default);
                            }
                        }

                    }

                    if (blWTSource)
                    {
                        if (wtForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            SAPbouiCOM.Item oItem = wtForm.Items.Item("1");
                            oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        }
                        wtForm.Close();
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

        static private void getWTCodesForBP(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                getWTCodesForBP(objForm);

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

        static private void getWTCodesForBP(SAPbouiCOM.Form objForm)
        {

            SAPbobsCOM.BusinessPartners objBP = null;
            XmlDocument objDocument = null;
            XmlNode objWTNode = null;
            string strBP = "";
            string strHashBP = "";

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(objForm.UniqueID);
                if (objForm.TypeEx == "133")
                {
                    strBP = objForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim();
                }
                else
                {
                    strBP = objForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim();
                }

                if (CardCodeHash.ContainsKey(objForm.UniqueID))
                {
                    strHashBP = (string)CardCodeHash[objForm.UniqueID];
                }

                if (strHashBP != strBP)
                {
                    BYBCache.Instance.addToCache("WTReset", true, BYBCache.objCachePriority.Default);
                    objBP = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                    if (objBP.GetByKey(strBP))
                    {
                        objDocument = new XmlDocument();
                        objDocument.LoadXml(objBP.GetAsXML());
                        objWTNode = objDocument.SelectSingleNode("/BOM/BO/BPWithholdingTax");
                        if (objWTNode != null)
                        {
                            BYBCache.Instance.addToCache("BPWT" + objForm.UniqueID, objWTNode, BYBCache.objCachePriority.Default);
                            if (CardCodeHash.ContainsKey(objForm.UniqueID))
                            {
                                CardCodeHash[objForm.UniqueID] = strBP;
                            }
                            else
                            {
                                CardCodeHash.Add(objForm.UniqueID, strBP);
                            }

                        }
                    }
                }
                BYBCache.Instance.addToCache("WTAddOnGenerated", true, BYBCache.objCachePriority.Default);
                BYBB1MainObject.Instance.B1Application.ActivateMenuItem("5897");









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

        static public void getMoneyFromForm(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            Hashtable hashInfo = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            string strCode = "";
            bool blShowOnRight = false;

            try
            {
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                hashInfo = (Hashtable)BYBCache.Instance.getFromCache("MoneyCodes");
                objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                SAPbobsCOM.AdminInfo objAdminInfo = objCompanyService.GetAdminInfo();
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == "70"))
                {
                    if (objAdminInfo.DisplayCurrencyontheRight == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        blShowOnRight = true;
                    }

                    SAPbouiCOM.ComboBox objCombo = oForm.Items.Item("70").Specific;
                    string strValue = objCombo.Selected.Value;
                    if (strValue == "L")
                    {
                        strCode = objAdminInfo.LocalCurrency;
                    }
                    else if (strValue == "S")
                    {
                        strCode = objAdminInfo.SystemCurrency;
                    }
                    else
                    {
                        SAPbouiCOM.ComboBox objComboTemp = oForm.Items.Item("63").Specific;
                        strCode = objComboTemp.Selected.Value;
                    }
                }
                else
                {
                    if (oForm.Items.Item("63").Visible)
                    {
                        SAPbouiCOM.ComboBox objComboTemp = oForm.Items.Item("63").Specific;
                        strCode = objComboTemp.Selected.Value;
                    }
                    else
                    {
                        SAPbouiCOM.ComboBox objCombo = oForm.Items.Item("70").Specific;
                        string strValue = objCombo.Selected.Value;
                        if (strValue == "L")
                        {
                            strCode = objAdminInfo.LocalCurrency;
                        }
                        else if (strValue == "S")
                        {
                            strCode = objAdminInfo.SystemCurrency;
                        }
                    }
                }

                BYBCache.Instance.addToCache("Curr" + pVal.FormUID, strCode, BYBCache.objCachePriority.Default);
                BYBCache.Instance.addToCache("ShowRight" + pVal.FormUID, blShowOnRight, BYBCache.objCachePriority.Default);


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

        static public void initWTForm(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;


            try
            {
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                oItem = oForm.Items.Add("BYBDetail", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = oForm.Items.Item("2").Left + 10 + oForm.Items.Item("2").Width;
                oItem.Top = oForm.Items.Item("2").Top;
                oItem.Width = oForm.Items.Item("2").Width;
                oButton = oItem.Specific;
                oButton.Caption = "Detalles";


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

        static private void openWTWindow(SAPbouiCOM.Form baseForm)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.FormCreationParams oParams = null;
            Hashtable hashWTCalculation = null;
            SAPbouiCOM.DataTable objDT = null;

            try
            {
                hashWTCalculation = new Hashtable();
                hashWTCalculation = BYBCache.Instance.getFromCache("WTCalc" + baseForm.UniqueID);

                //TODO get infomation based on origin of open

                oParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                oParams.XmlData = B1.WithholdingTax.Resources.WithholdingTax.WHTFRM002;
                oParams.FormType = "BYB_T1WTF002";
                oParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(oParams);
                if (hashWTCalculation.Count > 0)
                {
                    objDT = objForm.DataSources.DataTables.Item("DT_WHT");
                    objDT.Rows.Add(hashWTCalculation.Count);
                    int intLine = 0;
                    foreach (string strKey in hashWTCalculation.Keys)
                    {
                        ArrayList objArray = (ArrayList)hashWTCalculation[strKey];

                        objDT.SetValue("C_WTCode", intLine, strKey);
                        objDT.SetValue("C_Name", intLine, objArray[0]);
                        objDT.SetValue("C_Base", intLine, objArray[1]);
                        objDT.SetValue("C_WHBase", intLine, objArray[2]);
                        objDT.SetValue("C_WHTax", intLine, objArray[3]);
                        objDT.SetValue("C_Acct", intLine, objArray[4]);
                        intLine++;
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

        

        

        static public void UDOEvent(ref SAPbouiCOM.UDOEvent udoEventArgs, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (udoEventArgs.EventType == SAPbouiCOM.BoEventTypes.et_UDO_FORM_OPEN && udoEventArgs.UDOCode == "BYB_T1SWTU001")
                {
                    string strForm = localForm("BYB_WHTFRM003");
                    udoEventArgs.FormSrf = strForm;

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
            if (objReportMainClass == null)
            {
                objReportMainClass = new WithholdingTax();
            }

            BubbleEvent = true;

            SAPbouiCOM.Form objForm = null;

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                if (objForm.TypeEx == "bybwhtf001")
                {

                    if (eventInfo.BeforeAction && eventInfo.ItemUID == "0_U_G")
                    {
                        addRowMenu(objForm);
                    }
                    else
                    {
                        removeRowMenu();
                    }

                }
                else if (objForm.TypeEx == "133" || objForm.TypeEx == "392")
                {
                    if (eventInfo.BeforeAction)
                    {
                        BYBCache.Instance.addToCache("LastWTInvoice", objForm, BYBCache.objCachePriority.Default);
                        addWTAdjustmentMenu(objForm);
                    }
                    else
                    {
                        removeWTAdjustmentMenu();
                    }
                }
                else if (objForm.TypeEx == "BYB_WHTFRM005")
                {
                    if (eventInfo.BeforeAction && eventInfo.ItemUID == "newGrid")
                    {

                        addWTAdjustmentLine(objForm);
                    }
                    else
                    {
                        removeWTAdjustmentLine();
                    }
                }
                #region SelfWitholdintTax Window Definition
                else if (objForm.TypeEx == "BYB_WHTFRM003")
                {
                    if (eventInfo.BeforeAction && eventInfo.ItemUID == "0_U_G")
                    {

                        addLineMenuBPMatrixAdd(objForm);
                        deleteLineMenuBPMatrixAdd(objForm);
                        BYBCache.Instance.addToCache(eventInfo.FormUID + "_LastRow", eventInfo.Row, BYBCache.objCachePriority.Default);
                        BYBCache.Instance.addToCache("LastRightClickForm", eventInfo.FormUID, BYBCache.objCachePriority.Default);


                    }
                    else
                    {
                        addLineMenuBPMatrixRemove();
                        deleteLineMenuBPMatrixRemove();
                    }
                }
                #endregion SelfWitholdintTax Window Definition



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

        static private void addRowMenu(SAPbouiCOM.Form objForm)
        {
            string strMenuDescription = "";
            string strMenuId = "";
            try
            {
                strMenuId = "bybwtm002";
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
                strMenuId = "bybwtm002";
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

        static private void addWTAdjustmentLine(SAPbouiCOM.Form objForm)
        {
            string strMenuDescription = "";
            string strMenuId = "";
            try
            {
                strMenuId = "bybwtm004";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Añadir linea";
                    //strMenuId = T1.B1.FixedAssets.InteractionId.Default.fixedAssetsAddRowMenuId;

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Añadir Línea";
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                    objForm.Menu.AddEx(objMenuCreationParams);

                    ///TODO Eliminar Linea

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

        static private void removeWTAdjustmentLine()
        {
            string strMenuId = "";

            try
            {
                strMenuId = "bybwtm004";
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

        static private void addWTAdjustmentMenu(SAPbouiCOM.Form objForm)
        {
            string strMenuId = "";
            try
            {
                strMenuId = "bybwtm003";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Ajustar Impuestos y Retenciones";
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    objMenuCreationParams.Position = objForm.Menu.Count + 1;

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

        static private void removeWTAdjustmentMenu()
        {
            string strMenuId = "";

            try
            {
                strMenuId = "bybwtm003";
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

        static private void WTCalculation(SAPbouiCOM.Form objForm)
        {

            XmlDocument objAllWTCodes = null;
            XmlNode objBPCodes = null;
            XmlDocument objXMLMatrix = null;
            XmlDocument objXMLWHTDefinition = null;
            SAPbouiCOM.Matrix objMatrix;

            string strWTTypes = "";
            SAPbouiCOM.ComboBox objCombo = null;

            SAPbobsCOM.Recordset objRec = null;

            Hashtable hashWT = null;



            try
            {
                if (objForm.TypeEx == "141")
                {
                    strWTTypes = "P";
                }
                else
                {
                    strWTTypes = "S";
                }

                objCombo = objForm.Items.Item("3").Specific;
                string strDropValue = objCombo.Value;
                //Only run when on Item matrix of the form
                if (strDropValue == "I")
                {

                    ///TODO handle manual modifications of WT

                    //Get all WT by Item for comparison from Cache
                    objAllWTCodes = BYBCache.Instance.getFromCache("WTCodesByItem");

                    //Namespace to read from the recordset
                    var nsm = new XmlNamespaceManager(objAllWTCodes.NameTable);
                    nsm.AddNamespace("s", "http://www.sap.com/SBO/SDK/DI");

                    objBPCodes = BYBCache.Instance.getFromCache("BPWT" + objForm.UniqueID);
                    if (objBPCodes == null)
                    {
                        getWTCodesForBP(objForm);
                        objBPCodes = BYBCache.Instance.getFromCache("BPWT" + objForm.UniqueID);
                    }


                    XmlNodeList bpwht = objBPCodes.SelectNodes("//BPWithholdingTax/row/WTCode");
                    if (bpwht.Count > 0)
                    {


                        objMatrix = objForm.Items.Item("38").Specific;
                        objXMLMatrix = new XmlDocument();
                        objXMLMatrix.LoadXml(objMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All));
                        XmlNodeList objNodes = objXMLMatrix.SelectNodes("/Matrix/Rows/Row[./Visible = '1']");


                        hashWT = new Hashtable();


                        if (objNodes.Count > 0)
                        {
                            foreach (XmlNode rowNode in objNodes)
                            {
                                string strItemCode = rowNode.SelectSingleNode("Columns/Column[./ID = '1']/Value").InnerText;
                                string strPrice = rowNode.SelectSingleNode("Columns/Column[./ID = '21']/Value").InnerText;
                                string strWTLiable = rowNode.SelectSingleNode("Columns/Column[./ID = '174']/Value").InnerText;

                                double dblPrice = 0;
                                if (strPrice.Length > 0)
                                {
                                    BYBHelpers.Instance.getDoubleFromMoney(strPrice);
                                    dblPrice = BYBHelpers.Instance.getDoubleFromMoney(strPrice);
                                }

                                if (strItemCode.Length > 0)
                                {
                                    if (strWTLiable == "Y")
                                    {
                                        if (dblPrice > 0)
                                        {
                                            XmlNodeList objWTForItem = objAllWTCodes.SelectNodes("/s:Recordset/s:Rows/s:Row[./s:Fields/s:Field/s:Value = '" + strItemCode + "' and ./s:Fields/s:Field/s:Value = '" + strWTTypes + "' ]/s:Fields/s:Field[./s:Alias = 'U_Item']/s:Value", nsm);
                                            if (objWTForItem.Count > 0)
                                            {
                                                foreach (XmlNode objWTTax in objWTForItem)
                                                {
                                                    string strWTTax = objWTTax.InnerText;
                                                    if (!hashWT.ContainsKey(strWTTax))
                                                    {
                                                        hashWT.Add(strWTTax, dblPrice);
                                                    }
                                                    else
                                                    {
                                                        double dbActual = (double)hashWT[strWTTax];
                                                        dbActual += dblPrice;
                                                        hashWT[strWTTax] = dbActual;
                                                    }
                                                }
                                            }

                                        }
                                    }
                                }
                            }

                            if (hashWT.Count > 0)
                            {


                                objXMLWHTDefinition = BYBCache.Instance.getFromCache("WTCodesDefinitions");
                                double totalWT = 0;
                                Hashtable finalWTHash = new Hashtable();


                                foreach (string strKey in hashWT.Keys)
                                {
                                    ArrayList arrInfo = new ArrayList();

                                    XmlNode objNode = objXMLWHTDefinition.SelectSingleNode("/s:Recordset/s:Rows/s:Row[./s:Fields/s:Field/s:Value = '" + strKey + "']/s:Fields", nsm);
                                    if (objNode != null)
                                    {
                                        string strName = objNode.SelectSingleNode("s:Field[s:Alias = 'WTName']/s:Value", nsm).InnerText;
                                        double dblRate = BYBHelpers.Instance.getStandarNumericValue(objNode.SelectSingleNode("s:Field[s:Alias = 'PrctBsAmnt']/s:Value", nsm).InnerText);
                                        double dbPercent = BYBHelpers.Instance.getStandarNumericValue(objNode.SelectSingleNode("s:Field[s:Alias = 'Rate']/s:Value", nsm).InnerText);
                                        //TODO Check how this work with segmentation
                                        string strAccount = objNode.SelectSingleNode("s:Field[s:Alias = 'Account']/s:Value", nsm).InnerText;
                                        double dbValue = (double)hashWT[strKey];
                                        double wtValue = (dbValue * (dbPercent / 100)) * (dblRate / 100);
                                        totalWT += wtValue;
                                        arrInfo.Add(strName);
                                        arrInfo.Add(dbValue);
                                        arrInfo.Add(dbValue * (dbPercent / 100));
                                        arrInfo.Add(wtValue);
                                        arrInfo.Add(strAccount);
                                        finalWTHash.Add(strKey, arrInfo);
                                    }

                                }

                                if (totalWT > 0)
                                {


                                    BYBCache.Instance.addToCache("WTCalc" + objForm.UniqueID, finalWTHash, BYBCache.objCachePriority.Default);


                                }



                            }





                        }
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
            objForm.Freeze(false);
        }

        static public void formDataAddEvent(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool blBubbleEvent)
        {
            blBubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.ActionSuccess && BusinessObjectInfo.FormTypeEx == "133" && !BusinessObjectInfo.BeforeAction)
                {
                    addSelfWithHolding(BusinessObjectInfo);
                }


                //BYBB1MainObject.Instance.B1Application.ActivateMenuItem("5897");
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

        static private void addSelfWithHolding(SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbobsCOM.Recordset objRecordset = null;
            string strSql = "";


            SAPbobsCOM.JournalEntries objJE = null;

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oSelfWithHoldingService = null;
            SAPbobsCOM.GeneralData oSelfHolding = null;
            SAPbobsCOM.GeneralDataParams oSelfHoldingParams = null;


            SAPbobsCOM.Documents oDoc = null;

            string strObjectType = "";
            int intObjectType = 0;
            int intObjectentry = 0;
            double dbBaseAmnt = 0;
            string strCardCode = "";
            string strThirdParty = "";



            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                strObjectType = pVal.Type;
                if (strObjectType == "13")
                {
                    oDoc = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                    if (oDoc.Browser.GetByKeys(pVal.ObjectKey))
                    {
                        intObjectType = Convert.ToInt32(oDoc.DocObjectCodeEx);
                        intObjectentry = oDoc.DocEntry;
                        dbBaseAmnt = oDoc.BaseAmount;
                        strCardCode = oDoc.CardCode;

                        objRecordset = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        strSql = "select U_DebitAcct, U_CreditAccount, U_Percent, Code from [@BYB_T1SWT100] where U_Enabled='Y'";
                        objRecordset.DoQuery(strSql);


                        if (objRecordset.RecordCount > 0)
                        {
                            oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                            objRecordset.MoveFirst();
                            while (!objRecordset.EoF)
                            {
                                objJE = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                                string strCode = objRecordset.Fields.Item("Code").Value;
                                double dbPercent = objRecordset.Fields.Item("U_Percent").Value;
                                string strDebitAccount = objRecordset.Fields.Item("U_DebitAcct").Value;
                                string strCreditAccount = objRecordset.Fields.Item("U_CreditAccount").Value;


                                double dbValue = dbBaseAmnt * (dbPercent / 100);
                                objJE.Memo = "strCode";
                                objJE.Lines.Credit = dbValue;
                                objJE.Lines.AccountCode = strCreditAccount;
                                objJE.Lines.Add();
                                objJE.Lines.SetCurrentLine(1);
                                objJE.Lines.Debit = dbValue;
                                objJE.Lines.AccountCode = strDebitAccount;

                                if (objJE.Add() == 0)
                                {
                                    string strValue = BYBB1MainObject.Instance.B1Company.GetNewObjectKey();


                                    oSelfWithHoldingService = oCompanyService.GetGeneralService("BYB_T1SWTU002");
                                    oSelfHolding = oSelfWithHoldingService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                                    oSelfHolding.SetProperty("U_JEEntry", Convert.ToInt32(strValue));
                                    oSelfHolding.SetProperty("U_BaseAmnt", dbBaseAmnt);
                                    oSelfHolding.SetProperty("U_DocType", intObjectType);
                                    oSelfHolding.SetProperty("U_DocEntry", intObjectentry);
                                    oSelfHolding.SetProperty("U_CardCode", strCardCode);
                                    oSelfHolding.SetProperty("U_SWTCode", strCode);
                                    oSelfHolding.SetProperty("U_Total", dbValue);
                                    oSelfHolding.SetProperty("U_DocNum", oDoc.DocNum);
                                    oSelfHolding.SetProperty("U_DocSeries", Convert.ToString(oDoc.Series));
                                    oSelfHolding.SetProperty("LicTradNum", oDoc.FederalTaxID);
                                    oSelfHoldingParams = oSelfWithHoldingService.Add(oSelfHolding);

                                    int strCodeA = oSelfHoldingParams.GetProperty("DocEntry");

                                }




                                objRecordset.MoveNext();
                            }




                        }



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

        static private void getWithholdingInfo(SAPbouiCOM.Form objForm)
        {
            SAPbouiCOM.Grid objGridCurrent = null;
            SAPbouiCOM.Grid objGridNew = null;
            SAPbouiCOM.DataTable objDTCurrent = null;
            SAPbouiCOM.DataTable objDTNew = null;
            SAPbouiCOM.EditTextColumn oEditTExt = null;
            SAPbouiCOM.GridColumn objGridColumn = null;
            SAPbouiCOM.UserDataSources oUDSApplyNow = null;
            SAPbobsCOM.Documents oInvoice = null;
            SAPbobsCOM.JournalEntries oJournal = null;
            string strQuery = "";
            SAPbouiCOM.Form objInvoiceForm = null;

            try
            {
                objInvoiceForm = (SAPbouiCOM.Form)BYBCache.Instance.getFromCache("LastWTInvoice");

                SAPbouiCOM.UserDataSource oDocEntry = objForm.DataSources.UserDataSources.Add("oDocEntry", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10);
                SAPbouiCOM.UserDataSource oCardCode = objForm.DataSources.UserDataSources.Add("oCardCode", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 20);
                SAPbouiCOM.UserDataSource oDocDate = objForm.DataSources.UserDataSources.Add("oDocDate", SAPbouiCOM.BoDataType.dt_DATE, 20);
                SAPbouiCOM.UserDataSource oDocType = objForm.DataSources.UserDataSources.Add("oDocType", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 20);


                if (objInvoiceForm.BusinessObject.Type == "30")
                {
                    oJournal = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    if (oJournal.Browser.GetByKeys(objInvoiceForm.BusinessObject.Key))
                    {
                        oDocEntry.ValueEx = Convert.ToString(oJournal.JdtNum);
                        oCardCode.ValueEx = Convert.ToString(oJournal.TransactionCode);
                        oDocDate.ValueEx = oJournal.ReferenceDate.ToString("yyyyMMdd");
                        oDocType.ValueEx = objInvoiceForm.BusinessObject.Type;


                        objDTCurrent = objForm.DataSources.DataTables.Item("dtCurr");
                        //strQuery = B1.WithholdingTax.Resources.dbQueries.sqlGetCurrentWTandTax;
                        strQuery = B1.WithholdingTax.Resources.dbQueries.sqlGetCurrentInfoFromUDO;
                        strQuery = strQuery.Replace("[--DocEntry--]", Convert.ToString(oJournal.JdtNum));
                        objDTCurrent.ExecuteQuery(strQuery);

                        objGridCurrent = objForm.Items.Item("currGrid").Specific;

                        objGridColumn = objGridCurrent.Columns.Item(0);
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

                        SAPbouiCOM.ComboBoxColumn oCombo = (SAPbouiCOM.ComboBoxColumn)objGridCurrent.Columns.Item(0);

                        oCombo.ValidValues.Add("WT", "Retención");
                        oCombo.ValidValues.Add("TAX", "Impuesto");
                        oCombo.ValidValues.Add("ATO", "Anulación Total");
                        oCombo.ValidValues.Add("AJU", "Ajuste");

                        #region NewGrid


                        objDTNew = objForm.DataSources.DataTables.Item("dtNew");
                        objGridNew = objForm.Items.Item("newGrid").Specific;


                        objGridColumn = objGridNew.Columns.Item(0);
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                        SAPbouiCOM.ComboBoxColumn oNewCombo = (SAPbouiCOM.ComboBoxColumn)objGridNew.Columns.Item(0);

                        oNewCombo.ValidValues.Add("ATO", "Anulación Total");
                        oNewCombo.ValidValues.Add("AJU", "Ajuste");

                        objGridColumn = objGridNew.Columns.Item(1);
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                        SAPbouiCOM.ComboBoxColumn oNewOriginCombo = (SAPbouiCOM.ComboBoxColumn)objGridNew.Columns.Item(1);
                        oNewOriginCombo.ValidValues.Add("DOC", "Documento");
                        oNewOriginCombo.ValidValues.Add("GAS", "Gastos Adicionales");
                        ////TODO Agregar por linea
                        ///TODO Agregar en cambio poner el valor para no buscar

                        SAPbouiCOM.DBDataSource oDSWT = objForm.DataSources.DBDataSources.Add("OWHT");
                        oDSWT.Query();
                        objGridColumn = objGridNew.Columns.Item(2);
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                        SAPbouiCOM.ComboBoxColumn oNewWTCombo = (SAPbouiCOM.ComboBoxColumn)objGridNew.Columns.Item(2);
                        for (int i = 0; i < oDSWT.Size; i++)
                        {
                            oDSWT.Offset = i;
                            oNewWTCombo.ValidValues.Add(oDSWT.GetValue("WTCode", i), oDSWT.GetValue("WTCode", i));
                        }

                        SAPbouiCOM.DBDataSource oDSTax = objForm.DataSources.DBDataSources.Add("OSTC");
                        oDSTax.Query();
                        objGridColumn = objGridNew.Columns.Item(3);
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                        SAPbouiCOM.ComboBoxColumn oNewTaxCombo = (SAPbouiCOM.ComboBoxColumn)objGridNew.Columns.Item(3);
                        for (int i = 0; i < oDSTax.Size; i++)
                        {
                            oDSTax.Offset = i;
                            oNewTaxCombo.ValidValues.Add(oDSTax.GetValue("Code", i), oDSTax.GetValue("Code", i));
                        }

                        objDTNew.Rows.Add();

                        #endregion NewGrid







                        //objGridColumn.Editable = true;


                        //objGridColumn = objGrid.Columns.Item(1);
                        //objGridColumn.Editable = false;
                        //objGridColumn.RightJustified = true;



                        //objGridColumn = objGrid.Columns.Item(2);
                        //objGridColumn.Editable = false;
                        //objGridColumn.RightJustified = true;

                        //objGridColumn = objGrid.Columns.Item(3);
                        //objGridColumn.Editable = false;
                        //objGridColumn.RightJustified = true;

                        //objGridColumn = objGrid.Columns.Item(4);
                        //objGridColumn.Editable = false;
                        //objGridColumn.RightJustified = true;

                        objGridCurrent.AutoResizeColumns();



                    }
                    else
                    {

                        oInvoice = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        if (oInvoice.Browser.GetByKeys(objInvoiceForm.BusinessObject.Key))
                        {




                            oDocEntry.ValueEx = Convert.ToString(oInvoice.DocEntry);
                            oCardCode.ValueEx = Convert.ToString(oInvoice.CardCode);
                            oDocDate.ValueEx = oInvoice.DocDate.ToString("yyyyMMdd");
                            oDocType.ValueEx = objInvoiceForm.BusinessObject.Type;




                            objDTCurrent = objForm.DataSources.DataTables.Item("dtCurr");
                            //strQuery = B1.WithholdingTax.Resources.dbQueries.sqlGetCurrentWTandTax;
                            strQuery = B1.WithholdingTax.Resources.dbQueries.sqlGetCurrentInfoFromUDO;
                            strQuery = strQuery.Replace("[--DocEntry--]", Convert.ToString(oInvoice.DocEntry));
                            objDTCurrent.ExecuteQuery(strQuery);

                            objGridCurrent = objForm.Items.Item("currGrid").Specific;

                            objGridColumn = objGridCurrent.Columns.Item(0);
                            objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

                            SAPbouiCOM.ComboBoxColumn oCombo = (SAPbouiCOM.ComboBoxColumn)objGridCurrent.Columns.Item(0);

                            oCombo.ValidValues.Add("WT", "Retención");
                            oCombo.ValidValues.Add("TAX", "Impuesto");
                            oCombo.ValidValues.Add("ATO", "Anulación Total");
                            oCombo.ValidValues.Add("AJU", "Ajuste");

                            #region NewGrid


                            objDTNew = objForm.DataSources.DataTables.Item("dtNew");
                            objGridNew = objForm.Items.Item("newGrid").Specific;


                            objGridColumn = objGridNew.Columns.Item(0);
                            objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            SAPbouiCOM.ComboBoxColumn oNewCombo = (SAPbouiCOM.ComboBoxColumn)objGridNew.Columns.Item(0);

                            oNewCombo.ValidValues.Add("ATO", "Anulación Total");
                            oNewCombo.ValidValues.Add("AJU", "Ajuste");

                            objGridColumn = objGridNew.Columns.Item(1);
                            objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            SAPbouiCOM.ComboBoxColumn oNewOriginCombo = (SAPbouiCOM.ComboBoxColumn)objGridNew.Columns.Item(1);
                            oNewOriginCombo.ValidValues.Add("DOC", "Documento");
                            oNewOriginCombo.ValidValues.Add("GAS", "Gastos Adicionales");
                            ////TODO Agregar por linea
                            ///TODO Agregar en cambio poner el valor para no buscar

                            SAPbouiCOM.DBDataSource oDSWT = objForm.DataSources.DBDataSources.Add("OWHT");
                            oDSWT.Query();
                            objGridColumn = objGridNew.Columns.Item(2);
                            objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            SAPbouiCOM.ComboBoxColumn oNewWTCombo = (SAPbouiCOM.ComboBoxColumn)objGridNew.Columns.Item(2);
                            for (int i = 0; i < oDSWT.Size; i++)
                            {
                                oDSWT.Offset = i;
                                oNewWTCombo.ValidValues.Add(oDSWT.GetValue("WTCode", i), oDSWT.GetValue("WTCode", i));
                            }

                            SAPbouiCOM.DBDataSource oDSTax = objForm.DataSources.DBDataSources.Add("OSTC");
                            oDSTax.Query();
                            objGridColumn = objGridNew.Columns.Item(3);
                            objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            SAPbouiCOM.ComboBoxColumn oNewTaxCombo = (SAPbouiCOM.ComboBoxColumn)objGridNew.Columns.Item(3);
                            for (int i = 0; i < oDSTax.Size; i++)
                            {
                                oDSTax.Offset = i;
                                oNewTaxCombo.ValidValues.Add(oDSTax.GetValue("Code", i), oDSTax.GetValue("Code", i));
                            }

                            objDTNew.Rows.Add();

                            #endregion NewGrid







                            //objGridColumn.Editable = true;


                            //objGridColumn = objGrid.Columns.Item(1);
                            //objGridColumn.Editable = false;
                            //objGridColumn.RightJustified = true;



                            //objGridColumn = objGrid.Columns.Item(2);
                            //objGridColumn.Editable = false;
                            //objGridColumn.RightJustified = true;

                            //objGridColumn = objGrid.Columns.Item(3);
                            //objGridColumn.Editable = false;
                            //objGridColumn.RightJustified = true;

                            //objGridColumn = objGrid.Columns.Item(4);
                            //objGridColumn.Editable = false;
                            //objGridColumn.RightJustified = true;

                            objGridCurrent.AutoResizeColumns();

                        }


                    }
                    BYBCache.Instance.removeFromCache("LastWTInvoice");
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

        #region AutoRetencion
        static private void addAllBPs(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbobsCOM.Recordset oRecordSet = null;
            SAPbouiCOM.DBDataSource oDS = null;
            string strSQL = "";

            try
            {
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (
                    oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE ||
                    oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE ||
                    oForm.Mode == SAPbouiCOM.BoFormMode.fm_EDIT_MODE
                    )
                {

                    oMatrix = oForm.Items.Item("0_U_G").Specific;
                    oDS = oForm.DataSources.DBDataSources.Item("@BYB_T1SWT101");
                    strSQL = B1.WithholdingTax.Resources.dbQueries.sqlGetAllClients;
                    oRecordSet = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery(strSQL);
                    if (oRecordSet.RecordCount > 0)
                    {
                        oRecordSet.MoveFirst();
                        oDS.Clear();
                        while (!oRecordSet.EoF)
                        {
                            oDS.InsertRecord(0);

                            oDS.SetValue("U_CardCode", 0, oRecordSet.Fields.Item("CardCode").Value);
                            oDS.SetValue("U_CardName", 0, oRecordSet.Fields.Item("CardName").Value);

                            oRecordSet.MoveNext();
                        }
                        oMatrix.LoadFromDataSourceEx(true);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

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

        static private void filterAccountChooseFromList(SAPbouiCOM.Form objForm)
        {
            SAPbouiCOM.ChooseFromList oChooseFL = null;
            SAPbouiCOM.Conditions oConditions = null;
            SAPbouiCOM.Condition oCondition = null;
            try
            {
                oConditions = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                oCondition = oConditions.Add();
                oCondition.Alias = "Postable";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "Y";


                oChooseFL = objForm.ChooseFromLists.Item("CFL_2");
                oChooseFL.SetConditions(oConditions);

                oChooseFL = objForm.ChooseFromLists.Item("CFL_3");
                oChooseFL.SetConditions(oConditions);


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

        static private void addFirstLineToMatrixSWT(SAPbouiCOM.Form objForm)
        {

            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDS = null;
            try
            {
                oMatrix = objForm.Items.Item("0_U_G").Specific;
                oDS = objForm.DataSources.DBDataSources.Item("@BYB_T1SWT101");
                oDS.InsertRecord(1);
                if (oDS.Size > 1)
                {
                    oDS.RemoveRecord(1);
                }
                oMatrix.LoadFromDataSourceEx(false);

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

        static private void addLineMenuBPMatrixAdd(SAPbouiCOM.Form objForm)
        {
            string strMenuDescription = "";
            string strMenuId = "";
            try
            {
                strMenuId = "BYBSWTM80";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Añadir linea";
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Añadir Línea";
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    objMenuCreationParams.Position = objForm.Menu.Count + 1;
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

        static private void addLineMenuBPMatrixRemove()
        {
            string strMenuId = "";

            try
            {
                strMenuId = "BYBSWTM80";
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

        static private void deleteLineMenuBPMatrixAdd(SAPbouiCOM.Form objForm)
        {
            string strMenuDescription = "";
            string strMenuId = "";
            try
            {
                strMenuId = "BYBSWTM81";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Eliminar linea";
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = "Eliminar Línea";
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    objMenuCreationParams.Position = objForm.Menu.Count + 1;
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

        static private void deleteLineMenuBPMatrixRemove()
        {
            string strMenuId = "";

            try
            {
                strMenuId = "BYBSWTM81";
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

        static private void addNewLineSWTBP()
        {

            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDS = null;
            SAPbouiCOM.Form objForm = null;
            try
            {
                string strFormUID = BYBCache.Instance.getFromCache("LastRightClickForm");
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(strFormUID);

                oMatrix = objForm.Items.Item("0_U_G").Specific;
                oDS = objForm.DataSources.DBDataSources.Item("@BYB_T1SWT101");
                bool blRebuild = false;
                if (oDS.Size == 0)
                {
                    oDS.InsertRecord(0);
                    oDS.InsertRecord(1);
                    if (oDS.Size > 1)
                    {
                        oDS.RemoveRecord(1);
                    }

                    blRebuild = true;

                }
                else
                {
                    oDS.InsertRecord(oMatrix.RowCount);

                }

                oMatrix.LoadFromDataSourceEx(blRebuild);

                oMatrix.SetCellFocus(oMatrix.RowCount, 1);
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                BYBCache.Instance.removeFromCache("LastRightClickForm");

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

        static private void removeLineSWTBP()
        {
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.DBDataSource oDS = null;
            SAPbouiCOM.Form objForm = null;
            try
            {
                string strFormUID = BYBCache.Instance.getFromCache("LastRightClickForm");
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(strFormUID);
                oMatrix = objForm.Items.Item("0_U_G").Specific;
                oDS = objForm.DataSources.DBDataSources.Item("@BYB_T1SWT101");
                int intRow = BYBCache.Instance.getFromCache(objForm.UniqueID + "_LastRow") - 1;
                oMatrix.ClearSelections();
                objForm.Items.Item("0_U_E").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oDS.RemoveRecord(intRow);
                //oDS.InsertRecord(oMatrix.RowCount);
                oMatrix.LoadFromDataSourceEx(false);
                //oMatrix.SetCellFocus(oMatrix.RowCount, 1);
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                //objForm = BYBB1MainObject.Instance.B1Application.Forms.ActiveForm;
                //oMatrix = objForm.Items.Item("0_U_G").Specific;
                // oMatrix.DeleteRow((int)BYBCache.Instance.getFromCache(objForm.UniqueID + "_LastRow"));
                BYBCache.Instance.removeFromCache(objForm.UniqueID + "_LastRow");
                BYBCache.Instance.removeFromCache("LastRightClickForm");
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

        static private void addMissingSelfWithHolding(SAPbouiCOM.ItemEvent pVal)
        {

            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.UserDataSource oUDSDate = null;
            string strDateValue = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList objNodes = null;

            SAPbobsCOM.Recordset objDocuments = null;
            string strSqlDocuments = "";

            SAPbobsCOM.Recordset objRecordset = null;
            string strSql = "";

            SAPbobsCOM.SBObob oSBOBob = null;

            SAPbobsCOM.JournalEntries objJE = null;

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oSelfWithHoldingService = null;
            SAPbobsCOM.GeneralData oSelfHolding = null;
            SAPbobsCOM.GeneralDataParams oSelfHoldingParams = null;


            SAPbobsCOM.Documents oDoc = null;


            int intObjectType = 0;
            int intObjectentry = 0;
            double dbBaseAmnt = 0;
            string strCardCode = "";
            int intTotalDone = 0;
            int intTotalFound = 0;
            string strMessageResult = "";
            string strDocDate = "";

            bool blClose = false;

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objDT = objForm.DataSources.DataTables.Item("dtSelfWT");
                xmlDoc.LoadXml(objDT.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));
                objNodes = xmlDoc.SelectNodes("/DataTable/Rows/Row[./Cells/Cell[1]/Value/text() = 'Y']");

                if (objNodes.Count > 0)
                {
                    oUDSDate = objForm.DataSources.UserDataSources.Item("udsDate");
                    strDateValue = oUDSDate.ValueEx;
                    oUDSDate = objForm.DataSources.UserDataSources.Item("chkDDate");
                    strDocDate = oUDSDate.ValueEx;
                    if (strDateValue.Trim().Length > 0 || strDocDate == "Y")
                    {
                        intTotalFound = objNodes.Count;

                        #region Go thorough all checked nodes
                        foreach (XmlNode xn in objNodes)
                        {
                            string strDocNum = xn.SelectSingleNode("Cells/Cell[2]/Value").InnerText;

                            oSBOBob = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                            objDocuments = oSBOBob.GetObjectKeyBySingleValue(SAPbobsCOM.BoObjectTypes.oInvoices, "DocNum", strDocNum, SAPbobsCOM.BoQueryConditions.bqc_Equal);
                            if (objDocuments.RecordCount > 0)
                            {
                                objDocuments.MoveFirst();
                                int strDocEntry = (int)objDocuments.Fields.Item(0).Value;


                                oDoc = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                oDoc.GetByKey(strDocEntry);

                                objRecordset = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                strSql = B1.WithholdingTax.Resources.dbQueries.sqlGetActiveCodesSWTForBP;
                                strSql = strSql.Replace("[--CardCode--]", oDoc.CardCode);

                                objRecordset.DoQuery(strSql);

                                if (objRecordset.RecordCount > 0)
                                {
                                    objRecordset.MoveFirst();
                                    while (!objRecordset.EoF)
                                    {
                                        intObjectType = Convert.ToInt32(oDoc.DocObjectCodeEx);
                                        intObjectentry = oDoc.DocEntry;
                                        dbBaseAmnt = oDoc.BaseAmount;
                                        strCardCode = oDoc.CardCode;


                                        oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                                        #region Calculate all SelfWithholding Taxes

                                        objJE = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                                        string strCode = objRecordset.Fields.Item("Code").Value;
                                        double dbPercent = objRecordset.Fields.Item("U_Percent").Value;
                                        string strDebitAccount = objRecordset.Fields.Item("U_DebitAcct").Value;
                                        string strCreditAccount = objRecordset.Fields.Item("U_CreditAccount").Value;

                                        //double dbBase = oDoc.DocTotal - oDoc.VatSum + oDoc.TotalDiscount + oDoc.WTAmount - oDoc.RoundingDiffAmount + oDoc.WTApplied;
                                        double dbBase = dbBaseAmnt;//oDoc.DocTotal - oDoc.VatSum + oDoc.WTAmount - oDoc.RoundingDiffAmount - oDoc.ex;


                                        double dbValue = dbBase * (dbPercent / 100);
                                        objJE.Memo = "Autoretención para el documento " + oDoc.DocNum + " de " + strCode;
                                        objJE.Reference3 = Convert.ToString(oDoc.DocEntry);
                                        objJE.Lines.Credit = dbValue;
                                        objJE.Lines.AccountCode = strCreditAccount;
                                        objJE.Lines.Add();
                                        objJE.Lines.SetCurrentLine(1);
                                        objJE.Lines.Debit = dbValue;
                                        objJE.Lines.AccountCode = strDebitAccount;

                                        if (strDocDate == "Y")
                                        {
                                            objJE.TaxDate = oDoc.TaxDate;
                                        }
                                        else
                                        {

                                            objJE.TaxDate = BYBHelpers.Instance.getDateTimeFormString(strDateValue);
                                        }

                                        int intResult = objJE.Add();
                                        string strMessage = BYBB1MainObject.Instance.B1Company.GetLastErrorDescription();

                                        if (intResult == 0)
                                        {
                                            string strValue = BYBB1MainObject.Instance.B1Company.GetNewObjectKey();


                                            oSelfWithHoldingService = oCompanyService.GetGeneralService("BYB_T1SWTU002");
                                            oSelfHolding = oSelfWithHoldingService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                                            oSelfHolding.SetProperty("U_JEEntry", Convert.ToInt32(strValue));
                                            oSelfHolding.SetProperty("U_BaseAmnt", dbBaseAmnt);
                                            oSelfHolding.SetProperty("U_DocType", intObjectType);
                                            oSelfHolding.SetProperty("U_DocEntry", intObjectentry);
                                            oSelfHolding.SetProperty("U_CardCode", strCardCode.Trim());
                                            oSelfHolding.SetProperty("U_SWTCode", strCode.Trim());
                                            oSelfHolding.SetProperty("U_Total", dbValue);
                                            oSelfHolding.SetProperty("U_DocNum", oDoc.DocNum);
                                            oSelfHolding.SetProperty("U_DocSeries", Convert.ToString(oDoc.Series));
                                            oSelfHolding.SetProperty("U_docId", oDoc.FederalTaxID.Trim());
                                            oSelfHoldingParams = oSelfWithHoldingService.Add(oSelfHolding);

                                            int strCodeA = oSelfHoldingParams.GetProperty("DocEntry");
                                            intTotalDone++;

                                        }
                                        else
                                        {
                                            string strMessage1 = BYBB1MainObject.Instance.B1Company.GetLastErrorDescription();
                                        }





                                        #endregion Calculate all SelfWithholding Taxes

                                        objRecordset.MoveNext();
                                    }
                                }







                            }
                        }
                        strMessageResult = "Se contabilizaron " + intTotalDone.ToString() + " de " + intTotalFound.ToString() + " documentos seleccionados sin error.";
                        blClose = true;
                        #endregion Go thorough all checked nodes






                    }
                    else
                    {
                        strMessageResult = "Por favor seleccionar la fecha de contabilización.";
                    }
                }
                else
                {
                    strMessageResult = "No se han seleccionado autoretenciones para procesar.";
                }
                BYBB1MainObject.Instance.B1Application.MessageBox(strMessageResult);
                if (blClose)
                    objForm.Close();

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

        static private void getSelfWithholdingTaxDocuments(SAPbouiCOM.Form objForm)
        {
            SAPbouiCOM.Grid objGrid = null;
            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.EditTextColumn oEditTExt = null;
            SAPbouiCOM.GridColumn objGridColumn = null;

            try
            {
                objDT = objForm.DataSources.DataTables.Item("dtSelfWT");
                objDT.ExecuteQuery(B1.WithholdingTax.Resources.dbQueries.sqlGetMissingSWT);

                objGrid = objForm.Items.Item("grdSWT").Specific;

                objGridColumn = objGrid.Columns.Item(0);
                objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                objGridColumn.Editable = true;


                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;



                objGridColumn = objGrid.Columns.Item(2);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(3);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(4);
                objGridColumn.Editable = false;
                objGridColumn.RightJustified = true;

                objGrid.AutoResizeColumns();


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

        static private void setBPNameAfterCFL(SAPbouiCOM.ChooseFromListEvent pVal)
        {
            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.Form oForm = null;
            int intRow = -1;
            string strCardName = "";
            SAPbouiCOM.DBDataSource oDS = null;

            try
            {
                if (pVal.ActionSuccess)
                {
                    oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                    oMatrix = oForm.Items.Item("0_U_G").Specific;
                    oDS = oForm.DataSources.DBDataSources.Item("@BYB_T1SWT101");
                    intRow = pVal.Row;
                    if (intRow == 0)
                    {
                        intRow = oMatrix.RowCount;
                    }
                    objDT = pVal.SelectedObjects;
                    strCardName = objDT.GetValue("CardName", 0);
                    oDS.SetValue("U_CardName", intRow - 1, strCardName);
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

        static private void selectAllPendingDocuments(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = null;
            XmlDocument oXML = null;
            XmlNodeList oNodeList = null;
            try
            {
                oXML = new XmlDocument();
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oDT = oForm.DataSources.DataTables.Item("dtSelfWT");
                oXML.LoadXml(oDT.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));
                oNodeList = oXML.SelectNodes("/DataTable/Rows/Row/Cells/Cell[1]/Value");

                if (oNodeList.Count > 0)
                {
                    foreach (XmlNode xn in oNodeList)
                    {
                        if (xn.InnerText == "Y")
                        {
                            xn.InnerText = "N";
                        }
                        else
                        {
                            xn.InnerText = "Y";
                        }

                    }
                }
                oDT.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly, oXML.InnerXml);


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

        #endregion AutoRetencion

        #region Fix WithHolding tax Inconsistences
        static private void getAllNotRegisteredDocuments(SAPbouiCOM.ItemEvent pVal)
        {
            int intResult = 0;
            SAPbobsCOM.Recordset oRecordSet = null;
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.StaticText oStatic = null;

            try
            {

                oRecordSet = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(B1.WithholdingTax.Resources.dbQueries.sqlGetAllWTInconsistences);
                if (oRecordSet.RecordCount > 0)
                {
                    oRecordSet.MoveFirst();
                    intResult = oRecordSet.Fields.Item("Total").Value;

                }
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oStatic = oForm.Items.Item("txtSearch").Specific;
                oStatic.Caption = "Se encontraron " + intResult.ToString() + " inconsistencias en la base de datos";

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

        static private void fixAllNotRegisteredDocuments(SAPbouiCOM.ItemEvent pVal)
        {

            SAPbobsCOM.Recordset oRecordSet = null;
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.StaticText oStatic = null;
            SAPbobsCOM.Documents oDOc = null;
            XmlDocument oXML = null;
            int intTotal = 0;
            //sqlGetAllDocsWithWTInconsistences
            try
            {

                oRecordSet = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(B1.WithholdingTax.Resources.dbQueries.sqlGetAllDocsWithWTInconsistences);
                if (oRecordSet.RecordCount > 0)
                {
                    oRecordSet.MoveFirst();

                    while (!oRecordSet.EoF)
                    {
                        int intDocEntry = oRecordSet.Fields.Item("DocEntry").Value;
                        int intDocType = Convert.ToInt32(oRecordSet.Fields.Item("DocType").Value);

                        if (intDocType == 13)
                        {
                            oDOc = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        }
                        else
                        {
                            oDOc = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                        }
                        if (oDOc.GetByKey(intDocEntry))
                        {
                            Hashtable hashWT = new Hashtable();
                            Hashtable hashTaxExp = new Hashtable();
                            Hashtable hashTaxLines = new Hashtable();
                            int intJE = oDOc.TransNum;

                            for (int i = 0; i < oDOc.Lines.Count; i++)
                            {
                                ArrayList arrLineDet = new ArrayList();
                                #region Recover Tax Lines
                                oDOc.Lines.SetCurrentLine(i);
                                string strTaxCode = oDOc.Lines.TaxCode;
                                if (strTaxCode.Trim().Length > 0)
                                {
                                    double dbPercent = oDOc.Lines.TaxPercentagePerRow;
                                    double dbValue = oDOc.Lines.TaxTotal;
                                    double dbBase = (dbValue * 100) / dbPercent;

                                    if (hashTaxLines.ContainsKey(strTaxCode))
                                    {
                                        arrLineDet = (ArrayList)hashTaxLines[strTaxCode];
                                        arrLineDet[1] = (double)arrLineDet[1] + dbBase;
                                        arrLineDet[2] = (double)arrLineDet[2] + dbValue;
                                        hashTaxLines[strTaxCode] = arrLineDet;
                                    }
                                    else
                                    {
                                        arrLineDet = new ArrayList();
                                        arrLineDet.Add(dbPercent);
                                        arrLineDet.Add(dbBase);
                                        arrLineDet.Add(dbValue);
                                        hashTaxLines.Add(strTaxCode, arrLineDet);
                                    }
                                }

                                #endregion Recover Tax Lines
                            }

                            for (int i = 0; i < oDOc.Expenses.Count; i++)
                            {

                                ArrayList arrLineDet = new ArrayList();
                                #region Recover Tax Expenses
                                oDOc.Expenses.SetCurrentLine(i);
                                string strTaxCode = oDOc.Expenses.TaxCode;
                                if (strTaxCode.Trim().Length > 0)
                                {
                                    double dbPercent = oDOc.Expenses.TaxPercent;
                                    double dbValue = oDOc.Expenses.TaxSum;
                                    double dbBase = (dbValue * 100) / dbPercent;

                                    if (hashTaxExp.ContainsKey(strTaxCode))
                                    {
                                        arrLineDet = (ArrayList)hashTaxExp[strTaxCode];
                                        arrLineDet[1] = (double)arrLineDet[1] + dbBase;
                                        arrLineDet[2] = (double)arrLineDet[2] + dbValue;
                                        hashTaxExp[strTaxCode] = arrLineDet;
                                    }
                                    else
                                    {
                                        arrLineDet = new ArrayList();
                                        arrLineDet.Add(dbPercent);
                                        arrLineDet.Add(dbBase);
                                        arrLineDet.Add(dbValue);
                                        hashTaxExp.Add(strTaxCode, arrLineDet);
                                    }
                                }

                                #endregion Recover Tax Expenses

                            }

                            for (int i = 0; i < oDOc.WithholdingTaxData.Count; i++)
                            {
                                ArrayList arrLineDet = new ArrayList();
                                #region Recover WT
                                oDOc.WithholdingTaxData.SetCurrentLine(i);
                                string strTaxCode = oDOc.WithholdingTaxData.WTCode;
                                if (strTaxCode.Trim().Length > 0)
                                {
                                    double dbValue = oDOc.WithholdingTaxData.WTAmount;
                                    double dbPercent = B1.Helpers.WithHoldingTaxes.getPercentFromCode(strTaxCode);
                                    double dbBase = (dbValue * 100) / dbPercent;





                                    if (hashWT.ContainsKey(strTaxCode))
                                    {
                                        arrLineDet = (ArrayList)hashWT[strTaxCode];
                                        arrLineDet[1] = (double)arrLineDet[1] + dbBase;
                                        arrLineDet[2] = (double)arrLineDet[2] + dbValue;
                                        hashWT[strTaxCode] = arrLineDet;
                                    }
                                    else
                                    {
                                        arrLineDet = new ArrayList();
                                        arrLineDet.Add(dbPercent);
                                        arrLineDet.Add(dbBase);
                                        arrLineDet.Add(dbValue);
                                        hashWT.Add(strTaxCode, arrLineDet);
                                    }
                                }

                                #endregion Recover WT
                            }

                            SAPbobsCOM.GeneralService oGeneralService = null;
                            SAPbobsCOM.GeneralData oGeneralData = null;
                            SAPbobsCOM.GeneralData oChild = null;
                            SAPbobsCOM.GeneralDataCollection oChildren = null;
                            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                            SAPbobsCOM.CompanyService oCompanyService = null;



                            oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                            oGeneralService = oCompanyService.GetGeneralService("BYB_T1WHTMOV");
                            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                            oGeneralData.SetProperty("U_DocEntry", oDOc.DocEntry);
                            oGeneralData.SetProperty("U_DocType", intDocType);
                            oGeneralData.SetProperty("U_CardCode", oDOc.CardCode);

                            oChildren = oGeneralData.Child("BYB_T1WHT401");

                            #region Build Lines

                            foreach (string strKey in hashWT.Keys)
                            {
                                oChild = oChildren.Add();
                                ArrayList objArray = (ArrayList)hashWT[strKey];
                                oChild.SetProperty("U_Type", "I");
                                oChild.SetProperty("U_Operation", "WT");
                                oChild.SetProperty("U_Source", "D");
                                oChild.SetProperty("U_Code", strKey);
                                oChild.SetProperty("U_Percent", objArray[0]);
                                oChild.SetProperty("U_BaseAmnt", objArray[1]);
                                oChild.SetProperty("U_Value", objArray[2]);
                                oChild.SetProperty("U_JE", intJE);


                            }

                            foreach (string strKey in hashTaxLines.Keys)
                            {
                                oChild = oChildren.Add();
                                ArrayList objArray = (ArrayList)hashTaxLines[strKey];
                                oChild.SetProperty("U_Type", "I");
                                oChild.SetProperty("U_Operation", "TAX");
                                oChild.SetProperty("U_Source", "D");
                                oChild.SetProperty("U_Code", strKey);
                                oChild.SetProperty("U_Percent", objArray[0]);
                                oChild.SetProperty("U_BaseAmnt", objArray[1]);
                                oChild.SetProperty("U_Value", objArray[2]);
                                oChild.SetProperty("U_JE", intJE);
                            }

                            foreach (string strKey in hashTaxExp.Keys)
                            {
                                oChild = oChildren.Add();
                                ArrayList objArray = (ArrayList)hashTaxExp[strKey];
                                oChild.SetProperty("U_Type", "I");
                                oChild.SetProperty("U_Operation", "TAX");
                                oChild.SetProperty("U_Source", "E");
                                oChild.SetProperty("U_Code", strKey);
                                oChild.SetProperty("U_Percent", objArray[0]);
                                oChild.SetProperty("U_BaseAmnt", objArray[1]);
                                oChild.SetProperty("U_Value", objArray[2]);
                                oChild.SetProperty("U_JE", intJE);
                            }

                            #endregion


                            oGeneralParams = oGeneralService.Add(oGeneralData);
                            int intNumber = oGeneralParams.GetProperty("DocEntry");
                            intTotal++;
                        }



                        oRecordSet.MoveNext();
                    }


                }
                oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oStatic = oForm.Items.Item("txtFix").Specific;
                oStatic.Caption = "Se corrgieron  " + intTotal.ToString() + " inconsistencias de la base de datos";

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


        #endregion Fix WithHolding tax Inconsistences

        
        static public void formDataAddEvent(SAPbouiCOM.BusinessObjectInfo pVal, out bool blBubbleEvent)
        {
            bool blOpenForm = false;
            Hashtable objWTHash = new Hashtable();
            blBubbleEvent = true;
            SAPbouiCOM.Form objForm = null;
            SAPbobsCOM.Documents objPurch = null;

            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                WTCalculation(objForm, out blOpenForm);
                objWTHash = BYBCache.Instance.getFromCache("WTCalc" + pVal.FormUID);
                if(objWTHash != null)
                {
                    objPurch = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                    

                    oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("BYB_T1WHT200");
                    oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                    oChildren = oGeneralData.Child("BYB_T1WHT201");

                    foreach (string strKey in objWTHash.Keys)
                    {
                        oChild = oChildren.Add();
                        ArrayList objArray = (ArrayList)objWTHash[strKey];
                        oChild.SetProperty("U_WTCode", strKey);
                        oChild.SetProperty("U_WTName", objArray[0]);
                        oChild.SetProperty("U_DbValue", objArray[1]);
                        oChild.SetProperty("U_WTTaxBase", objArray[2]);
                        oChild.SetProperty("U_WTValue", objArray[3]);
                        oChild.SetProperty("U_Account", objArray[4]);
                    }

                    oGeneralData.SetProperty("Code", );
                    oGeneralData.SetProperty("U_Type", "P");
                    oGeneralData.SetProperty("U_DocEntry", dbTotal);

                    oGeneralParams = oGeneralService.Add(oGeneralData);
                    int intNumber = oGeneralParams.GetProperty("DocEntry");
                    if (intNumber > 0)
                    {
                        BYBB1MainObject.Instance.B1Application.MessageBox(MessageStrings.Default.impairmentDoneMessage + "Asiento número: " + BYBB1MainObject.Instance.B1Company.GetNewObjectKey());
                      
                    }
                    else
                    {
                        BYBB1MainObject.Instance.B1Application.MessageBox("Se produjo un error al registrar el detalle de los socios de negocio. El deterioro se contabilizó con el asiento número: " + BYBB1MainObject.Instance.B1Company.GetNewObjectKey());

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
        */
    }
}
