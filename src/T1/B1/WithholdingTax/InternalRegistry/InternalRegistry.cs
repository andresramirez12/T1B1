using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using T1.Classes;

namespace T1.B1.WithholdingTax.InternalRegistry
{
    public class InternalRegistry
    {
        static private InternalRegistry objInternalRegistry = null;

        private InternalRegistry()
        {
            
        }



        public static void ObjApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if(
                    ( BusinessObjectInfo.FormTypeEx == "133")
                    && !BusinessObjectInfo.BeforeAction
                    && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    )
                {
                    addSalesInvoiceInternalRegistry(BusinessObjectInfo);
                }
                if (
                    (BusinessObjectInfo.FormTypeEx == "141")
                    && !BusinessObjectInfo.BeforeAction
                    && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    )
                {
                    addPurchaseInvoiceInternalRegistry(BusinessObjectInfo);
                }
                if (
                    (BusinessObjectInfo.FormTypeEx == "392")
                    && !BusinessObjectInfo.BeforeAction
                    && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    )
                {
                    addJournalEntryInternalRegistry(BusinessObjectInfo);
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        private static void addSalesInvoiceInternalRegistry(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            SAPbobsCOM.Documents objDocuments = null;
            SAPbobsCOM.CompanyService srvCompanyService = null;
            SAPbobsCOM.GeneralService srvInternalRegistration = null;
            SAPbobsCOM.GeneralData objInternalResgistrationHeader = null;
            SAPbobsCOM.GeneralData objInternalRegistrationDetail = null;

            SAPbobsCOM.GeneralDataCollection oDetailLines = null;
            SAPbobsCOM.GeneralDataParams oInternalRegistrationResults = null;

            int JournalEntry = -1;
            


            try
            {
                objDocuments = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                if (objDocuments.Browser.GetByKeys(BusinessObjectInfo.ObjectKey))
                {
                    Hashtable hashWT = new Hashtable();
                    Hashtable hashTaxExp = new Hashtable();
                    Hashtable hashTaxLines = new Hashtable();

                    JournalEntry = objDocuments.TransNum;

                    for (int i = 0; i < objDocuments.Lines.Count; i++)
                    {
                        ArrayList arrLineDet = new ArrayList();
                        #region Recover Tax Lines
                        objDocuments.Lines.SetCurrentLine(i);
                        string strTaxCode = objDocuments.Lines.TaxCode;
                        if (strTaxCode.Trim().Length > 0)
                        {
                            double dbPercent = objDocuments.Lines.TaxPercentagePerRow;
                            double dbValue = objDocuments.Lines.TaxTotal;
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

                    for (int i = 0; i < objDocuments.Expenses.Count; i++)
                    {

                        ArrayList arrLineDet = new ArrayList();
                        #region Recover Tax Expenses
                        objDocuments.Expenses.SetCurrentLine(i);
                        string strTaxCode = objDocuments.Expenses.TaxCode;
                        if (strTaxCode.Trim().Length > 0)
                        {
                            double dbPercent = objDocuments.Expenses.TaxPercent;
                            double dbValue = objDocuments.Expenses.TaxSum;
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

                    for (int i = 0; i < objDocuments.WithholdingTaxData.Count; i++)
                    {
                        ArrayList arrLineDet = new ArrayList();
                        #region Recover WT
                        objDocuments.WithholdingTaxData.SetCurrentLine(i);
                        string strTaxCode = objDocuments.WithholdingTaxData.WTCode;
                        if (strTaxCode.Trim().Length > 0)
                        {
                            double dbValue = objDocuments.WithholdingTaxData.WTAmount;
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

                    srvCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                    srvInternalRegistration = srvCompanyService.GetGeneralService("BYB_T1WHTMOV");
                    objInternalResgistrationHeader = ((SAPbobsCOM.GeneralData)(srvInternalRegistration.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                    objInternalResgistrationHeader.SetProperty("U_DocEntry", objDocuments.DocEntry);
                    objInternalResgistrationHeader.SetProperty("U_DocType", "13");
                    objInternalResgistrationHeader.SetProperty("U_CardCode", objDocuments.CardCode);

                    oDetailLines = objInternalRegistrationDetail.Child("BYB_T1WHT401");

                    #region Build Lines

                    foreach (string strKey in hashWT.Keys)
                    {
                        objInternalRegistrationDetail = oDetailLines.Add();
                        ArrayList objArray = (ArrayList)hashWT[strKey];
                        objInternalRegistrationDetail.SetProperty("U_Type", "I");
                        objInternalRegistrationDetail.SetProperty("U_Operation", "WT");
                        objInternalRegistrationDetail.SetProperty("U_Source", "D");
                        objInternalRegistrationDetail.SetProperty("U_Code", strKey);
                        objInternalRegistrationDetail.SetProperty("U_Percent", objArray[0]);
                        objInternalRegistrationDetail.SetProperty("U_BaseAmnt", objArray[1]);
                        objInternalRegistrationDetail.SetProperty("U_Value", objArray[2]);
                        objInternalRegistrationDetail.SetProperty("U_JE", JournalEntry);


                    }

                    foreach (string strKey in hashTaxLines.Keys)
                    {
                        objInternalRegistrationDetail = oDetailLines.Add();
                        ArrayList objArray = (ArrayList)hashTaxLines[strKey];
                        objInternalRegistrationDetail.SetProperty("U_Type", "I");
                        objInternalRegistrationDetail.SetProperty("U_Operation", "TAX");
                        objInternalRegistrationDetail.SetProperty("U_Source", "D");
                        objInternalRegistrationDetail.SetProperty("U_Code", strKey);
                        objInternalRegistrationDetail.SetProperty("U_Percent", objArray[0]);
                        objInternalRegistrationDetail.SetProperty("U_BaseAmnt", objArray[1]);
                        objInternalRegistrationDetail.SetProperty("U_Value", objArray[2]);
                        objInternalRegistrationDetail.SetProperty("U_JE", JournalEntry);
                    }

                    foreach (string strKey in hashTaxExp.Keys)
                    {
                        objInternalRegistrationDetail = oDetailLines.Add();
                        ArrayList objArray = (ArrayList)hashTaxExp[strKey];
                        objInternalRegistrationDetail.SetProperty("U_Type", "I");
                        objInternalRegistrationDetail.SetProperty("U_Operation", "TAX");
                        objInternalRegistrationDetail.SetProperty("U_Source", "E");
                        objInternalRegistrationDetail.SetProperty("U_Code", strKey);
                        objInternalRegistrationDetail.SetProperty("U_Percent", objArray[0]);
                        objInternalRegistrationDetail.SetProperty("U_BaseAmnt", objArray[1]);
                        objInternalRegistrationDetail.SetProperty("U_Value", objArray[2]);
                        objInternalRegistrationDetail.SetProperty("U_JE", JournalEntry);
                    }

                    #endregion


                    oInternalRegistrationResults = srvInternalRegistration.Add(objInternalResgistrationHeader);
                    int intNumber = oInternalRegistrationResults.GetProperty("DocEntry");

                }
                }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
        BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        private static void addPurchaseInvoiceInternalRegistry(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            SAPbobsCOM.Documents objDocuments = null;
            SAPbobsCOM.CompanyService srvCompanyService = null;
            SAPbobsCOM.GeneralService srvInternalRegistration = null;
            SAPbobsCOM.GeneralData objInternalResgistrationHeader = null;
            SAPbobsCOM.GeneralData objInternalRegistrationDetail = null;

            SAPbobsCOM.GeneralDataCollection oDetailLines = null;
            SAPbobsCOM.GeneralDataParams oInternalRegistrationResults = null;

            int JournalEntry = -1;



            try
            {
                objDocuments = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                if (objDocuments.Browser.GetByKeys(BusinessObjectInfo.ObjectKey))
                {
                    Hashtable hashWT = new Hashtable();
                    Hashtable hashTaxExp = new Hashtable();
                    Hashtable hashTaxLines = new Hashtable();

                    JournalEntry = objDocuments.TransNum;

                    for (int i = 0; i < objDocuments.Lines.Count; i++)
                    {
                        ArrayList arrLineDet = new ArrayList();
                        #region Recover Tax Lines
                        objDocuments.Lines.SetCurrentLine(i);
                        string strTaxCode = objDocuments.Lines.TaxCode;
                        if (strTaxCode.Trim().Length > 0)
                        {
                            double dbPercent = objDocuments.Lines.TaxPercentagePerRow;
                            double dbValue = objDocuments.Lines.TaxTotal;
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

                    for (int i = 0; i < objDocuments.Expenses.Count; i++)
                    {

                        ArrayList arrLineDet = new ArrayList();
                        #region Recover Tax Expenses
                        objDocuments.Expenses.SetCurrentLine(i);
                        string strTaxCode = objDocuments.Expenses.TaxCode;
                        if (strTaxCode.Trim().Length > 0)
                        {
                            double dbPercent = objDocuments.Expenses.TaxPercent;
                            double dbValue = objDocuments.Expenses.TaxSum;
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

                    for (int i = 0; i < objDocuments.WithholdingTaxData.Count; i++)
                    {
                        ArrayList arrLineDet = new ArrayList();
                        #region Recover WT
                        objDocuments.WithholdingTaxData.SetCurrentLine(i);
                        string strTaxCode = objDocuments.WithholdingTaxData.WTCode;
                        if (strTaxCode.Trim().Length > 0)
                        {
                            double dbValue = objDocuments.WithholdingTaxData.WTAmount;
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

                    srvCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                    srvInternalRegistration = srvCompanyService.GetGeneralService("BYB_T1WHTMOV");
                    objInternalResgistrationHeader = ((SAPbobsCOM.GeneralData)(srvInternalRegistration.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                    objInternalResgistrationHeader.SetProperty("U_DocEntry", objDocuments.DocEntry);
                    objInternalResgistrationHeader.SetProperty("U_DocType", "18");
                    objInternalResgistrationHeader.SetProperty("U_CardCode", objDocuments.CardCode);

                    oDetailLines = objInternalRegistrationDetail.Child("BYB_T1WHT401");

                    #region Build Lines

                    foreach (string strKey in hashWT.Keys)
                    {
                        objInternalRegistrationDetail = oDetailLines.Add();
                        ArrayList objArray = (ArrayList)hashWT[strKey];
                        objInternalRegistrationDetail.SetProperty("U_Type", "I");
                        objInternalRegistrationDetail.SetProperty("U_Operation", "WT");
                        objInternalRegistrationDetail.SetProperty("U_Source", "D");
                        objInternalRegistrationDetail.SetProperty("U_Code", strKey);
                        objInternalRegistrationDetail.SetProperty("U_Percent", objArray[0]);
                        objInternalRegistrationDetail.SetProperty("U_BaseAmnt", objArray[1]);
                        objInternalRegistrationDetail.SetProperty("U_Value", objArray[2]);
                        objInternalRegistrationDetail.SetProperty("U_JE", JournalEntry);


                    }

                    foreach (string strKey in hashTaxLines.Keys)
                    {
                        objInternalRegistrationDetail = oDetailLines.Add();
                        ArrayList objArray = (ArrayList)hashTaxLines[strKey];
                        objInternalRegistrationDetail.SetProperty("U_Type", "I");
                        objInternalRegistrationDetail.SetProperty("U_Operation", "TAX");
                        objInternalRegistrationDetail.SetProperty("U_Source", "D");
                        objInternalRegistrationDetail.SetProperty("U_Code", strKey);
                        objInternalRegistrationDetail.SetProperty("U_Percent", objArray[0]);
                        objInternalRegistrationDetail.SetProperty("U_BaseAmnt", objArray[1]);
                        objInternalRegistrationDetail.SetProperty("U_Value", objArray[2]);
                        objInternalRegistrationDetail.SetProperty("U_JE", JournalEntry);
                    }

                    foreach (string strKey in hashTaxExp.Keys)
                    {
                        objInternalRegistrationDetail = oDetailLines.Add();
                        ArrayList objArray = (ArrayList)hashTaxExp[strKey];
                        objInternalRegistrationDetail.SetProperty("U_Type", "I");
                        objInternalRegistrationDetail.SetProperty("U_Operation", "TAX");
                        objInternalRegistrationDetail.SetProperty("U_Source", "E");
                        objInternalRegistrationDetail.SetProperty("U_Code", strKey);
                        objInternalRegistrationDetail.SetProperty("U_Percent", objArray[0]);
                        objInternalRegistrationDetail.SetProperty("U_BaseAmnt", objArray[1]);
                        objInternalRegistrationDetail.SetProperty("U_Value", objArray[2]);
                        objInternalRegistrationDetail.SetProperty("U_JE", JournalEntry);
                    }

                    #endregion


                    oInternalRegistrationResults = srvInternalRegistration.Add(objInternalResgistrationHeader);
                    int intNumber = oInternalRegistrationResults.GetProperty("DocEntry");

                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        private static void addJournalEntryInternalRegistry(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            SAPbobsCOM.JournalEntries objDocuments = null;
            SAPbobsCOM.CompanyService srvCompanyService = null;
            SAPbobsCOM.GeneralService srvInternalRegistration = null;
            SAPbobsCOM.GeneralData objInternalResgistrationHeader = null;
            SAPbobsCOM.GeneralData objInternalRegistrationDetail = null;

            SAPbobsCOM.GeneralDataCollection oDetailLines = null;
            SAPbobsCOM.GeneralDataParams oInternalRegistrationResults = null;

            int JournalEntry = -1;



            try
            {
                objDocuments = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                if (objDocuments.Browser.GetByKeys(BusinessObjectInfo.ObjectKey))
                {
                    Hashtable hashWT = new Hashtable();
                    Hashtable hashTaxExp = new Hashtable();
                    Hashtable hashTaxLines = new Hashtable();

                    JournalEntry = objDocuments.JdtNum;

                    for (int i = 0; i < objDocuments.Lines.Count; i++)
                    {
                        ArrayList arrLineDet = new ArrayList();
                        #region Recover Tax Lines
                        objDocuments.Lines.SetCurrentLine(i);
                        string strTaxCode = objDocuments.Lines.TaxCode;
                        if (strTaxCode.Trim().Length > 0)
                        {
                            double dbPercent = (100* objDocuments.Lines.TotalTax)/(objDocuments.Lines.BaseSum - objDocuments.Lines.TotalTax);
                            double dbValue = objDocuments.Lines.TotalTax;
                            double dbBase = objDocuments.Lines.BaseSum - objDocuments.Lines.TotalTax;

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


                    for (int i = 0; i < objDocuments.WithholdingTaxData.Count; i++)
                    {
                        ArrayList arrLineDet = new ArrayList();
                        #region Recover WT
                        objDocuments.WithholdingTaxData.SetCurrentLine(i);
                        string strTaxCode = objDocuments.WithholdingTaxData.WTCode;
                        if (strTaxCode.Trim().Length > 0)
                        {
                            double dbValue = objDocuments.WithholdingTaxData.WTAmount;
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

                    srvCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                    srvInternalRegistration = srvCompanyService.GetGeneralService("BYB_T1WHTMOV");
                    objInternalResgistrationHeader = ((SAPbobsCOM.GeneralData)(srvInternalRegistration.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                    objInternalResgistrationHeader.SetProperty("U_DocEntry", objDocuments.JdtNum);
                    objInternalResgistrationHeader.SetProperty("U_DocType", "30");
                    //objInternalResgistrationHeader.SetProperty("U_CardCode", objDocuments.CardCode);

                    oDetailLines = objInternalRegistrationDetail.Child("BYB_T1WHT401");

                    #region Build Lines

                    foreach (string strKey in hashWT.Keys)
                    {
                        objInternalRegistrationDetail = oDetailLines.Add();
                        ArrayList objArray = (ArrayList)hashWT[strKey];
                        objInternalRegistrationDetail.SetProperty("U_Type", "I");
                        objInternalRegistrationDetail.SetProperty("U_Operation", "WT");
                        objInternalRegistrationDetail.SetProperty("U_Source", "D");
                        objInternalRegistrationDetail.SetProperty("U_Code", strKey);
                        objInternalRegistrationDetail.SetProperty("U_Percent", objArray[0]);
                        objInternalRegistrationDetail.SetProperty("U_BaseAmnt", objArray[1]);
                        objInternalRegistrationDetail.SetProperty("U_Value", objArray[2]);
                        objInternalRegistrationDetail.SetProperty("U_JE", JournalEntry);


                    }

                    foreach (string strKey in hashTaxLines.Keys)
                    {
                        objInternalRegistrationDetail = oDetailLines.Add();
                        ArrayList objArray = (ArrayList)hashTaxLines[strKey];
                        objInternalRegistrationDetail.SetProperty("U_Type", "I");
                        objInternalRegistrationDetail.SetProperty("U_Operation", "TAX");
                        objInternalRegistrationDetail.SetProperty("U_Source", "D");
                        objInternalRegistrationDetail.SetProperty("U_Code", strKey);
                        objInternalRegistrationDetail.SetProperty("U_Percent", objArray[0]);
                        objInternalRegistrationDetail.SetProperty("U_BaseAmnt", objArray[1]);
                        objInternalRegistrationDetail.SetProperty("U_Value", objArray[2]);
                        objInternalRegistrationDetail.SetProperty("U_JE", JournalEntry);
                    }

                    foreach (string strKey in hashTaxExp.Keys)
                    {
                        objInternalRegistrationDetail = oDetailLines.Add();
                        ArrayList objArray = (ArrayList)hashTaxExp[strKey];
                        objInternalRegistrationDetail.SetProperty("U_Type", "I");
                        objInternalRegistrationDetail.SetProperty("U_Operation", "TAX");
                        objInternalRegistrationDetail.SetProperty("U_Source", "E");
                        objInternalRegistrationDetail.SetProperty("U_Code", strKey);
                        objInternalRegistrationDetail.SetProperty("U_Percent", objArray[0]);
                        objInternalRegistrationDetail.SetProperty("U_BaseAmnt", objArray[1]);
                        objInternalRegistrationDetail.SetProperty("U_Value", objArray[2]);
                        objInternalRegistrationDetail.SetProperty("U_JE", JournalEntry);
                    }

                    #endregion


                    oInternalRegistrationResults = srvInternalRegistration.Add(objInternalResgistrationHeader);
                    int intNumber = oInternalRegistrationResults.GetProperty("DocEntry");

                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
        }
    }
}
