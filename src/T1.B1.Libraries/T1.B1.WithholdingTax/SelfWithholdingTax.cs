using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;
using System.Xml;

namespace T1.B1.WithholdingTax
{
    public class SelfWithholdingTax
    {
        private static SelfWithholdingTax objWithHoldingTax;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private SelfWithholdingTax()
        {
            if (objWithHoldingTax == null)
            {
                objWithHoldingTax = new SelfWithholdingTax();
            }
        }

        public List<SelfWithholdingTaxTransaction> getWTTransaction(int DocEntry)
        {
            List<SelfWithholdingTaxTransaction> objResult = new List<SelfWithholdingTaxTransaction>();
            SAPbobsCOM.GeneralService objService = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralDataParams objParams = null;
            SAPbobsCOM.GeneralData objGeneralData = null;

            try
            {
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objService = objCompanyService.GetGeneralService(Settings._SelfWithHoldingTax.SWtaxUDOTransaction);
                objParams = objService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objParams.SetProperty("DocEntry", DocEntry);
                objGeneralData = objService.GetByParams(objParams);


            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
            return objResult;
        }

        #region Add SWT normal operation

        public static void addSelfWithHoldingTax(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            try
            {
                SAPbobsCOM.Documents objDoc = null;
                if(BusinessObjectInfo.Type == "13")
                {
                    objDoc = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                }
                else if(BusinessObjectInfo.Type == "14")
                {
                    objDoc = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                }

                
                if(objDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey))
                {
                    calcSelfWTax(objDoc,BusinessObjectInfo.Type);
                }
                else
                {
                    _Logger.Error("Could not retrive Document with key " + BusinessObjectInfo.ObjectKey);
                    MainObject.Instance.B1Application.SetStatusBarMessage("T1: Could not retrive Document. Self WithHolding was not calculated");
                }
                
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }
        private static SelfWothholdingTaxResult calcSelfWTax(SAPbobsCOM.Documents objDoc, string DocType)
        {
            SelfWothholdingTaxResult objResult= new SelfWothholdingTaxResult();
            try
            {

                List<SelfWithholdingTaxInfo> lSelfWithHolding = getSelfWithholdingTax(objDoc.CardCode, DocType);
                List<SelfWithholdingTaxInfo> lCalcSWTax = new List<SelfWithholdingTaxInfo>();
                bool blCancelation = false;
                string strThirdParty = "";

                if(DocType == "14")
                {
                    blCancelation = true;

                    #region Check if Based
                    for(int i=0; i < objDoc.Lines.Count; i++)
                    {
                        objDoc.Lines.SetCurrentLine(0);
                        if(objDoc.Lines.BaseType != 13)
                        {
                            lSelfWithHolding = new List<SelfWithholdingTaxInfo>();
                            break;
                        }
                    }
                    #endregion


                }

                if (lSelfWithHolding.Count > 0)
                {
                    double dbBaseAmount = getBaseAmount(objDoc);
                    strThirdParty = getRelatedParty(objDoc);


                    SAPbobsCOM.JournalEntries objJournal = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    objJournal.Memo = string.Concat(new object[] { "Autoretención para el documento ", objDoc.DocNum });
                    objJournal.Reference2 = objDoc.DocEntry.ToString();
                    objJournal.Reference = objDoc.DocNum.ToString();
                    objJournal.ReferenceDate = objDoc.DocDate;
                    objJournal.DueDate = objDoc.DocDueDate;
                    objJournal.TaxDate = objDoc.TaxDate;

                    if (Settings._SelfWithHoldingTax.WTaxTransCode.Length > 0)
                    {
                        objJournal.TransactionCode = Settings._SelfWithHoldingTax.WTaxTransCode;
                    }

                    bool blFirst = true;
                    bool blLines = false;
                    foreach (SelfWithholdingTaxInfo sInfo in lSelfWithHolding)
                    {
                        if (sInfo.MinMount < dbBaseAmount)
                        {
                            blLines = true;
                            double dbWTax = dbBaseAmount * (sInfo.Percentage / 100);
                            if (!blFirst)
                            {
                                objJournal.Lines.Add();
                            }
                            objJournal.Lines.Credit = dbWTax;
                            objJournal.Lines.AccountCode = !blCancelation ? sInfo.Credit : sInfo.Debit;
                            objJournal.Lines.Reference1 = objDoc.DocNum.ToString();
                            objJournal.Lines.Reference2 = objDoc.DocEntry.ToString();
                            objJournal.Lines.LineMemo = string.Concat(new object[] { "Autoretención de ", sInfo.Code, " para el documento ", objDoc.DocNum });
                            objJournal.Lines.UserFields.Fields.Item(Settings._SelfWithHoldingTax.relatedpartyFieldInLines).Value = strThirdParty;
                            objJournal.Lines.Add();
                            objJournal.Lines.Debit = dbWTax;
                            objJournal.Lines.AccountCode = !blCancelation ? sInfo.Debit : sInfo.Credit;
                            objJournal.Lines.Reference1 = objDoc.DocNum.ToString();
                            objJournal.Lines.Reference2 = objDoc.DocEntry.ToString();
                            objJournal.Lines.LineMemo = string.Concat(new object[] { "Autoretención de ", sInfo.Code, " para el documento ", objDoc.DocNum });
                            objJournal.Lines.UserFields.Fields.Item(Settings._SelfWithHoldingTax.relatedpartyFieldInLines).Value = strThirdParty;
                            blFirst = false;

                            sInfo.DocEntry = objDoc.DocEntry;
                            sInfo.DocNum = objDoc.DocNum;
                            sInfo.CardCode = objDoc.CardCode;
                            sInfo.dbBaseAmount = dbBaseAmount;
                            sInfo.dbWtAmount = dbWTax;
                            sInfo.DocType = DocType;
                            lCalcSWTax.Add(sInfo);
                        }

                    }
                    if (blLines)
                    {
                        if (objJournal.Add() == 0)
                        {
                            SAPbobsCOM.CompanyService companyService = null;
                            SAPbobsCOM.GeneralService generalService = null;
                            SAPbobsCOM.GeneralData generalData = null;
                            SAPbobsCOM.GeneralDataParams generalDataParams = null;
                            string newObjectKey = MainObject.Instance.B1Company.GetNewObjectKey();
                            companyService = MainObject.Instance.B1Company.GetCompanyService();
                            generalService = companyService.GetGeneralService(Settings._SelfWithHoldingTax.SWtaxUDOTransaction);

                            Dictionary<string, int> objAdded = new Dictionary<string, int>();
                            foreach (SelfWithholdingTaxInfo sInfo in lCalcSWTax)
                            {
                                generalData = generalService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                                generalData.SetProperty("U_JEEntry", Convert.ToInt32(newObjectKey));
                                generalData.SetProperty("U_BaseAmnt", sInfo.dbBaseAmount);
                                generalData.SetProperty("U_DocType", sInfo.DocType);
                                generalData.SetProperty("U_DocEntry", sInfo.DocEntry);
                                generalData.SetProperty("U_CardCode", sInfo.CardCode);
                                generalData.SetProperty("U_SWTCode", sInfo.Code);
                                generalData.SetProperty("U_Total", sInfo.dbWtAmount);
                                generalData.SetProperty("U_DocNum", sInfo.DocNum);
                                generalData.SetProperty("U_DocDate", objDoc.DocDate);
                                generalData.SetProperty("U_DocTaxDate", objDoc.TaxDate);
                                generalData.SetProperty("U_RELPART", strThirdParty);
                                generalData.SetProperty("U_DocSeries", Convert.ToString(objDoc.Series));
                                generalData.SetProperty("U_LicTradNum", objDoc.FederalTaxID);
                                generalDataParams = generalService.Add(generalData);
                                int property = (int)((dynamic)generalDataParams.GetProperty("DocEntry"));
                                objAdded.Add(sInfo.Code, property);
                            }
                            if (objAdded.Count == lCalcSWTax.Count)
                            {
                                MainObject.Instance.B1Application.SetStatusBarMessage("T1: Las autoretenciones se causaron con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                objResult.Message = "T1: Las autoretenciones se causaron con éxito.";
                                objResult.MessageCode = "";
                            }
                            else
                            {
                                MainObject.Instance.B1Application.SetStatusBarMessage("T1: Ocurrio un error durante la asociacion de la autoretención. Las autoretenciones se causaron con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                objResult.Message = "T1: Ocurrio un error durante la asociacion de la autoretención. Las autoretenciones se causaron con éxito.";
                                objResult.MessageCode = "-2";
                            }

                        }
                        else
                        {
                            string strMessage = MainObject.Instance.B1Company.GetLastErrorDescription();
                            _Logger.Error("Could not create SelfWithHolding Tax. " + strMessage);
                            MainObject.Instance.B1Application.SetStatusBarMessage("T1: Could not create SelfWithHolding Tax. Self WithHolding was not calculated." + strMessage);
                            objResult.Message = "T1: Could not create SelfWithHolding Tax. Self WithHolding was not calculated." + strMessage;
                            objResult.MessageCode = MainObject.Instance.B1Company.GetLastErrorCode().ToString();
                        }
                    }
                }


            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("T1: Could not create SelfWithHolding Tax. Self WithHolding was not calculated." + er.Message);
                objResult.Message = "T1: Could not create SelfWithHolding Tax. Self WithHolding was not calculated." + er.Message;
                objResult.MessageCode = "-1";
            }
            
                return objResult;
            
        }
        #endregion

        private static List<SelfWithholdingTaxInfo> getSelfWithholdingTax(string strCardCode, string DocType)
        {
            List<SelfWithholdingTaxInfo> lSWTH = new List<SelfWithholdingTaxInfo>();
            string isSales = "N";
            string isPurchase = "N";
            string strSQL = "";
            try
            {

                if(DocType == "13" || DocType == "14")
                {
                    isSales = "Y";
                }

                if (DocType == "18" || DocType == "19")
                {
                    isPurchase = "Y";
                }


                if (isSales == "Y")
                {
                    strSQL = Settings._SelfWithHoldingTax.getSelfWithHoldingTaxQuery
                        .Replace("[--CardCode--]", strCardCode)
                        .Replace("[--isSales--]", isSales);
                }

                if (isPurchase == "Y")
                {
                    strSQL = Settings._SelfWithHoldingTax.getSelfWithHoldingTaxQueryPurchase
                    .Replace("[--CardCode--]", strCardCode)
                    
                    .Replace("[--isPurchase--]", isPurchase);
                }

                SAPbobsCOM.Recordset objRS = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRS.DoQuery(strSQL);
                while(!objRS.EoF)
                {
                    SelfWithholdingTaxInfo objSWI= new SelfWithholdingTaxInfo();
                    objSWI.Code = objRS.Fields.Item("Code").Value;
                    objSWI.Credit = objRS.Fields.Item("U_CreditAcct").Value;
                    objSWI.Debit = objRS.Fields.Item("U_DebitAcct").Value;
                    objSWI.Percentage = objRS.Fields.Item("U_Percent").Value;
                    objSWI.MinMount = objRS.Fields.Item("U_MINMOUNT").Value;
                    lSWTH.Add(objSWI);
                    objRS.MoveNext();
                }

            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                lSWTH = null;
            }
            return lSWTH;
        }

        private static double getBaseAmount(SAPbobsCOM.Documents objDoc)
        {
            double dbBase = 0;
            try
            {
                //dbBase = objDoc.BaseAmount;
                if(dbBase == 0)
                {
                    //double dbExpenses = 0;
                    //SAPbobsCOM.DocumentsAdditionalExpenses objExp = objDoc.Expenses;
                    //for(int i=0; i < objExp.Count; i++)
                    //{
                      //  objExp.SetCurrentLine(i);
                        //dbExpenses += objExp.LineTotal;
                    //}
                    dbBase = objDoc.DocTotal - objDoc.VatSum + objDoc.WTAmount + objDoc.RoundingDiffAmount;
                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                dbBase = -1;
            }
            return dbBase;
        }

        private static string getRelatedParty(SAPbobsCOM.Documents objDoc)
        {
            
            string strResult = "";
            
            try
            {
                
                strResult = T1.B1.ReletadParties.Instance.gotRPInfo(objDoc.CardCode);
                
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                strResult = "";
            }
            return strResult;
        }


        #region add Missing Self Withholding Tax

        public static void loadMissingSWTaxForm()
        {
            try
            {
                SAPbouiCOM.Form objForm = T1.B1.Base.UIOperations.Operations.openFormfromXML(T1.B1.WithholdingTax.SWTaxResources.AutoretencionesFaltantes, Settings._SelfWithHoldingTax.MissingSWTFormUID, false);
                objForm.VisibleEx = true;
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void getMissingSWTaxDocuments(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Grid specific = null;
            SAPbouiCOM.DataTable objDTDocuments = null;
            SAPbouiCOM.GridColumn objGridDocuments = null;
            SAPbouiCOM.UserDataSource startDate = null;
            SAPbouiCOM.UserDataSource endDate = null;
            SAPbouiCOM.UserDataSource salesWtax = null;
            SAPbouiCOM.UserDataSource purchWtax = null;
            string strCHKPurch = "N";
            string strCHKSales = "N";
            
            string str = "";
            string str1 = "";
            string str2 = "";
            try
            {
                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
                startDate = objForm.DataSources.UserDataSources.Item("getInDate");
                str1 = startDate.ValueEx.Trim();
                if (str1.Trim().Length == 0)
                {
                    MainObject.Instance.B1Application.MessageBox("Por favor seleccione la fecha de inicio.", 1, "Ok", "", "");
                }
                else
                {
                    endDate = objForm.DataSources.UserDataSources.Item("getEndDate");
                    str2 = endDate.ValueEx.Trim();
                    if (str2.Trim().Length == 0)
                    {
                        MainObject.Instance.B1Application.MessageBox("Por favor seleccione la fecha de fin.", 1, "Ok", "", "");
                    }
                    else
                    {


                        salesWtax = objForm.DataSources.UserDataSources.Item("udsSales");
                        strCHKSales = salesWtax.ValueEx.Trim() == "" ? "N" : salesWtax.ValueEx.Trim();
                        purchWtax = objForm.DataSources.UserDataSources.Item("udsPurch");
                        strCHKPurch = purchWtax.ValueEx.Trim() == "" ? "N" : purchWtax.ValueEx.Trim();
                        

                        
                        if (strCHKSales == "N" && strCHKPurch == "N")
                        {
                            MainObject.Instance.B1Application.MessageBox("Por favor seleccione el tipo de autoretención que desea buscar.", 1, "Ok", "", "");
                        }
                        else
                        {
                            string str3 = Settings._SelfWithHoldingTax.getMissingSWT;
                            
                            str3 = str3.Replace("[--StartDate--]", str1);
                            str3 = str3.Replace("[--EndDate--]", str2);
                            objDTDocuments = objForm.DataSources.DataTables.Item("dtSelfWT");
                            objDTDocuments.ExecuteQuery(str3);
                            if (objDTDocuments.Rows.Count <= 0)
                            {
                                MainObject.Instance.B1Application.MessageBox("No se encontraron documentos faltantes en la fecha especificada", 1, "Ok", "", "");
                            }
                            else
                            {
                                specific = objForm.Items.Item("grdSWT").Specific;

                                objGridDocuments = specific.Columns.Item(0);
                                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                                objGridDocuments.Editable = true;

                                objGridDocuments = specific.Columns.Item(1);
                                SAPbouiCOM.EditTextColumn oCol = (SAPbouiCOM.EditTextColumn)objGridDocuments;
                                oCol.LinkedObjectType = "13";
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;



                                objGridDocuments = specific.Columns.Item(2);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;

                                objGridDocuments = specific.Columns.Item(3);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;

                                objGridDocuments = specific.Columns.Item(4);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = true;

                                objGridDocuments = specific.Columns.Item(5);
                                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;

                                specific.AutoResizeColumns();
                                objForm.Items.Item("grdSWT").Visible = true;
                                objForm.Items.Item("btnCalc").Visible = true;
                            }
                        }
                    }
                }
            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception exception2)
            {
                Exception exception1 = exception2;
                _Logger.Error("", exception1);

            }
        }

        public static void addMisingSWTDocuments(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable variable = null;
            SAPbouiCOM.Form oForm = null;
            SAPbobsCOM.Documents objDoc = null;

            XmlDocument xmlDocument = new XmlDocument();
            XmlNodeList xmlNodeLists = null;

            try
            {
                oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);
                variable = oForm.DataSources.DataTables.Item("dtSelfWT");
                xmlDocument.LoadXml(variable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));

                xmlNodeLists = xmlDocument.SelectNodes("/DataTable/Rows/Row[./Cells/Cell[1]/Value/text() = 'Y']");

                Dictionary<int, List<string>> objResult = new Dictionary<int, List<string>>();
                

                int countJE = xmlNodeLists.Count;
                int intProgress = 1;
                T1.B1.Base.UIOperations.Operations.startProgressBar("Iniciando registro", countJE * 2);
                foreach (XmlNode xmlNodes in xmlNodeLists)
                {

                    string innerText = xmlNodes.SelectSingleNode("Cells/Cell[2]/Value").InnerText;

                    objDoc = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                    
                    List<string> ListMessage = new List<string>();
                    
                    if (objDoc.GetByKey(Convert.ToInt32(innerText)))
                    {
                        Base.UIOperations.Operations.setProgressBarMessage("Calculando el documento " + innerText, intProgress);
                        SelfWothholdingTaxResult objResultOperation = calcSelfWTax(objDoc, "13");
                        ListMessage.Add(objResultOperation.MessageCode + " " + objResultOperation.Message);

                    }
                    else
                    {
                        _Logger.Error("No se pudo recuperar el documento " + innerText);
                    }
                    objResult.Add(Convert.ToInt32(innerText), ListMessage);
                    
                    intProgress++;
                }

                int intLastLine = -1;
                for (int i = 0; i < variable.Rows.Count; i++)
                {
                    Base.UIOperations.Operations.setProgressBarMessage("Actualizando resultados", intProgress);
                    int strJournalEntry = variable.GetValue(1, i);
                    if (objResult.ContainsKey(strJournalEntry))
                    {
                        List<string> objListMsg = objResult[strJournalEntry];
                        string strMsg = "";
                        for (int k = 0; k < objListMsg.Count; k++)
                        {
                            strMsg += objListMsg[k] + ".";
                        }
                        variable.SetValue(5, i, strMsg);
                        intLastLine = i;
                        intProgress++;
                    }
                }

                
                //SAPbouiCOM.Grid specific = oForm.Items.Item("grdSWT").Specific;
                //SAPbouiCOM.GridColumn objGridDocuments = specific.Columns.Item(0);
                //objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                //objGridDocuments.Editable = false;
                Base.UIOperations.Operations.stopProgressBar();

                MainObject.Instance.B1Application.MessageBox("La operación finalizó con éxito. Por favor revise los resultados en el listado", 1, "Ok", "", "");


            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception2 = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception2);
                Base.UIOperations.Operations.stopProgressBar();
            }
            catch (Exception exception4)
            {
                Exception exception3 = exception4;
                _Logger.Error("", exception3);
                Base.UIOperations.Operations.stopProgressBar();
            }
            finally
            {
                if (oForm != null)
                {
                    oForm.Freeze(false);
                }
            }
        }


        #endregion

        #region cancelSWtax Wizard


        public static void setSelectedCode(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.UserDataSource oUDS = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.ChooseFromListEvent oCFLE = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if(objForm != null)
                {
                    oCFLE = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    objDT = oCFLE.SelectedObjects;
                    if(!objDT.IsEmpty)
                    {
                        oUDS = objForm.DataSources.UserDataSources.Item("UD_SWTC");
                        oUDS.ValueEx = objDT.GetValue("Code", 0);
                    }

                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }
        private static bool cancelSWtaxPosting(int JournalEntry, out string Result)
        {
            bool blResult = false;
            Result = "";
            try
            {
                SAPbobsCOM.JournalEntries objJE = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                objJE.GetByKey(JournalEntry);
                if (objJE.Cancel() != 0)
                {
                    Result = "No se pudo cancelar al asiento de Autoretención " + JournalEntry.ToString() + "." + MainObject.Instance.B1Company.GetLastErrorDescription();
                    _Logger.Error("Could not cancel JE " + JournalEntry.ToString() + "." + MainObject.Instance.B1Company.GetLastErrorDescription());

                }
                else
                {
                    blResult = true;
                    Result = "El asiento " + JournalEntry.ToString() + " se canceló con éxito";
                }


            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            return blResult;
        }
        private static bool cancelSWTaxRegistration(int JournalEntry, out string Result)
        {
            bool blResult = false;
            Result = "";
            try
            {
                    SAPbobsCOM.CompanyService companyService = null;
                    SAPbobsCOM.GeneralService generalService = null;
                    SAPbobsCOM.GeneralData generalData = null;
                    SAPbobsCOM.GeneralDataParams generalDataParams = null;
                    companyService = MainObject.Instance.B1Company.GetCompanyService();
                    generalService = companyService.GetGeneralService(Settings._SelfWithHoldingTax.SWtaxUDOTransaction);

                SAPbobsCOM.Recordset objRs = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRs.DoQuery(Settings._SelfWithHoldingTax.getRegistrationFromJEQuery.Replace("[--JE--]", JournalEntry.ToString()));
                while (!objRs.EoF)
                {
                    int DocEntry = objRs.Fields.Item("DocEntry").Value;

                    generalDataParams = generalService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    generalDataParams.SetProperty("DocEntry", DocEntry);


                    try
                    {
                        //generalData = generalService.GetByParams(generalDataParams);
                        //generalData.SetProperty("Canceled", "Y");
                        //generalService.Update(generalData);
                        //generalData.

                        generalService.Cancel(generalDataParams);
                        generalData = generalService.GetByParams(generalDataParams);
                        string strResult = generalData.GetProperty("Canceled");
                        Result += "Internal registration " + DocEntry.ToString() + " canceled";
                    }
                    catch(Exception er)
                    {
                        _Logger.Error("Could not cancel internal registration " + DocEntry.ToString() + "."+ er.Message);
                        Result += "Could not cancel internal registration " + DocEntry.ToString() + "." + er.Message;
                    }
                    
                    objRs.MoveNext();
                }

                if(Result.Length == 0)
                {
                    Result = "No Internal registration found";
                }

                    
                    
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }

            return blResult;
        }
        public static void loadCancelSWTaxForm()
        {
            SAPbouiCOM.Item objItem = null;
            try
            {
                SAPbouiCOM.Form objForm = T1.B1.Base.UIOperations.Operations.openFormfromXML(T1.B1.WithholdingTax.SWTaxResources.CancelarAutoretenciones, Settings._SelfWithHoldingTax.CancelFormUID, false);
                if (objForm != null)
                {

                    if (!Settings._SelfWithHoldingTax.TransactionCodeBase)
                    {
                        objItem = objForm.Items.Item("Item_6");
                        objItem.Visible = true;
                        objItem = objForm.Items.Item("txtSWTCode");
                        objItem.Visible = true;
                    }
                    objForm.VisibleEx = true;
                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }


        public static void getPostedSWTaxDocuments(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Grid specific = null;
            SAPbouiCOM.DataTable objDTDocuments = null;
            SAPbouiCOM.GridColumn objGridDocuments = null;
            SAPbouiCOM.UserDataSource startDate = null;
            SAPbouiCOM.UserDataSource endDate = null;
            SAPbouiCOM.UserDataSource sWtaxCode = null;
            string str = "";
            string str1 = "";
            string str2 = "";
            try
            {
                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
                startDate = objForm.DataSources.UserDataSources.Item("getInDate");
                str1 = startDate.ValueEx.Trim();
                if (str1.Trim().Length == 0)
                {
                    MainObject.Instance.B1Application.MessageBox("Por favor seleccione la fecha de inicio.", 1, "Ok", "", "");
                }
                else
                {
                    endDate = objForm.DataSources.UserDataSources.Item("getEndDate");
                    str2 = endDate.ValueEx.Trim();
                    if (str2.Trim().Length == 0)
                    {
                        MainObject.Instance.B1Application.MessageBox("Por favor seleccione la fecha de fin.", 1, "Ok", "", "");
                    }
                    else
                    {
                        sWtaxCode = objForm.DataSources.UserDataSources.Item("UD_SWTC");
                        str = sWtaxCode.ValueEx.Trim();
                        if (Settings._SelfWithHoldingTax.TransactionCodeBase)
                        {
                            str = "All";
                        }

                        
                        if (str.Trim().Length == 0)
                        {
                            MainObject.Instance.B1Application.MessageBox("Por favor seleccione el codigo de autoretención.", 1, "Ok", "", "");
                        }
                        else
                        {
                            string str3 = Settings._SelfWithHoldingTax.getPostedSWtaxQueryV1;
                            if (Settings._SelfWithHoldingTax.TransactionCodeBase)
                            {
                                str3 = Settings._SelfWithHoldingTax.getPostedSWtaxQueryV2.Replace("[--TransCode--]", Settings._SelfWithHoldingTax.WTaxTransCode);
                            }
                            else
                            {
                                str3 = Settings._SelfWithHoldingTax.getPostedSWtaxQueryV1.Replace("[--SWTCode--]", str);
                            }

                            str3 = str3.Replace("[--StartDate--]", str1);
                            str3 = str3.Replace("[--EndDate--]", str2);
                            objDTDocuments = objForm.DataSources.DataTables.Item("dtSelfWT");
                            objDTDocuments.ExecuteQuery(str3);
                            if (objDTDocuments.Rows.Count <= 0)
                            {
                                MainObject.Instance.B1Application.MessageBox("No se encontraron documentos contabilizados en la fecha especificada", 1, "Ok", "", "");
                            }
                            else
                            {
                                specific = objForm.Items.Item("grdSWT").Specific;
                                objGridDocuments = specific.Columns.Item(0);
                                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                                objGridDocuments.Editable = true;
                                objGridDocuments = specific.Columns.Item(1);
                                SAPbouiCOM.EditTextColumn oCol = (SAPbouiCOM.EditTextColumn)objGridDocuments;
                                oCol.LinkedObjectType = "30";
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                objGridDocuments = specific.Columns.Item(2);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                objGridDocuments = specific.Columns.Item(3);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                objGridDocuments = specific.Columns.Item(4);
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                objGridDocuments = specific.Columns.Item(5);
                                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                                objGridDocuments.Editable = false;
                                objGridDocuments.RightJustified = false;
                                specific.AutoResizeColumns();
                                objForm.Items.Item("grdSWT").Visible = true;
                                objForm.Items.Item("btnCalc").Visible = true;
                            }
                        }
                    }
                }
            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception exception2)
            {
                Exception exception1 = exception2;
                _Logger.Error("", exception1);
                
            }
        }

        public static void cancelPostedTaxDocuments(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable variable = null;
            SAPbouiCOM.Form oForm = null;
            
            XmlDocument xmlDocument = new XmlDocument();
            XmlNodeList xmlNodeLists = null;
            
            try
            {
                oForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);
                variable = oForm.DataSources.DataTables.Item("dtSelfWT");
                xmlDocument.LoadXml(variable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));

                xmlNodeLists = xmlDocument.SelectNodes("/DataTable/Rows/Row[./Cells/Cell[1]/Value/text() = 'Y']");

                Dictionary<int, List<string>> objResult = new Dictionary<int, List<string>>();
                SAPbouiCOM.DBDataSource variable4 = oForm.DataSources.DBDataSources.Item("@BYB_T1SWT100");
                
                        int countJE = xmlNodeLists.Count;
                int intProgress = 1;
                T1.B1.Base.UIOperations.Operations.startProgressBar("Iniciando reversión", countJE * 2);
                        foreach (XmlNode xmlNodes in xmlNodeLists)
                        {
                    
                    string innerText = xmlNodes.SelectSingleNode("Cells/Cell[2]/Value").InnerText;
                    Base.UIOperations.Operations.setProgressBarMessage("Reversando JE " + innerText, intProgress);
                            string strResultJe = "";
                            string strResultInternal = "";
                            List<string> ListMessage = new List<string>();

                            cancelSWtaxPosting(Convert.ToInt32(innerText), out strResultJe);
                            ListMessage.Add(strResultJe);
                            cancelSWTaxRegistration(Convert.ToInt32(innerText), out strResultInternal);
                            ListMessage.Add(strResultInternal);
                            objResult.Add(Convert.ToInt32(innerText), ListMessage);
                    intProgress++;
                        }

                int intLastLine = -1;
                    for(int i=0; i < variable.Rows.Count; i++)
                    {
                    Base.UIOperations.Operations.setProgressBarMessage("Actualizando resultados", intProgress);
                    int strJournalEntry = variable.GetValue(1, i);
                        if (objResult.ContainsKey(strJournalEntry))
                        {
                            List<string> objListMsg = objResult[strJournalEntry];
                            string strMsg = "";
                            for (int k=0; k < objListMsg.Count; k++)
                            {
                                strMsg += objListMsg[k] + ".";
                            }
                            variable.SetValue(5, i, strMsg);
                        intLastLine = i;
                        intProgress++;
                        }
                    }

                oForm.Items.Item("Item_7").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                SAPbouiCOM.Grid specific = oForm.Items.Item("grdSWT").Specific;
                SAPbouiCOM.GridColumn objGridDocuments = specific.Columns.Item(0);
                objGridDocuments.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                objGridDocuments.Editable = false;
                Base.UIOperations.Operations.stopProgressBar();

                MainObject.Instance.B1Application.MessageBox("La operación finalizó con éxito. Por favor revise los resultados en el listado", 1, "Ok", "", "");
                                
                
            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception2 = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception2);
                Base.UIOperations.Operations.stopProgressBar();
            }
            catch (Exception exception4)
            {
                Exception exception3 = exception4;
                _Logger.Error("", exception3);
                Base.UIOperations.Operations.stopProgressBar();
            }
            finally
            {
                if(oForm != null)
                {
                    oForm.Freeze(false);
                }
            }
        }


        #endregion

        #region Autoretención Configuración

        public static void loadSWTaxConfigForm()
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.FormCreationParams objParams = null;
            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objParams.XmlData =  SWTaxResources.SWTaxConfigForm;
                objParams.FormType = "BYB_T1SWT100UDO";
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
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
                objParams.UniqueID = "BYB_MWTRU";
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
                if (MainObject.Instance.B1Application.Menus.Exists("BYB_MWTRU"))
                {
                    MainObject.Instance.B1Application.Menus.RemoveEx("BYB_MWTRU");
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
                objParams.UniqueID = "BYB_MWTDRU";
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
                if (MainObject.Instance.B1Application.Menus.Exists("BYB_MWTDRU"))
                {
                    MainObject.Instance.B1Application.Menus.RemoveEx("BYB_MWTDRU");
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
                objMatrix = objForm.Items.Item("0_U_G").Specific;

                int intRow = eventInfo.Row;
                switch (Action)
                {
                    case "Add":
                        objMatrix.AddRow(1, intRow);

                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_1", "");
                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_2", "");
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

        static public void addAllPBS(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            string strIsSales = "";
            string strIsPurchase = "";
            SAPbouiCOM.DBDataSource objDS = null;
            SAPbouiCOM.Matrix objMatrix = null;
            SAPbobsCOM.Recordset objRS = null;

            SAPbobsCOM.SBObob objBridge = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                
                objDS = objForm.DataSources.DBDataSources.Item("@BYB_T1SWT100");
                strIsSales = objDS.GetValue("U_SALES", objDS.Offset);
                strIsPurchase = objDS.GetValue("U_PURCHASE", objDS.Offset);
                objMatrix = objForm.Items.Item("0_U_G").Specific;
                objMatrix.Clear();
                objForm.Freeze(true);
                objMatrix.FlushToDataSource();
                objBridge = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                int intRow = 0;
                objForm.Freeze(true);

                T1.B1.Base.UIOperations.Operations.startProgressBar("Agregando socios de negocio...", 2);
                if (strIsSales == "Y")
                {
                    
                    objRS = objBridge.GetBPList(SAPbobsCOM.BoCardTypes.cCustomer);
                    while(!objRS.EoF)
                    {
                        string strCardCode = objRS.Fields.Item("CardCode").Value;
                        string strCardName = objRS.Fields.Item("CardName").Value;

                        objMatrix.AddRow(1, intRow);

                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_1", strCardCode);
                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_2", strCardName);
                        objMatrix.FlushToDataSource();
                        intRow++;
                        objRS.MoveNext();
                    }
                }
                if (strIsPurchase == "Y")
                {
                   
                    objRS = objBridge.GetBPList(SAPbobsCOM.BoCardTypes.cSupplier);
                    while (!objRS.EoF)
                    {
                        string strCardCode = objRS.Fields.Item("CardCode").Value;
                        string strCardName = objRS.Fields.Item("CardName").Value;

                        objMatrix.AddRow(1, intRow);

                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_1", strCardCode);
                        objMatrix.SetCellWithoutValidation(intRow + 1, "C_0_2", strCardName);
                        objMatrix.FlushToDataSource();
                        intRow++;
                        objRS.MoveNext();
                    }
                }
                objForm.Freeze(false);
               




            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                if(objForm != null)
                {
                    objForm.Freeze(false);
                }
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
                T1.B1.Base.UIOperations.Operations.setStatusBarMessage("Operación finalizada.", false, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }

        static public void clearAllPBS(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            
            SAPbouiCOM.Matrix objMatrix = null;
            

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                
                objMatrix = objForm.Items.Item("0_U_G").Specific;
                objMatrix.Clear();
                
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                if (objForm != null)
                {
                    objForm.Freeze(false);
                }
            }
        }

        static public void filterAccounts(SAPbouiCOM.ItemEvent pVal, string CFLId)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Conditions objCOnditions = null;
            SAPbouiCOM.Condition objCOnd = null;
            SAPbouiCOM.ChooseFromList objCFL = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item(CFLId);
                objCOnditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                objCOnd = objCOnditions.Add();
                objCOnd.Alias = "Postable";
                objCOnd.CondVal = "Y";
                objCOnd.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                objCFL.SetConditions(objCOnditions);
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void clearfilterAccounts(SAPbouiCOM.ItemEvent pVal, string CFLId)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.ChooseFromList objCFL = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item(CFLId);
                objCFL.SetConditions(null);
                
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void filterBPs(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Conditions objCOnditions = null;
            SAPbouiCOM.Condition objCOnd = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            string strIsSales = "";
            string strIsPurchase = "";
            SAPbouiCOM.DBDataSource objDS = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objDS = objForm.DataSources.DBDataSources.Item("@BYB_T1SWT100");
                strIsSales = objDS.GetValue("U_SALES", objDS.Offset);
                strIsPurchase = objDS.GetValue("U_PURCHASE", objDS.Offset);

                objCFL = objForm.ChooseFromLists.Item("CFL_BP");
                if (strIsSales == "Y" || strIsPurchase == "Y")
                {
                    objCOnditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    if (strIsSales == "Y")
                    {
                        objCOnd = objCOnditions.Add();
                        objCOnd.Alias = "CardType";
                        objCOnd.CondVal = "C";
                        objCOnd.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    }
                    if(strIsSales == "Y" && strIsPurchase == "Y")
                    {
                        objCOnd.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                    }
                    if (strIsPurchase == "Y")
                    {
                        objCOnd = objCOnditions.Add();
                        objCOnd.Alias = "CardType";
                        objCOnd.CondVal = "S";
                        objCOnd.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    }

                    objCFL.SetConditions(objCOnditions);
                }
                else
                {
                    objCOnditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    
                        objCOnd = objCOnditions.Add();
                        objCOnd.Alias = "CardType";
                        objCOnd.CondVal = "X";
                        objCOnd.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    objCFL.SetConditions(objCOnditions);

                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void clearfilterBPs(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.ChooseFromList objCFL = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_BP");
                objCFL.SetConditions(null);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void setBPNameColumn(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;
            SAPbouiCOM.ChooseFromListEvent objCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
            SAPbouiCOM.DataTable objDT = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objMatrix = objForm.Items.Item("0_U_G").Specific;
                SAPbouiCOM.CellPosition objPos = objMatrix.GetCellFocus();
                objDT = objCFLEvent.SelectedObjects;
                if (objDT != null)
                {
                    string strValue = objDT.GetValue("CardName", 0);
                    int intRowNum = objPos.rowIndex;
                    objMatrix.SetCellWithoutValidation(intRowNum, "C_0_2", strValue);
                    objMatrix.FlushToDataSource();
                }
            }
            catch(Exception er)
            {
                _Logger.Error("",er);
            }
        }



        #endregion

        #region SelfwithHlding in Documents

        static public void BYBSelfWithHoldingFolderAdd(string strFormUID)
        {

            SAPbouiCOM.Form objForm = null;
            int intLeft = 0;
            string strUID = "";
            SAPbouiCOM.Item objItemBase = null;
            SAPbouiCOM.Item objItem = null;
            //SAPbouiCOM.Matrix objMatrix = null;
            
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
                objNode = objFormXML.SelectSingleNode("/Application/forms/action/form/items/action/item[@uid='BYB_FSWT']");
                if (objNode != null)
                {
                    blFolderFound = true;
                }
                objForm.Freeze(true);
                if (blFolderFound)
                {
                    objForm.Freeze(true);
                    objItem = objForm.Items.Item(Settings._SelfWithHoldingTax.SelfWithHoldingFolderId);
                }
                else
                {
                    objForm.Freeze(true);
                    string strFolderXML = SWTaxResources.SWTaxFolderDocuments;
                    strFolderXML = strFolderXML.Replace("[--UniqueId--]", strFormUID);
                    MainObject.Instance.B1Application.LoadBatchActions(ref strFolderXML);
                    string strResult = MainObject.Instance.B1Application.GetLastBatchResults();
                    xmlResult = new XmlDocument();
                    xmlResult.LoadXml(strResult);
                    bool errors = xmlResult.SelectSingleNode("/result/errors").HasChildNodes != true ? false : true;
                    if (!errors)
                    {
                        objItem = objForm.Items.Item(Settings._SelfWithHoldingTax.SelfWithHoldingFolderId);

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

                    objItemBase = objForm.Items.Item(Settings._SelfWithHoldingTax.lastFolderId);
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

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

            if (objForm != null)
            {
                objForm.Freeze(false);

            }
        }

        static public void getSWTaxInfoForDocument(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            SAPbobsCOM.Documents objDoc = null;
            string strDocType = "";
            string strSQL = "";
            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.Grid objGrid = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.EditTextColumn oEditTExt = null;
            SAPbouiCOM.GridColumn objGridColumn = null;
            try
            {
                strDocType = BusinessObjectInfo.Type;
                objDoc = MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), strDocType));
                if(objDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey))
                {
                    objForm = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                    objForm.Freeze(true);
                    try
                    {
                        objDT = objForm.DataSources.DataTables.Item("BYB_DSWT");
                    }
                    catch
                    {
                        BYBSelfWithHoldingFolderAdd(BusinessObjectInfo.FormUID);
                        objDT = objForm.DataSources.DataTables.Item("BYB_DSWT");

                    }
                    if (objDT != null)
                    {
                        strSQL = Settings._SelfWithHoldingTax.getAppliedSWTinDOc
                            .Replace("[--DocType--]", strDocType)
                            .Replace("[--DocEntry--]", objDoc.DocEntry.ToString());
                        objDT.ExecuteQuery(strSQL);

                        objGrid = objForm.Items.Item("BYB_GR01").Specific;

                        objGridColumn = objGrid.Columns.Item(0);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;

                        objGridColumn = objGrid.Columns.Item(1);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                        oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                        oEditTExt.RightJustified = true;

                        objGridColumn = objGrid.Columns.Item(2);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                        oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                        oEditTExt.RightJustified = true;

                        objGridColumn = objGrid.Columns.Item(3);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                        oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                        oEditTExt.LinkedObjectType = "30";
                        oEditTExt.RightJustified = true;

                        objGridColumn = objGrid.Columns.Item(4);
                        objGridColumn.Editable = false;
                        objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                        oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;

                        oEditTExt.RightJustified = false;

                    }

                }

            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                if(objForm != null)
                {
                    objForm.Freeze(false);
                }
            }
        }
        #endregion

    }
}
