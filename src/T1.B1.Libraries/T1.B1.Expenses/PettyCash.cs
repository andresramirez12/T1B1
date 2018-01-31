using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;
using System.Xml;
using System.Globalization;

namespace T1.B1.Expenses
{
    public class PettyCash
    {
        private static PettyCash objExpenses;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private PettyCash()
        {

        }

        public static List<projectInfo> getProjectList()
        {
            List<projectInfo> objResult = null;
            SAPbobsCOM.CompanyService objCmpService = null;
            SAPbobsCOM.ProjectsService objProjectService = null;
            SAPbobsCOM.ProjectsParams objProjectList = null;
            SAPbobsCOM.ProjectParams objProjectInfo = null;

            try
            {
                objResult = new List<projectInfo>();
                objCmpService = MainObject.Instance.B1Company.GetCompanyService();
                objProjectService = objCmpService.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService);
                objProjectList = objProjectService.GetProjectList();
                for (int i = 0; i < objProjectList.Count; i++)
                {
                    objProjectInfo = objProjectList.Item(i);
                    projectInfo oProj = new projectInfo();
                    oProj.ProjectCode = objProjectInfo.Code;
                    oProj.ProjectName = objProjectInfo.Name;
                    objResult.Add(oProj);

                }


            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                objResult = new List<projectInfo>();
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objResult = new List<projectInfo>();
            }
            return objResult;
        }

        public static List<costCenterInfo> getProfitCenterList()
        {
            List<costCenterInfo> objResult = null;
            SAPbobsCOM.CompanyService objCmpService = null;
            SAPbobsCOM.ProfitCentersService objProfitCenterService = null;
            SAPbobsCOM.ProfitCentersParams objProfitCenterList = null;
            SAPbobsCOM.ProfitCenterParams objProfitCenterInfo = null;


            try
            {
                objResult = new List<costCenterInfo>();
                objCmpService = MainObject.Instance.B1Company.GetCompanyService();
                objProfitCenterService = objCmpService.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService);
                objProfitCenterList = objProfitCenterService.GetProfitCenterList();
                for (int i = 0; i < objProfitCenterList.Count; i++)
                {
                    objProfitCenterInfo = objProfitCenterList.Item(i);
                    SAPbobsCOM.ProfitCenterParams objTemp = objProfitCenterService.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenterParams);
                    objTemp.CenterCode = objProfitCenterInfo.CenterCode;
                    SAPbobsCOM.ProfitCenter objProfitCenter = objProfitCenterService.GetProfitCenter(objTemp);
                    costCenterInfo oProfit = new costCenterInfo();
                    oProfit.CostCenterCode = objProfitCenter.CenterCode;
                    oProfit.CostCenterName = objProfitCenter.CenterName;
                    oProfit.DimensionCode = objProfitCenter.InWhichDimension;


                    objResult.Add(oProfit);

                }


            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                objResult = new List<costCenterInfo>();
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objResult = new List<costCenterInfo>();
            }
            return objResult;
        }

        public static List<dimensionInfo> getDimensionsList()
        {
            List<dimensionInfo> objResult = null;
            SAPbobsCOM.CompanyService objCmpService = null;
            SAPbobsCOM.DimensionsService objPDimensionsService = null;
            SAPbobsCOM.DimensionsParams objDimensionsList = null;
            SAPbobsCOM.DimensionParams objDimensionInfo = null;


            try
            {
                objResult = new List<dimensionInfo>();
                objCmpService = MainObject.Instance.B1Company.GetCompanyService();
                objPDimensionsService = objCmpService.GetBusinessService(SAPbobsCOM.ServiceTypes.DimensionsService);
                objDimensionsList = objPDimensionsService.GetDimensionList();
                for (int i = 0; i < objDimensionsList.Count; i++)
                {
                    objDimensionInfo = objDimensionsList.Item(i);
                    SAPbobsCOM.DimensionParams objTemp = objPDimensionsService.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams);
                    objTemp.DimensionCode = objDimensionInfo.DimensionCode;
                    SAPbobsCOM.Dimension objDimensions = objPDimensionsService.GetDimension(objTemp);
                    dimensionInfo oDImension = new dimensionInfo();
                    oDImension.DimentionCode = objDimensions.DimensionCode;
                    oDImension.DimensionName = objDimensions.DimensionDescription;
                    oDImension.isActive = objDimensions.IsActive == SAPbobsCOM.BoYesNoEnum.tYES ? true : false;

                    //oDImension.DimensionLevel = objDimensions..InWhichDimension;


                    objResult.Add(oDImension);

                }


            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                objResult = new List<dimensionInfo>();
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objResult = new List<dimensionInfo>();
            }
            return objResult;
        }

        public static bool isMultipleDimension()
        {
            bool blResult = false;
            bool blIsHANA = false;
            SAPbobsCOM.Recordset objRcordSet = null;
            string strSQL = "";
            string strResult = "N";
            try
            {
                blIsHANA = CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName) != null ? CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName) : false;
                strSQL = blIsHANA ? Settings._HANA.getMultipleDimQuery : Settings._SQL.getMultipleDimQuery;
                objRcordSet = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRcordSet.DoQuery(strSQL);
                if (objRcordSet.RecordCount > 0)
                {
                    strResult = objRcordSet.Fields.Item(0).Value;
                    blResult = strResult.Trim() == "N" ? false : true;
                }
            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                blResult = false;
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                blResult = false;
            }
            return blResult;
        }

        #region Apertura
        static public void loadPettyCashPaymentForm()
        {
            string strSQL = "";
            
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.FormCreationParams objParams = null;

            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objParams.XmlData = ExpensesRes.BYBPTC_CMPayment;
                objParams.FormType = Settings._MainPettyCash.pettyCashPaymentFormType.Trim();
                objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objForm.Visible = true;

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public static void filterAccountPCPaymentForm(SAPbouiCOM.ItemEvent pVal, bool clearFilter)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.UserDataSource oDS = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_ACT");
                if (clearFilter)
                {
                    objCFL.SetConditions(null);
                    SAPbouiCOM.ChooseFromListEvent oEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    if (oEvent.SelectedObjects != null && oEvent.SelectedObjects.Rows.Count == 1)
                    {
                        oDS = objForm.DataSources.UserDataSources.Item("UD_ACCT");
                        oDS.Value = oEvent.SelectedObjects.GetValue("AcctCode", 0);
                    }


                }
                else
                {
                    objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    objCondition = objConditions.Add();
                    objCondition.Alias = "Postable";
                    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    objCondition.CondVal = "Y";
                    objCFL.SetConditions(objConditions);
                }





            }
            catch (COMException comEx)
            {


                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void filterPettyCashPCPaymentForm(SAPbouiCOM.ItemEvent pVal, bool clearFilter)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.UserDataSource oDS = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_PC");
                if (clearFilter)
                {
                    objCFL.SetConditions(null);
                    SAPbouiCOM.ChooseFromListEvent oEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    if (oEvent.SelectedObjects != null && oEvent.SelectedObjects.Rows.Count == 1)
                    {
                        oDS = objForm.DataSources.UserDataSources.Item("UD_PC");
                        oDS.Value = oEvent.SelectedObjects.GetValue("Code", 0);
                    }


                }
                else
                {
                    objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    objCondition = objConditions.Add();
                    objCondition.Alias = "U_ISPOSTED";
                    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    objCondition.CondVal = "N";
                    objCFL.SetConditions(objConditions);
                }





            }
            catch (COMException comEx)
            {


                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void addPaymentDocument(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable objDocs = null;
            SAPbouiCOM.Form objForm = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objPettyCash = null;
            SAPbobsCOM.GeneralService objRelatedParties = null;


            SAPbobsCOM.GeneralData objPettyCashData = null;
            SAPbobsCOM.GeneralData objRelatedPartiesData = null;
            SAPbobsCOM.GeneralDataParams objFilterParams = null;

            SAPbouiCOM.UserDataSource objPaymentDate = null;
            SAPbouiCOM.UserDataSource objCashAccount = null;
            SAPbouiCOM.UserDataSource objPC = null;

            SAPbobsCOM.Payments objPayment = null;
            SAPbobsCOM.JournalEntries objJE = null;
            bool blResultPayment = false;
            int intResult = -1;
            string strMessage = "";

            #region DocumentValues

            string strCashAccount = "";
            string strPettyCash = "";
            DateTime dtPaymentDate;
            double dbValue = 0;
            string strPettyCashAccount = "";
            string strControlAccount = "";
            string strCardCode = "";
            string strResponsable = "";
            string strCardName = "";

            bool blIsAssociated = false;


            #endregion

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                #region Get Form information
                objPaymentDate = objForm.DataSources.UserDataSources.Item("UD_DATE");
                objCashAccount = objForm.DataSources.UserDataSources.Item("UD_ACCT");
                objPC = objForm.DataSources.UserDataSources.Item("UD_PC");

                strCashAccount = objCashAccount.Value.Trim();
                strPettyCash = objPC.Value.Trim();

                SAPbobsCOM.SBObob objTemp = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                SAPbobsCOM.Recordset objRS = objTemp.Format_StringToDate(objPaymentDate.Value);
                dtPaymentDate = objRS.Fields.Item(0).Value;// Convert.ToDateTime(objPaymentDate.Value,CultureInfo.InvariantCulture);
                if(dtPaymentDate < new DateTime(2013,1,1))
                {
                    dtPaymentDate = DateTime.Now;
                }
                #endregion

                if (strCashAccount.Length > 0
                    && strPettyCash.Length > 0
                    )
                {
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    objPettyCash = objCompanyService.GetGeneralService(Settings._MainPettyCash.pettyCashUDO);
                    objFilterParams = objPettyCash.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    objFilterParams.SetProperty("Code", strPettyCash);
                    objPettyCashData = objPettyCash.GetByParams(objFilterParams);
                    strPettyCashAccount = objPettyCashData.GetProperty("U_PCACCOUNT");
                    strControlAccount = objPettyCashData.GetProperty("U_CTRLACCT");
                    dbValue = objPettyCashData.GetProperty("U_VALUE");
                    strResponsable = objPettyCashData.GetProperty("U_TERRELA");

                    blIsAssociated = T1.B1.Base.DIOperations.Operations.isAccountAsociated(strControlAccount);
                    
                    
                    if (strResponsable.Trim().Length > 0)
                    {
                        objRelatedParties = objCompanyService.GetGeneralService(Settings._Main.RelatedPartyUDO);
                        objFilterParams = objRelatedParties.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        objFilterParams.SetProperty("Code", strResponsable);
                        objRelatedPartiesData = objRelatedParties.GetByParams(objFilterParams);
                        if (objRelatedPartiesData != null)
                        {
                            strCardCode = objRelatedPartiesData.GetProperty("U_CARDCODE");
                            strCardName = objRelatedPartiesData.GetProperty("U_LEGALNAME");
                        }

                        #region Asiento Contable de Apertura

                        objJE = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        objJE.ReferenceDate = dtPaymentDate;
                        objJE.TaxDate = dtPaymentDate;
                        objJE.DueDate = dtPaymentDate;
                        objJE.Reference3 = strPettyCash;
                        objJE.Memo = "Contabilización de apertura Cajas Menor " + strPettyCash;
                        objJE.Lines.AccountCode = strPettyCashAccount;

                        objJE.Lines.Debit = dbValue;
                        objJE.Lines.UserFields.Fields.Item("U_BYB_RELPAR").Value = strResponsable;
                        objJE.Lines.Add();
                        if (blIsAssociated)
                        {
                            objJE.Lines.ShortName = strCardCode;
                        }
                        objJE.Lines.ControlAccount = strControlAccount;
                        
                        objJE.Lines.Credit = dbValue;
                        objJE.Lines.UserFields.Fields.Item("U_BYB_RELPAR").Value = strResponsable;

                        #endregion

                        #region Desembolso

                        
                        if (blIsAssociated)
                        {

                            objPayment = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                        objPayment.CardCode = strCardCode;
                        objPayment.DocDate = dtPaymentDate;
                        objPayment.ControlAccount = strControlAccount;
                        objPayment.TransferAccount = strCashAccount;
                        objPayment.TransferSum = dbValue;
                        objPayment.TransferDate = dtPaymentDate;
                        objPayment.TaxDate = dtPaymentDate;
                        objPayment.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;
                        objPayment.Remarks = "Contabilización de apertura Cajas Menor " + strPettyCash;
                        }
                        else
                        {
                            objPayment = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                            objPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
                            objPayment.CardCode = strCardCode;
                            objPayment.CardName = strCardName;
                            objPayment.DocDate = dtPaymentDate;
                            objPayment.AccountPayments.AccountCode = strControlAccount;
                            objPayment.AccountPayments.Decription = "Desembolso caja menor "  + strPettyCash;
                            objPayment.AccountPayments.SumPaid = dbValue;
                            
                            objPayment.TransferAccount = strCashAccount;
                            objPayment.TransferSum = dbValue;
                            objPayment.TransferDate = dtPaymentDate;
                            objPayment.TaxDate = dtPaymentDate;
                            objPayment.Remarks = "Contabilización de apertura Cajas Menor " + strPettyCash;
                        }


                        #endregion

                        if (MainObject.Instance.B1Company.InTransaction)
                        {
                            MainObject.Instance.B1Application.MessageBox("La compañia se encuentra en una transacción. Por favor inténtelo en unos minutos.");
                        }
                        else
                        {
                           
                            MainObject.Instance.B1Company.StartTransaction();
                            intResult = objJE.Add();
                            if (intResult == 0)
                            {
                                intResult = objPayment.Add();
                                if (intResult == 0)
                                {
                                    blResultPayment = true;
                                    objPettyCashData.SetProperty("U_ISPOSTED", "Y");
                                    objPettyCash.Update(objPettyCashData);
                                    MainObject.Instance.B1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                }
                                else
                                {
                                    strMessage = MainObject.Instance.B1Company.GetLastErrorDescription();
                                    _Logger.Error(strMessage);
                                    MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                }
                            }
                            else
                            {
                                strMessage = MainObject.Instance.B1Company.GetLastErrorDescription();
                                _Logger.Error(strMessage);
                                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            }
                        }
                    }
                    else
                    {
                        strMessage = "El responsable de la caja menor es un campo requerido. Por favor verifique su configuración";
                        MainObject.Instance.B1Application.MessageBox(strMessage);
                    }
                }
                else
                {
                    MainObject.Instance.B1Application.MessageBox("Por favor seleccione todos los valores para hacer el desembolso.");
                }
                if (blResultPayment)
                {
                    objForm.Close();
                    MainObject.Instance.B1Application.SetStatusBarMessage("BYB: La apertura de la caja se realizó con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    MainObject.Instance.B1Application.SetStatusBarMessage("BYB: "+ strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }


            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                if(MainObject.Instance.B1Company.InTransaction)
                {
                    MainObject.Instance.B1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                if (MainObject.Instance.B1Company.InTransaction)
                {
                    MainObject.Instance.B1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
            }
        }
        #endregion

        #region Conceptos

        public static void loadBPNameCFL(SAPbouiCOM.Form objForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                SAPbouiCOM.ChooseFromListEvent objCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                SAPbouiCOM.DataTable objDT = objCFLEvent.SelectedObjects;


                if (objDT != null && objDT.Rows.Count > 0)
                {
                    SAPbouiCOM.DBDataSource variable5 = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC202");
                    variable5.SetValue("U_RELPARNAME", variable5.Offset, objDT.GetValue("U_LEGALNAME", 0));

                }





            }
            catch (COMException comEx)
            {


                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }


        }

        public static void filterAccountConceptsUDO(SAPbouiCOM.ItemEvent pVal, bool clearFilter)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_ACCT");
                if (clearFilter)
                {
                    objCFL.SetConditions(null);
                }
                else
                {
                    objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    objCondition = objConditions.Add();
                    objCondition.Alias = "Postable";
                    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    objCondition.CondVal = "Y";
                    objCFL.SetConditions(objConditions);
                }





            }
            catch (COMException comEx)
            {


                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
        #endregion

        #region Registro Legalizacion Caja Menor
        public static void openLegalizationForm(SAPbouiCOM.MenuEvent pVal)
        {
            if (PettyCash.objExpenses == null)
            {
                PettyCash.objExpenses = new T1.B1.Expenses.PettyCash();
            }

            try
            {
                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "BYB_T1PTC300", "");
                configureLegalizationForm(CacheManager.CacheManager.Instance.getFromCache(Settings._Main.LegalizationFormLastId), true);

            }
            catch (COMException comEx)
            {


                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void configureLegalizationForm(string FormUID, bool ReloadCombos)
        {
            if (PettyCash.objExpenses == null)
            {
                PettyCash.objExpenses = new T1.B1.Expenses.PettyCash();
            }
            
            SAPbouiCOM.ComboBox objCombo = null;
            SAPbouiCOM.EditText objEdit = null;
            SAPbouiCOM.Form objForm = null;
            
            string strNextNumber = "";

            try
            {
                if (FormUID.Trim().Length > 0)
                {
                    objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
                    objForm.Freeze(true);


                    #region fill series drop

                    if (ReloadCombos)
                    {
                        objCombo = objForm.Items.Item("cmbSeries").Specific;
                        objCombo.ValidValues.LoadSeries("BYB_T1PTC300", SAPbouiCOM.BoSeriesMode.sf_Add);


                        objEdit = objForm.Items.Item("1_U_E").Specific;
                        strNextNumber = objForm.BusinessObject.GetNextSerialNumber(objCombo.Value).ToString();
                        objEdit.Value = strNextNumber;

                        loadMatrixComboBoxes(FormUID);
                    }

                    #endregion


                    
                }
            }
            catch (COMException comEx)
            {


                _Logger.Error("", comEx);

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
                CacheManager.CacheManager.Instance.removeFromCache(Settings._Main.ExpenseRequestFormLastId);
            }
        }

        public static void loadMatrixComboBoxes(string FormUID)
        {
            if (PettyCash.objExpenses == null)
            {
                PettyCash.objExpenses = new T1.B1.Expenses.PettyCash();
            }

            bool isMultiDim = false;
            List<dimensionInfo> objDimList = null;
            List<costCenterInfo> objProfitCenterList = null;
            List<projectInfo> objProjectList = null;
            SAPbouiCOM.ComboBox objCombo = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;
            SAPbouiCOM.Column objColumn = null;
            int intNumberOfRows = -1;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
                objForm.Freeze(true);


                if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    isMultiDim = isMultipleDimension();
                    if (isMultiDim)
                    {
                        objDimList = getDimensionsList();

                    }


                    objProfitCenterList = getProfitCenterList();

                    objProjectList = getProjectList();
                    objMatrix = objForm.Items.Item("0_U_G").Specific;
                    intNumberOfRows = objMatrix.RowCount;


                    #region load DropDowns Dimensions and ProfitCenter


                    if (!isMultiDim)
                    {
                        #region ProfitCenter
                        if (intNumberOfRows > 0)
                        {
                            objColumn = objMatrix.Columns.Item("C_0_10");
                            objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
                            for (int i = 0; i < objCombo.ValidValues.Count; i++)
                            {
                                objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }

                            foreach (costCenterInfo oPC in objProfitCenterList)
                            {
                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                            }
                            updateVisualizarionString(objForm.TypeEx, "C_0_10", "0_U_G", "Centro de Costo");
                            objColumn = objMatrix.Columns.Item("C_0_11");
                            objColumn.Visible = false;
                            objColumn = objMatrix.Columns.Item("C_0_12");
                            objColumn.Visible = false;
                            objColumn = objMatrix.Columns.Item("C_0_13");
                            objColumn.Visible = false;
                            objColumn = objMatrix.Columns.Item("Col_14");
                            objColumn.Visible = false;

                        }
                        #endregion
                    }
                    else
                    {
                        #region Dimensions

                        foreach (dimensionInfo oDIM in objDimList)
                        {
                            switch (oDIM.DimentionCode)
                            {
                                case 1:
                                    if (oDIM.isActive)
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_10");
                                        objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
                                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                                        {
                                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                        }

                                        foreach (costCenterInfo oPC in objProfitCenterList)
                                        {
                                            if (oPC.DimensionCode == 1)
                                            {
                                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                                            }
                                        }
                                        updateVisualizarionString(objForm.TypeEx, "C_0_10", "0_U_G", oDIM.DimensionName);

                                    }

                                    break;
                                case 2:
                                    if (oDIM.isActive)
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_11");
                                        objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
                                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                                        {
                                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                        }

                                        foreach (costCenterInfo oPC in objProfitCenterList)
                                        {
                                            if (oPC.DimensionCode == 2)
                                            {
                                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                                            }
                                        }
                                        updateVisualizarionString(objForm.TypeEx, "C_0_11", "0_U_G", oDIM.DimensionName);

                                    }
                                    else
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_11");
                                        objColumn.Visible = false;
                                    }
                                    break;
                                case 3:
                                    if (oDIM.isActive)
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_12");
                                        objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
                                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                                        {
                                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                        }

                                        foreach (costCenterInfo oPC in objProfitCenterList)
                                        {
                                            if (oPC.DimensionCode == 3)
                                            {
                                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                                            }
                                        }
                                        updateVisualizarionString(objForm.TypeEx, "C_0_12", "0_U_G", oDIM.DimensionName);

                                    }
                                    else
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_12");
                                        objColumn.Visible = false;
                                    }
                                    break;
                                case 4:
                                    if (oDIM.isActive)
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_13");
                                        objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
                                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                                        {
                                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                        }

                                        foreach (costCenterInfo oPC in objProfitCenterList)
                                        {
                                            if (oPC.DimensionCode == 4)
                                            {
                                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                                            }
                                        }
                                        updateVisualizarionString(objForm.TypeEx, "C_0_13", "0_U_G", oDIM.DimensionName);

                                    }
                                    else
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_13");
                                        objColumn.Visible = false;
                                    }
                                    break;
                                case 5:
                                    if (oDIM.isActive)
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_14");
                                        objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
                                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                                        {
                                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                        }

                                        foreach (costCenterInfo oPC in objProfitCenterList)
                                        {
                                            if (oPC.DimensionCode == 5)
                                            {
                                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                                            }
                                        }
                                        updateVisualizarionString(objForm.TypeEx, "C_0_14", "0_U_G", oDIM.DimensionName);

                                    }
                                    else
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_14");
                                        objColumn.Visible = false;
                                    }
                                    break;
                            }
                        }
                        #endregion


                    }
                    #endregion


                    #region Projects

                    if (intNumberOfRows > 0)
                    {
                        objColumn = objMatrix.Columns.Item("C_0_9");
                        objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                        {
                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }

                        foreach (projectInfo oPC in objProjectList)
                        {
                            objCombo.ValidValues.Add(oPC.ProjectCode, oPC.ProjectName);

                        }
                    }

                    #endregion

                }

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                objForm.Refresh();
                objForm.Freeze(false);
            }
        }

        private static void updateVisualizarionString(string FormUID, string ColumnID, string ItemID, string NewValue)
        {
            SAPbobsCOM.DynamicSystemStrings objDynamicStrings = null;
            int intResult = -1;
            try
            {
                objDynamicStrings = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDynamicSystemStrings);
                if (!objDynamicStrings.GetByKey(FormUID, ColumnID, ItemID))
                {
                    objDynamicStrings.FormID = FormUID;
                    objDynamicStrings.ItemID = ItemID;
                    if (ColumnID.Trim().Length > 0)
                    {
                        objDynamicStrings.ColumnID = ColumnID;
                    }
                    objDynamicStrings.ItemString = NewValue;
                    intResult = objDynamicStrings.Add();
                    if (intResult != 0)
                    {
                        _Logger.Error(MainObject.Instance.B1Company.GetLastErrorDescription());
                    }

                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }


        public static void filterValidPCUDO(SAPbouiCOM.ItemEvent pVal, bool clearFilter)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.ChooseFromListEvent oEvent = null;
            SAPbouiCOM.DBDataSource oDBDS = null;
            SAPbouiCOM.UserDataSource oUDS = null;
            int intDocEntry = -1;
            string strPC = "";
            PCLegalizationFormCache objLegalizationFormCache = null;
            double dbValue = 0;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_PC");
                if (clearFilter)
                {
                    objCFL.SetConditions(null);
                    if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (oEvent.SelectedObjects != null)
                        {
                            strPC = oEvent.SelectedObjects.GetValue("Code", 0);
                            if (strPC.Trim().Length > 0)
                            {
                                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC300");
                                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("DocEntry", oDBDS.Offset));
                                objLegalizationFormCache = new PCLegalizationFormCache();
                                objLegalizationFormCache.DocEntry = intDocEntry;
                                objLegalizationFormCache.pettyCash = getPCInfo(strPC);
                                CacheManager.CacheManager.Instance.addToCache(pVal.FormUID + "_" + intDocEntry.ToString(), objLegalizationFormCache, CacheManager.CacheManager.objCachePriority.Default);

                                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC300");
                                if (oDBDS != null)
                                {
                                    oDBDS.SetValue("U_VALUE", oDBDS.Offset, objLegalizationFormCache.pettyCash.AvailableValue.ToString(CultureInfo.InvariantCulture));
                                }

                            }
                        }
                    }
                }
                else
                {
                    objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    objCondition = objConditions.Add();
                    objCondition.Alias = "U_ISPOSTED";
                    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    objCondition.CondVal = "Y";
                    objCFL.SetConditions(objConditions);
                }

            }
            catch (COMException comEx)
            {


                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        private static PC getPCInfo(string strPCCode)
        {
            PC objPC = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;
            SAPbobsCOM.GeneralData objPCInfo = null;
            SAPbobsCOM.GeneralService objPCObject = null;

            SAPbobsCOM.GeneralDataCollection objThirdPartyLines = null;
            SAPbobsCOM.GeneralData objExpenseThirdPartyInfo = null;

            try
            {
                objPC = new PC();

                objPC.Code = strPCCode;
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objPCObject = objCompanyService.GetGeneralService(Settings._MainPettyCash.pettyCashUDO);
                objFilter = objPCObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("Code", strPCCode);
                objPCInfo = objPCObject.GetByParams(objFilter);
                objPC.AvailableValue = objPCInfo.GetProperty("U_AVVALUE");
                objPC.ControlAccount = objPCInfo.GetProperty("U_CTRLACCT");
                objPC.isPosted = objPCInfo.GetProperty("U_ISPOSTED") == "Y" ? true : false; ;
                objPC.Name = objPCInfo.GetProperty("Name");
                objPC.PCAccount = objPCInfo.GetProperty("U_PCACCOUNT");
                objPC.Remark = objPCInfo.GetProperty("U_REMARK");
                objPC.ThirdParty = objPCInfo.GetProperty("U_TERRELA");
                objPC.Value = objPCInfo.GetProperty("U_VALUE");

                            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objPC = null;
            }
            return objPC;
        }

        private static PCLegalizationFormCache getPCLegalizationInformation(int intLegalizationDocEntry, string PCCode, string FormUID, bool ForceLoad)
        {
            PCLegalizationFormCache objLegalizationFormCache = null;

            bool blLoadLegalizationInformation = false;

            try
            {
                objLegalizationFormCache = CacheManager.CacheManager.Instance.getFromCache(FormUID + "_" + intLegalizationDocEntry.ToString());
                if (objLegalizationFormCache != null)
                {
                    if (objLegalizationFormCache.pettyCash.Code.Trim() != PCCode)
                    {
                        blLoadLegalizationInformation = true;
                    }
                }
                else
                {
                    blLoadLegalizationInformation = true;
                }

                if (ForceLoad)
                {
                    blLoadLegalizationInformation = ForceLoad;
                }
                if (blLoadLegalizationInformation)
                {
                    objLegalizationFormCache = getPCLegalizationInfo(intLegalizationDocEntry);
                    CacheManager.CacheManager.Instance.addToCache(FormUID + "_" + intLegalizationDocEntry.ToString(), objLegalizationFormCache, CacheManager.CacheManager.objCachePriority.Default);
                }

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objLegalizationFormCache = null;
            }
            return objLegalizationFormCache;
        }

        private static PCLegalizationFormCache getPCLegalizationInfo(int intLegalizationDocEntry)
        {
            PCLegalizationFormCache objLegalizationFormCache = new PCLegalizationFormCache();
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;

            SAPbobsCOM.GeneralData objPCLegalizationInfo = null;
            SAPbobsCOM.GeneralService objPCLegalizationObject = null;

            SAPbobsCOM.GeneralDataCollection objPCLegalizationLines = null;
            SAPbobsCOM.GeneralData objPCLegalizationLineInfo = null;

            try
            {
                objLegalizationFormCache.DocEntry = intLegalizationDocEntry;
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objPCLegalizationObject = objCompanyService.GetGeneralService(Settings._MainPettyCash.pettyCashLegalizationUDO);
                objFilter = objPCLegalizationObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("DocEntry", intLegalizationDocEntry);
                objPCLegalizationInfo = objPCLegalizationObject.GetByParams(objFilter);
                objLegalizationFormCache.DocNum = objPCLegalizationInfo.GetProperty("DocNum");
                objLegalizationFormCache.Series = objPCLegalizationInfo.GetProperty("Series");
                objLegalizationFormCache.DocumentLines = new List<PCConceptLines>();
                objLegalizationFormCache.ExternalDocuments = new List<PCExternalDocuments>();
                objLegalizationFormCache.isPosted = objPCLegalizationInfo.GetProperty("U_ISPOSTED") == "Y" ? true : false;
                objLegalizationFormCache.JournalEntries = new List<PCJournalEntries>();
                objLegalizationFormCache.PCValue = objPCLegalizationInfo.GetProperty("U_VALUE");
                objLegalizationFormCache.pettyCash = new PC();
                string strPTCode = objPCLegalizationInfo.GetProperty("U_PTCODE");
                objLegalizationFormCache.pettyCash = getPCInfo(strPTCode.Trim());
                objLegalizationFormCache.remarks = objPCLegalizationInfo.GetProperty("Remark");
                objLegalizationFormCache.TotalValue = objPCLegalizationInfo.GetProperty("U_TOTVALUE");


                objPCLegalizationLines = objPCLegalizationInfo.Child("BYB_T1PTC301");
                for (int i = 0; i < objPCLegalizationLines.Count; i++)
                {
                    objPCLegalizationLineInfo = objPCLegalizationLines.Item(i);
                    PCConceptLines objPCConceptLine = new PCConceptLines();
                    string strConcept = objPCLegalizationLineInfo.GetProperty("U_CONCEPT");
                    if (strConcept.Trim().Length > 0)
                    {
                        
                        objPCConceptLine.ConceptCode = strConcept.Trim();
                        objPCConceptLine.Concept = getConceptInformation(strConcept);
                        objPCConceptLine.Date = objPCLegalizationLineInfo.GetProperty("U_DATE");
                        objPCConceptLine.Description = objPCLegalizationLineInfo.GetProperty("U_DESCRIPTION");
                        objPCConceptLine.DIM1 = objPCLegalizationLineInfo.GetProperty("U_DIM1");
                        objPCConceptLine.DIM2 = objPCLegalizationLineInfo.GetProperty("U_DIM2");
                        objPCConceptLine.DIM3 = objPCLegalizationLineInfo.GetProperty("U_DIM3");
                        objPCConceptLine.DIM4 = objPCLegalizationLineInfo.GetProperty("U_DIM4");
                        objPCConceptLine.DIM5 = objPCLegalizationLineInfo.GetProperty("U_DIM5");
                        objPCConceptLine.LineTotal = objPCLegalizationLineInfo.GetProperty("U_LINETOT");
                        objPCConceptLine.ProfitCenter = objPCLegalizationLineInfo.GetProperty("U_DIM1");
                        objPCConceptLine.Project = objPCLegalizationLineInfo.GetProperty("U_PROJECT");
                        objPCConceptLine.ThirdParty = objPCLegalizationLineInfo.GetProperty("U_TERRELA");
                        objPCConceptLine.TotalBeforeTaxes = objPCLegalizationLineInfo.GetProperty("U_VALUE");
                        objPCConceptLine.VAT = objPCLegalizationLineInfo.GetProperty("U_VAT");
                        objPCConceptLine.WHTax = objPCLegalizationLineInfo.GetProperty("U_WHTAX");
                        objLegalizationFormCache.DocumentLines.Add(objPCConceptLine);
                    }
                }



            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objLegalizationFormCache = null;
            }
            return objLegalizationFormCache;
        }

        public static PCConcept getConceptInformation(string strCode)
        {
            PCConcept objConcept = null;
            SAPbobsCOM.CompanyService objService = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;

            SAPbobsCOM.GeneralService objPCConceptObject = null;
            SAPbobsCOM.GeneralData objPCConceptInfo = null;

            SAPbobsCOM.GeneralDataCollection objWHTaxChild = null;
            SAPbobsCOM.GeneralDataCollection objThirdpartyChild = null;

            string strDescription = "";
            List<string> WHTaxCodeList = null;
            List<string> VATCodeList = null;

            try
            {
                objService = MainObject.Instance.B1Company.GetCompanyService();
                objPCConceptObject = objService.GetGeneralService(Settings._MainPettyCash.pettyCashConceptUDO);
                objFilter = objPCConceptObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("Code", strCode);
                objPCConceptInfo = objPCConceptObject.GetByParams(objFilter);

                objConcept = new PCConcept();
                
                objConcept.Account = objPCConceptInfo.GetProperty("U_ACCTCODE");
                objConcept.Code = objPCConceptInfo.GetProperty("Code");
                objConcept.Name = objPCConceptInfo.GetProperty("Name");
                objConcept.ValidThirdparty = new List<PCConceptThirdParty>();
                objConcept.ValidWHTax = new List<PCConceptWHTax>();
                objConcept.VATCode = objPCConceptInfo.GetProperty("U_TAXCODE");
                

                objThirdpartyChild = objPCConceptInfo.Child("BYB_T1PTC202");
                for (int i = 0; i < objThirdpartyChild.Count; i++)
                {
                    PCConceptThirdParty objTP = new PCConceptThirdParty();
                    objTP.Code = objThirdpartyChild.Item(i).GetProperty("U_RELPAR");
                    objTP.Default = Convert.ToString(objThirdpartyChild.Item(i).GetProperty("U_ISDEFAULT")) == "Y" ? true : false;
                    objConcept.ValidThirdparty.Add(objTP);
                }

                objWHTaxChild = objPCConceptInfo.Child("BYB_T1PTC201");
                for (int i = 0; i < objWHTaxChild.Count; i++)
                {
                    string strWCode = objWHTaxChild.Item(i).GetProperty("U_WTCODE");
                    if (strWCode.Trim().Length > 0)
                    {
                        PCConceptWHTax objWHtax = new PCConceptWHTax();
                        objWHtax.Code = objWHTaxChild.Item(i).GetProperty("U_WTCODE");
                        objConcept.ValidWHTax.Add(objWHtax);
                    }
                }
            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                objConcept = null;
            }
            catch (Exception ex)
            {
                _Logger.Error("", ex);
                objConcept = null;
            }
            return objConcept;
        }

        public static void setValidConceptInfo(SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;

            string strType = "";
            SAPbobsCOM.Recordset objRS = null;
            PCLegalizationFormCache objLegalizationFormCache = null;
            string strSQL = "";
            SAPbouiCOM.DBDataSource oDBDS = null;
            int intDocEntry = -1;
            string strPCCode = "";
            string strConceptCode = "";

            SAPbouiCOM.ChooseFromListEvent oEvent = null;
            SAPbouiCOM.DBDataSource oDBLinesDS = null;

            BubbleEvent = true;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC300");
                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("DocEntry", oDBDS.Offset));
                strPCCode = oDBDS.GetValue("U_PTCODE", oDBDS.Offset).Trim();
                if (strPCCode.Trim().Length > 0)
                {
                    objLegalizationFormCache = getPCLegalizationInformation(intDocEntry, strPCCode, pVal.FormUID, false);
                }
                else
                {
                    MainObject.Instance.B1Application.MessageBox("Por favor seleccione la caja menor que desea legalizar antes de selccionar los conceptos.");
                    BubbleEvent = false;
                    objLegalizationFormCache = null;
                }

                if (objLegalizationFormCache != null)
                {
                    objCFL = objForm.ChooseFromLists.Item("CFL_CON");
                    if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        oEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (oEvent.SelectedObjects != null)
                        {
                            strConceptCode = oEvent.SelectedObjects.GetValue("Code", 0);
                            PCConcept objConcept = getConceptInformation(strConceptCode);
                            if (objConcept != null)
                            {

                                PCConceptLines objConceptLine = new PCConceptLines();
                                objConceptLine.Concept = objConcept;
                                objConceptLine.ConceptCode = objConcept.Code;
                                objConceptLine.Description = objConcept.Name;
                                objConceptLine.FormLineNum = pVal.Row;
                                
                                if (objLegalizationFormCache.DocumentLines == null)
                                {
                                    objLegalizationFormCache.DocumentLines = new List<PCConceptLines>();
                                    objLegalizationFormCache.DocumentLines.Add(objConceptLine);
                                }
                                else
                                {

                                    if (objLegalizationFormCache.DocumentLines.Count == 0)
                                    {
                                        objLegalizationFormCache.DocumentLines.Add(objConceptLine);
                                    }
                                    else
                                    {
                                        bool blFound = false;
                                        for (int i = 0; i < objLegalizationFormCache.DocumentLines.Count; i++)
                                        {
                                            PCConceptLines tempLine = objLegalizationFormCache.DocumentLines[i];
                                            if (tempLine.FormLineNum == pVal.Row)
                                            {
                                                blFound = true;
                                                objLegalizationFormCache.DocumentLines[i] = objConceptLine;
                                                break;
                                            }

                                        }
                                        if (!blFound)
                                        {
                                            objLegalizationFormCache.DocumentLines.Add(objConceptLine);
                                        }
                                    }
                                }
                                CacheManager.CacheManager.Instance.addToCache(pVal.FormUID + "_" + intDocEntry.ToString(), objLegalizationFormCache, CacheManager.CacheManager.objCachePriority.Default);

                                oDBLinesDS = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC301");
                                oDBLinesDS.SetValue("U_DESCRIPTION", oDBLinesDS.Offset, objConcept.Name);
                                oDBLinesDS.SetValue("U_VALUE", oDBLinesDS.Offset, "0");
                                oDBLinesDS.SetValue("U_DATE", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_TERRELA", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_WHTAX", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_VAT", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_LINETOT", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_PROJECT", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_DIM1", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_DIM2", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_DIM3", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_DIM4", oDBLinesDS.Offset, "");
                                oDBLinesDS.SetValue("U_DIM5", oDBLinesDS.Offset, "");

                            }
                        }
                    }
                }

            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void filterThirdPartyConceptsUDO(SAPbouiCOM.ItemEvent pVal, bool clearFilter, out bool BubbleEvent)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;

            string strType = "";
            SAPbobsCOM.Recordset objRS = null;
            PCLegalizationFormCache objLegalizationFormCache = null;
            string strSQL = "";
            SAPbouiCOM.DBDataSource oDBDS = null;
            int intDocEntry = -1;
            string intExpenseDocEntry = "";
            string strConceptCode = "";
            BubbleEvent = true;

            SAPbouiCOM.ChooseFromListEvent oEvent = null;
            SAPbouiCOM.DBDataSource oDBLinesDS = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC300");
                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("DocEntry", oDBDS.Offset));
                intExpenseDocEntry = oDBDS.GetValue("U_PTCODE", oDBDS.Offset).Trim();
                if (intExpenseDocEntry.Trim().Length > 0)
                {
                    objLegalizationFormCache = getPCLegalizationInformation(intDocEntry, intExpenseDocEntry, pVal.FormUID, false);
                }

                if (objLegalizationFormCache != null)
                {
                    objCFL = objForm.ChooseFromLists.Item("CFL_RP");
                    if (clearFilter)
                    {
                        objCFL.SetConditions(null);
                    }
                    else
                    {
                        for (int i = 0; i < objLegalizationFormCache.DocumentLines.Count; i++)
                        {
                            if (objLegalizationFormCache.DocumentLines[i].FormLineNum == pVal.Row)
                            {
                                bool TPFilterFound = false;
                                objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                                int intTotal = objLegalizationFormCache.DocumentLines[i].Concept.ValidThirdparty.Count;
                                int intCounter = intTotal;
                                foreach (PCConceptThirdParty oTP in objLegalizationFormCache.DocumentLines[i].Concept.ValidThirdparty)
                                {
                                    if (oTP.Code.Trim().Length > 0)
                                    {
                                        TPFilterFound = true;
                                        string strCode = oTP.Code.Trim();
                                        objCondition = objConditions.Add();
                                        if (intTotal > 1)
                                        {
                                            objCondition.BracketOpenNum = 1;
                                        }

                                        objCondition.Alias = "Code";
                                        objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                        objCondition.CondVal = strCode;
                                        intCounter--;
                                        if (intTotal > 1)
                                        {
                                            objCondition.BracketCloseNum = 1;
                                            if (intCounter != 0)
                                            {
                                                objCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        intCounter--;
                                    }

                                }
                                if (TPFilterFound)
                                {
                                    objCFL.SetConditions(objConditions);
                                }
                                break;
                            }
                        }
                    }
                }
            }
            catch (COMException comEx)
            {


                _Logger.Error("", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void getPCLegalizationDocEntryOnLoad(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            SAPbouiCOM.Form objForm = null;
            int intDocEntry = -1;
            string strPCCode = "";
            SAPbouiCOM.DBDataSource oDBDS = null;
            SAPbouiCOM.UserDataSource oUDS = null;


            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC300");
                strPCCode = oDBDS.GetValue("U_PTCODE", oDBDS.Offset).Trim();
                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset));

                PCLegalizationFormCache objFormCacheInfo = getPCLegalizationInformation(intDocEntry, strPCCode, BusinessObjectInfo.FormUID, false);

            }
            catch (COMException comEX)
            {
                _Logger.Error("", comEX);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void refreshFormValues(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DBDataSource objDBDS = null;
            SAPbouiCOM.DBDataSource objDBDSLines = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;
            PCLegalizationFormCache objLegalizationFormCache = null;
            int intDocEntry = -1;
            string strPCCode = "";
            double dbRowValue = 0;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC300");
                intDocEntry = Convert.ToInt32(objDBDS.GetValue("DocEntry", objDBDS.Offset) == "" ? "0" : objDBDS.GetValue("DocEntry", objDBDS.Offset));
                strPCCode = objDBDS.GetValue("U_PTCODE", objDBDS.Offset).Trim();
                if (strPCCode.Trim().Length > 0)
                {
                    objLegalizationFormCache = getPCLegalizationInformation(intDocEntry, strPCCode, pVal.FormUID, false);
                }

                if (objLegalizationFormCache != null)
                {

                    objMatrix = objForm.Items.Item(pVal.ItemUID).Specific;
                    objMatrix.FlushToDataSource();
                    objDBDSLines = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC301");
                    if (!Double.TryParse(objDBDSLines.GetValue("U_VALUE", objDBDSLines.Offset), out dbRowValue))
                    {
                        dbRowValue = 0;
                    }
                    if (dbRowValue > 0)
                    {
                        double dbVATTotal = 0;
                        double dbWTTAX = 0;
                        for (int i = 0; i < objLegalizationFormCache.DocumentLines.Count; i++)
                        {
                            PCConceptLines objLines = objLegalizationFormCache.DocumentLines[i];
                            if (objLines.FormLineNum == pVal.Row)
                            {
                                objLegalizationFormCache.DocumentLines[i].TotalBeforeTaxes = dbRowValue;
                                #region calculateVAT (128)

                                
                                    SAPbobsCOM.SalesTaxCodes objVAT = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesTaxCodes);
                                    if (objVAT.GetByKey(objLegalizationFormCache.DocumentLines[i].Concept.VATCode))
                                    {
                                        double dbPercent = objVAT.Rate;
                                        dbVATTotal += dbRowValue * (dbPercent / 100);

                                    }
                                
                                objLegalizationFormCache.DocumentLines[i].VAT = dbVATTotal;

                                #endregion


                                #region calculateWHT (178)

                                foreach (PCConceptWHTax objConcetpWHT in objLines.Concept.ValidWHTax)
                                {
                                    SAPbobsCOM.WithholdingTaxCodes objWHT = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                                    if (objWHT.GetByKey(objConcetpWHT.Code))
                                    {
                                        double dbPercent = objWHT.BaseAmount;
                                        dbWTTAX += dbRowValue * (dbPercent / 100);

                                    }
                                }
                                objLegalizationFormCache.DocumentLines[i].WHTax = dbWTTAX;
                                #endregion
                                objLegalizationFormCache.DocumentLines[i].LineTotal = dbRowValue + dbVATTotal - dbWTTAX;
                                objLegalizationFormCache.TotalValue += objLegalizationFormCache.DocumentLines[i].LineTotal;

                                CacheManager.CacheManager.Instance.addToCache(pVal.FormUID + "_" + intDocEntry.ToString(), objLegalizationFormCache, CacheManager.CacheManager.objCachePriority.Default);



                                objDBDSLines.SetValue("U_VALUE", objDBDSLines.Offset, Convert.ToString(objLegalizationFormCache.DocumentLines[i].TotalBeforeTaxes.ToString(CultureInfo.InvariantCulture)));

                                objDBDSLines.SetValue("U_WHTAX", objDBDSLines.Offset, Convert.ToString(objLegalizationFormCache.DocumentLines[i].WHTax.ToString(CultureInfo.InvariantCulture)));
                                objDBDSLines.SetValue("U_VAT", objDBDSLines.Offset, Convert.ToString(objLegalizationFormCache.DocumentLines[i].VAT.ToString(CultureInfo.InvariantCulture)));
                                objDBDSLines.SetValue("U_LINETOT", objDBDSLines.Offset, Convert.ToString(objLegalizationFormCache.DocumentLines[i].LineTotal.ToString(CultureInfo.InvariantCulture)));
                                objDBDS.SetValue("U_TOTVALUE", objDBDS.Offset, Convert.ToString(objLegalizationFormCache.TotalValue, CultureInfo.InvariantCulture));

                                break;
                            }

                        }
                    }
                    objMatrix.LoadFromDataSource();


                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void postLegalization(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.DBDataSource objDBDS = null;
            int intLegalizationDocEntry = -1;
            string strPCCode = "";

            PCLegalizationFormCache objFormCache = null;

            SAPbobsCOM.GeneralDataParams oFilter = null;

            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objLegalizationObject = null;
            SAPbobsCOM.GeneralData objLegalizationInfo = null;

            
            
            SAPbobsCOM.JournalEntries objAdditionalJE = null;
            SAPbobsCOM.JournalEntries objMainJE = null;

            List<SAPbobsCOM.JournalEntries> objAllJE = null;
            List<int> JETransIdListE = null;

            //Incluir un campo para la fecha de contabilizacion general diferente a la fecha actual
            DateTime dtPostingDate = DateTime.Now;

            //Double totalCreditSCMain = 0;
            //Double totalDebitSCMain = 0;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    objDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1PTC300");
                    intLegalizationDocEntry = Convert.ToInt32(objDBDS.GetValue("DocEntry", objDBDS.Offset));
                    strPCCode = objDBDS.GetValue("U_PTCODE", objDBDS.Offset).Trim();

                    objFormCache = getPCLegalizationInformation(intLegalizationDocEntry, strPCCode, pVal.FormUID, true);


                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    objLegalizationObject = objCompanyService.GetGeneralService(Settings._MainPettyCash.pettyCashLegalizationUDO);
                    oFilter = objLegalizationObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oFilter.SetProperty("DocEntry", intLegalizationDocEntry);
                    objLegalizationInfo = objLegalizationObject.GetByParams(oFilter);

                    if (objLegalizationInfo != null && objFormCache != null)
                    {
                        
                        
                        //Incluir un campo de proyecto general en la caja menor
                        //objJE.ProjectCode = objFormCache.expense.expenseCostAccounting.Project;

                        objMainJE = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        objMainJE.ReferenceDate = dtPostingDate;
                        objMainJE.TaxDate = dtPostingDate;
                        objMainJE.DueDate = dtPostingDate;
                        objMainJE.Reference = intLegalizationDocEntry.ToString();
                        objMainJE.Memo = "Legalizacion de Caja Menor: " + strPCCode;
                        if (Settings._Main.LegalizationTransactionCode.Trim().Length > 0)
                        {
                            objMainJE.TransactionCode = Settings._MainPettyCash.PCLegalizationTransactionCode.Trim();
                        }
                        objMainJE.AutomaticWT = SAPbobsCOM.BoYesNoEnum.tYES;
                        objMainJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES;
                        
                        bool blFirstMain = true;
                        double dbPostingTotal = 0;
                        string strOutMessage = "";

                        //Dictionary<string, double> objWT = new Dictionary<string, double>();
                        foreach (PCConceptLines objLines in objFormCache.DocumentLines)
                        {
                            if (objLines.ConceptCode.Trim().Length > 0)
                            {
                                if (objLines.Concept.ValidWHTax.Count > 0)
                                {
                                    Double totalCreditSCAdditional = 0;
                                    Double totalDebitSCAdditional = 0;

                                    //double dbTotal = objLines.TotalBeforeTaxes + objLines.VAT;

                                    objAdditionalJE = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                                    objAdditionalJE.ReferenceDate = dtPostingDate;
                                    objAdditionalJE.TaxDate = dtPostingDate;
                                    objAdditionalJE.DueDate = dtPostingDate;
                                    objAdditionalJE.Reference = intLegalizationDocEntry.ToString();
                                    objAdditionalJE.Memo = "Legalizacion de Caja Menor: " + strPCCode;
                                    if (Settings._Main.LegalizationTransactionCode.Trim().Length > 0)
                                    {
                                        objAdditionalJE.TransactionCode = Settings._MainPettyCash.PCLegalizationTransactionCode.Trim();
                                    }
                                    objAdditionalJE.AutomaticWT = SAPbobsCOM.BoYesNoEnum.tYES;
                                    objAdditionalJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES;
                                    double dbTotal = objLines.TotalBeforeTaxes + objLines.VAT - objLines.WHTax;
                                    objAdditionalJE.Lines.ShortName = objLines.ThirdParty;
                                    objAdditionalJE.Lines.Credit = dbTotal;
                                    double dbCreditSys = T1.B1.Base.DIOperations.Operations.getSCValue(dbTotal, dtPostingDate, out strOutMessage, SAPbobsCOM.RoundingContextEnum.rcSum);
                                    objAdditionalJE.Lines.CreditSys = dbCreditSys;
                                    totalCreditSCAdditional = dbCreditSys;
                                    objAdditionalJE.Lines.ProjectCode = objLines.Project;
                                    objAdditionalJE.Lines.CostingCode = objLines.DIM1;
                                    objAdditionalJE.Lines.CostingCode2 = objLines.DIM2;
                                    objAdditionalJE.Lines.CostingCode3 = objLines.DIM3;
                                    objAdditionalJE.Lines.CostingCode4 = objLines.DIM4;
                                    objAdditionalJE.Lines.CostingCode5 = objLines.DIM5;

                                    objAdditionalJE.Lines.Add();
                                    objAdditionalJE.Lines.AccountCode = objLines.Concept.Account;
                                    objAdditionalJE.Lines.Debit = objLines.TotalBeforeTaxes;
                                    double dbSys = T1.B1.Base.DIOperations.Operations.getSCValue(objLines.TotalBeforeTaxes, dtPostingDate, out strOutMessage, SAPbobsCOM.RoundingContextEnum.rcLineGrossTotal);
                                    objAdditionalJE.Lines.DebitSys = dbSys;
                                    totalDebitSCAdditional = dbSys;
                                    objAdditionalJE.Lines.DebitSys = dbSys;
                                    if (objLines.Concept.VATCode.Trim().Length > 0)
                                    {
                                        objAdditionalJE.Lines.TaxPostAccount = SAPbobsCOM.BoTaxPostAccEnum.tpa_PurchaseTaxAccount;
                                        objAdditionalJE.Lines.TaxCode = objLines.Concept.VATCode.Trim();
                                        totalDebitSCAdditional += T1.B1.Base.DIOperations.Operations.getSCValue(T1.B1.Base.DIOperations.Operations.getTaxAmountLC(objLines.Concept.VATCode.Trim(),objLines.TotalBeforeTaxes), dtPostingDate, out strOutMessage, SAPbobsCOM.RoundingContextEnum.rcTax);
                                    }
                                    objAdditionalJE.Lines.WTLiable = SAPbobsCOM.BoYesNoEnum.tYES;
                                    objAdditionalJE.Lines.ProjectCode = objLines.Project;
                                    objAdditionalJE.Lines.CostingCode = objLines.DIM1;
                                    objAdditionalJE.Lines.CostingCode2 = objLines.DIM2;
                                    objAdditionalJE.Lines.CostingCode3 = objLines.DIM3;
                                    objAdditionalJE.Lines.CostingCode4 = objLines.DIM4;
                                    objAdditionalJE.Lines.CostingCode5 = objLines.DIM5;

                                    bool blFirstWT = true;
                                    foreach (PCConceptWHTax oWTCode in objLines.Concept.ValidWHTax)
                                    {
                                        SAPbobsCOM.WithholdingTaxCodes objWHT = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                                        if (objWHT.GetByKey(oWTCode.Code))
                                        {
                                            if (!blFirstWT)
                                            {
                                                objAdditionalJE.WithholdingTaxData.Add();
                                            }
                                            double dbWTTAX = T1.B1.Base.DIOperations.Operations.getWTAmountLC(oWTCode.Code, objLines.TotalBeforeTaxes);
                                            double dbSySWTTax = T1.B1.Base.DIOperations.Operations.getSCValue(dbWTTAX, dtPostingDate, out strOutMessage, SAPbobsCOM.RoundingContextEnum.rcWTax);
                                            totalCreditSCAdditional += dbSySWTTax;

                                            //double dbPercent = objWHT.BaseAmount;
                                            //double dbWTTAX = objLines.TotalBeforeTaxes * (dbPercent / 100);
                                            objAdditionalJE.WithholdingTaxData.WTCode = oWTCode.Code;
                                            //objAdditionalJE.WithholdingTaxData.TaxableAmount = dbWTTAX;
                                            //objAdditionalJE.WithholdingTaxData.TaxableAmountinSys = dbSySWTTax;
                                            //objAdditionalJE.WithholdingTaxData.WTAmount = dbWTTAX;
                                            //objAdditionalJE.WithholdingTaxData.WTAmountSys = dbSySWTTax;

                                            blFirstWT = false;
                                        }
                                    }
                                    double dbDifference = totalCreditSCAdditional - totalDebitSCAdditional;
                                                                        
                                    if(dbDifference > 0)
                                    {
                                        objAdditionalJE.Lines.Add();
                                        objAdditionalJE.Lines.AccountCode = Settings._MainPettyCash.SysCurrDeviationAccount;
                                        objAdditionalJE.Lines.Debit = 0;
                                        objAdditionalJE.Lines.DebitSys = Math.Abs(dbDifference);
                                    }
                                    else if (dbDifference < 0)
                                    {
                                        objAdditionalJE.Lines.Add();
                                        objAdditionalJE.Lines.AccountCode = Settings._MainPettyCash.SysCurrDeviationAccount;
                                        objAdditionalJE.Lines.Credit = 0;
                                        objAdditionalJE.Lines.CreditSys = Math.Abs(dbDifference);

                                    }

                                    if (objAllJE == null )
                                    {
                                        objAllJE = new List<SAPbobsCOM.JournalEntries>(); 
                                    }

                                    objAllJE.Add(objAdditionalJE);

                                }
                                else
                                {

                                    if (!blFirstMain)
                                    {
                                        objMainJE.Lines.Add();
                                    }
                                    if (blFirstMain)
                                    {
                                        if (objFormCache.pettyCash.ThirdParty.Trim().Length > 0)
                                        {
                                            objMainJE.Lines.ShortName = objFormCache.pettyCash.ThirdParty.Trim();
                                            objMainJE.Lines.ControlAccount = objFormCache.pettyCash.ControlAccount;
                                        }
                                        else
                                        {
                                            objMainJE.Lines.AccountCode = objFormCache.pettyCash.ControlAccount;
                                            objMainJE.Lines.Credit = 0;
                                        }
                                        objMainJE.Lines.Add();
                                        objMainJE.Lines.AccountCode = objLines.Concept.Account;
                                        objMainJE.Lines.Debit = objLines.TotalBeforeTaxes;
                                        blFirstMain = false;
                                    }
                                    else
                                    {
                                        objMainJE.Lines.AccountCode = objLines.Concept.Account;
                                        objMainJE.Lines.Debit = objLines.TotalBeforeTaxes;
                                    }
    
                                            dbPostingTotal += objLines.TotalBeforeTaxes;
                                        objMainJE.Lines.ProjectCode = objLines.Project;
                                        objMainJE.Lines.CostingCode = objLines.DIM1;
                                        objMainJE.Lines.CostingCode2 = objLines.DIM2;
                                        objMainJE.Lines.CostingCode3 = objLines.DIM3;
                                        objMainJE.Lines.CostingCode4 = objLines.DIM4;
                                        objMainJE.Lines.CostingCode5 = objLines.DIM5;
                                        
                                            if (objLines.Concept.VATCode.Trim().Length > 0)
                                            {
                                            objMainJE.Lines.TaxPostAccount = SAPbobsCOM.BoTaxPostAccEnum.tpa_PurchaseTaxAccount;
                                        objMainJE.Lines.TaxCode = objLines.Concept.VATCode.Trim();
                                            dbPostingTotal += objLines.VAT;
                                            }
                                            

                                        
                                    
                                }
                            }



                            
                        }
                        
                        objMainJE.Lines.SetCurrentLine(0);
                        objMainJE.Lines.Credit = dbPostingTotal;

                        if(objAllJE == null)
                        {
                            objAllJE = new List<SAPbobsCOM.JournalEntries>();
                        }
                        objAllJE.Add(objMainJE);

                        if (MainObject.Instance.B1Company.InTransaction)
                        {
                            MainObject.Instance.B1Application.MessageBox("La base de datos se encuentra bloqueada por otra transacción. Por favor inténtelo en unos minutos");
                        }
                        else
                        {
                            MainObject.Instance.B1Company.StartTransaction();
                            JETransIdListE = new List<int>();
                            bool blResult = false;
                            foreach (SAPbobsCOM.JournalEntries objJEItem in objAllJE)
                            {
                                int intResult = -1;
                                string strXML = objJEItem.GetAsXML();
                                intResult = objJEItem.Add();
                                if (intResult == 0)
                                {
                                    int intTemp = Convert.ToInt32(MainObject.Instance.B1Company.GetNewObjectKey());
                                    SAPbobsCOM.JournalEntries oTemp = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                                    oTemp.GetByKey(intTemp);
                                    string strTemp = oTemp.GetAsXML();
                                    JETransIdListE.Add(Convert.ToInt32(MainObject.Instance.B1Company.GetNewObjectKey()));
                                    blResult = true;
                                }
                                else
                                {
                                    string strMessage = MainObject.Instance.B1Company.GetLastErrorDescription();

                                    _Logger.Error(strMessage);
                                    MainObject.Instance.B1Application.SetStatusBarMessage(strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    blResult = false;
                                    break;
                                }
                            }
                            if (blResult)
                            {

                                //objLegalizationInfo.SetProperty("U_isPosted", "Y");
                                //objLegalizationInfo.SetProperty("U_postDate", objFormCache.PostingDate);
                                //objLegalizationInfo.SetProperty("U_JEEntry", Convert.ToInt32(MainObject.Instance.B1Company.GetNewObjectKey(), CultureInfo.InvariantCulture));
                                //objLegalizationInfo.SetProperty("U_TOTVALUE", dbPostingTotal);
                                //objLegalizationObject.Update(objLegalizationInfo);
                                MainObject.Instance.B1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            }
                                objForm.Refresh();
                            
                        }

                    }
                }
                else
                {
                    MainObject.Instance.B1Application.MessageBox("Por favor actualice o cree el documento de legalización antes de contabilizar.");
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                if (MainObject.Instance.B1Company.InTransaction)
                {
                    MainObject.Instance.B1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + er.Message,SAPbouiCOM.BoMessageTime.bmt_Short,true);

            }
        }

        #endregion




    }
}
