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
    public class Expenses
    {
        private static Expenses objExpenses;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private Expenses()
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

        #region Concepts
        public static void openConceptsForm(SAPbouiCOM.MenuEvent pVal)
        {
            if (Expenses.objExpenses == null)
            {
                Expenses.objExpenses = new T1.B1.Expenses.Expenses();
            }
            
            try
            {
                MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "BYB_T1EXP100", "");
                
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
                objCFL = objForm.ChooseFromLists.Item("CFL_COA");
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

        public static void loadBPNameCFL(SAPbouiCOM.Form objForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                SAPbouiCOM.ChooseFromListEvent objCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                SAPbouiCOM.DataTable objDT = objCFLEvent.SelectedObjects;
                
                
                if (objDT != null && objDT.Rows.Count > 0)
                {
                    SAPbouiCOM.DBDataSource variable5 = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP101");
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

        #endregion

        #region Expense Type

        public static void openExpenseTypeForm(SAPbouiCOM.MenuEvent pVal)
        {
            if (Expenses.objExpenses == null)
            {
                Expenses.objExpenses = new T1.B1.Expenses.Expenses();
            }

            try
            {
                MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "BYB_T1EXP200", "");

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

        public static void filterAccountExpTypeUDO(SAPbouiCOM.ItemEvent pVal, bool clearFilter)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_COA");
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

        #region Request

        

        public static void openExpenseRequestForm(SAPbouiCOM.MenuEvent pVal)
        {
            if (Expenses.objExpenses == null)
            {
                Expenses.objExpenses = new T1.B1.Expenses.Expenses();
            }

            try
            {
                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "BYB_T1EXP600", "");
                configureExpenseRequestFOrm(CacheManager.CacheManager.Instance.getFromCache(Settings._Main.ExpenseRequestFormLastId),true);
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

        public static void configureExpenseRequestFOrm(string FormUID, bool ReloadCombos)
        {
            if (Expenses.objExpenses == null)
            {
                Expenses.objExpenses = new T1.B1.Expenses.Expenses();
            }

            bool isMultiDim = false;
            List<dimensionInfo> objDimList = null;
            List<costCenterInfo> objProfitCenterList = null;
            List<projectInfo> objProjectList = null;
            SAPbouiCOM.ComboBox objCombo = null;
            SAPbouiCOM.EditText objEdit = null;
            SAPbouiCOM.StaticText objStatic = null;
            SAPbouiCOM.Item objItem = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.DBDataSource oDS = null;
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
                        objCombo = objForm.Items.Item("Item_2").Specific;
                        objCombo.ValidValues.LoadSeries("BYB_T1EXP600", SAPbouiCOM.BoSeriesMode.sf_Add);


                        objEdit = objForm.Items.Item("1_U_E").Specific;
                        strNextNumber = objForm.BusinessObject.GetNextSerialNumber(objCombo.Value).ToString();
                        objEdit.Value = strNextNumber;
                    }

                        #endregion
                    

                    #region control Status
                    if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        objItem = objForm.Items.Item("26_U_E");
                        objItem.Enabled = true;
                        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (Settings._Main.DefaultCreateStatus.Length > 0)
                            {
                                oDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP600");
                                oDS.SetValue("U_STATUS", oDS.Offset, Settings._Main.DefaultCreateStatus);
                            }
                        }
                        
                            objItem.Enabled = false;
                        
                    }
                    else
                    {
                        objItem = objForm.Items.Item("26_U_E");
                        objItem.Enabled = true;

                    }
                    #endregion
                    if (ReloadCombos)
                    {
                        isMultiDim = isMultipleDimension();
                        if (isMultiDim)
                        {
                            objDimList = getDimensionsList();

                        }


                        objProfitCenterList = getProfitCenterList();

                        objProjectList = getProjectList();


                        #region load DropDowns Dimensions and ProfitCenter

                        #region ProfitCenter
                        if (!isMultiDim)
                        {
                            objCombo = objForm.Items.Item("Item_6").Specific;
                            for (int i = 0; i < objCombo.ValidValues.Count; i++)
                            {
                                objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }

                            foreach (costCenterInfo oPC in objProfitCenterList)
                            {
                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                            }
                            objItem = objForm.Items.Item("Item_6");
                            objItem.Visible = true;
                            objItem = objForm.Items.Item("25_U_S");
                            objItem.Visible = true;


                        }
                        #endregion

                        #region Dimensions
                        if (isMultiDim)
                        {

                            foreach (dimensionInfo oDIM in objDimList)
                            {
                                switch (oDIM.DimentionCode)
                                {
                                    case 1:
                                        if (oDIM.isActive)
                                        {
                                            objCombo = objForm.Items.Item("Item_7").Specific;
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
                                            objStatic = objForm.Items.Item("27_U_S").Specific;
                                            objStatic.Caption = oDIM.DimensionName;
                                            objItem = objForm.Items.Item("Item_7");
                                            objItem.Visible = true;
                                            objItem = objForm.Items.Item("27_U_S");
                                            objItem.Visible = true;
                                        }
                                        break;
                                    case 2:
                                        if (oDIM.isActive)
                                        {
                                            objCombo = objForm.Items.Item("Item_8").Specific;
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
                                            objStatic = objForm.Items.Item("28_U_S").Specific;
                                            objStatic.Caption = oDIM.DimensionName;
                                            objItem = objForm.Items.Item("Item_8");
                                            objItem.Visible = true;
                                            objItem = objForm.Items.Item("28_U_S");
                                            objItem.Visible = true;
                                        }
                                        break;
                                    case 3:
                                        if (oDIM.isActive)
                                        {
                                            objCombo = objForm.Items.Item("Item_9").Specific;
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
                                            objStatic = objForm.Items.Item("29_U_S").Specific;
                                            objStatic.Caption = oDIM.DimensionName;
                                            objItem = objForm.Items.Item("Item_9");
                                            objItem.Visible = true;
                                            objItem = objForm.Items.Item("29_U_S");
                                            objItem.Visible = true;
                                        }
                                        break;
                                    case 4:
                                        if (oDIM.isActive)
                                        {
                                            objCombo = objForm.Items.Item("Item_10").Specific;
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
                                            objStatic = objForm.Items.Item("30_U_S").Specific;
                                            objStatic.Caption = oDIM.DimensionName;
                                            objItem = objForm.Items.Item("Item_10");
                                            objItem.Visible = true;
                                            objItem = objForm.Items.Item("30_U_S");
                                            objItem.Visible = true;
                                        }
                                        break;
                                    case 5:
                                        if (oDIM.isActive)
                                        {
                                            objCombo = objForm.Items.Item("Item_11").Specific;
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
                                            objStatic = objForm.Items.Item("31_U_S").Specific;
                                            objStatic.Caption = oDIM.DimensionName;
                                            objItem = objForm.Items.Item("Item_11");
                                            objItem.Visible = true;
                                            objItem = objForm.Items.Item("31_U_S");
                                            objItem.Visible = true;
                                        }
                                        break;
                                }
                            }








                        }
                        #endregion

                        #endregion
                        #region Projects
                        objCombo = objForm.Items.Item("Item_0").Specific;
                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                        {
                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }

                        foreach (projectInfo oPC in objProjectList)
                        {
                            objCombo.ValidValues.Add(oPC.ProjectCode, oPC.ProjectName);

                        }
                        #endregion
                    }
                    //#region Assign Default Values
                    //if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    //{
                    //    if (Settings._Main.DefaultCreateStatus.Length > 0)
                    //    {
                    //        oDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP600");
                    //        oDS.SetValue("U_STATUS", oDS.Offset, Settings._Main.DefaultCreateStatus);
                    //    }
                    //}
                    //#endregion
                    
                    objForm.PaneLevel = objForm.PaneLevel;
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
                if(objForm != null)
                {
                    objForm.Freeze(false);
                }
                CacheManager.CacheManager.Instance.removeFromCache(Settings._Main.ExpenseRequestFormLastId);
            }
        }

        public static void filterStepTypeUDO(SAPbouiCOM.ItemEvent pVal, bool clearFilter)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_STEP");
                if (clearFilter)
                {
                    objCFL.SetConditions(null);
                }
                else
                {
                    objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    objCondition = objConditions.Add();
                    objCondition.Alias = "U_STEPTYPE";
                    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    objCondition.CondVal = "SOL";
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

        public static void filterStatusUDO(SAPbouiCOM.ItemEvent pVal, bool clearFilter)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_STAT");
                if (clearFilter)
                {
                    objCFL.SetConditions(null);
                }
                else
                {
                    objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    objCondition = objConditions.Add();
                    objCondition.Alias = "U_STATYPE";
                    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    objCondition.CondVal = "SOL";
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

        #region Expense Report
        public void getExpenses(SAPbouiCOM.ItemEvent pVal, string StatusFilter)
        {
            try
            {

            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }
        #endregion

        #region Approve Request

        static public void loadPendingAppovedRequestForm()
        {
            string strSQL = "";
            SAPbobsCOM.Recordset objRS = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.FormCreationParams objParams = null;
            SAPbouiCOM.DataTable objDT = null;
            //SAPbouiCOM.Item objItem = null;
            SAPbouiCOM.Grid objGrid = null;
            SAPbouiCOM.GridColumn objGridColumn = null;
            SAPbouiCOM.EditTextColumn oEditTExt = null;


            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objParams.XmlData = ExpensesRes.BYB_ApproveRequestsForm;
                objParams.FormType = "BYB_REQAPR";
                objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objDT = objForm.DataSources.DataTables.Item("BYB_RELIST");


                objRS = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName))
                {
                    strSQL = Settings._HANA.getApprovedRequests;
                }
                else
                {
                    strSQL = Settings._SQL.getApprovedRequests;
                }

                objDT.ExecuteQuery(strSQL);

                #region Format Grid
                objGrid = objForm.Items.Item("aprGrid").Specific;

                objGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;


                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;

                objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                oEditTExt.LinkedObjectType = "BYB_T1EXP600";
                
                


                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;


                objGridColumn = objGrid.Columns.Item(2);
                objGridColumn.Editable = false;

                objGridColumn = objGrid.Columns.Item(3);
                objGridColumn.Editable = false;

                objGridColumn = objGrid.Columns.Item(4);
                objGridColumn.Editable = false;

                objGridColumn.RightJustified = true;





                objGrid.AutoResizeColumns();


                #endregion


                objForm.Visible = true;

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

        static public void updateRequestStatus(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable objDT = null;
            
            SAPbouiCOM.Form objForm = null;
            
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objGeneralService = null;
            SAPbobsCOM.GeneralData objGeneralData = null;
            
            SAPbobsCOM.GeneralDataParams objSearch = null;
            


            

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objDT = objForm.DataSources.DataTables.Item("BYB_RELIST");

                if (objDT.Rows.Count > 0)
                {
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    objGeneralService = objCompanyService.GetGeneralService("BYB_T1EXP600");
                    objGeneralData = objGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    
                    T1.B1.Base.UIOperations.Operations.startProgressBar("Procesando...", objDT.Rows.Count);
                    for (int i = 0; i < objDT.Rows.Count; i++)
                    {
                        int strDocEntry = objDT.GetValue(1, i);
                        string strChecked = objDT.GetValue(0, i);
                        if (strChecked.Trim().ToUpper() == "Y")
                        {
                            objSearch = objGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            objSearch.SetProperty("DocEntry", strDocEntry);

                            objGeneralData = objGeneralService.GetByParams(objSearch);
                            if (objGeneralData != null)
                            {
                                objGeneralData.SetProperty("U_STATUS", Settings._Main.ApprovedStatusValue);
                                objGeneralService.Update(objGeneralData);
                            }
                        }
                        T1.B1.Base.UIOperations.Operations.setProgressBarMessage("Documento " + strDocEntry + " procesado.", i + 1);
                    }
                    T1.B1.Base.UIOperations.Operations.stopProgressBar();
                    objForm.Close();
                    MainObject.Instance.B1Application.MessageBox("El proceso finalizó con éxito");
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
            finally
            {
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
                if (objForm != null)
                {
                    objForm.Freeze(false);
                }
            }
        }
        #endregion

        #region Desembolso

        static public void loadPaymentForm()
        {
            string strSQL = "";
            SAPbobsCOM.Recordset objRS = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.FormCreationParams objParams = null;
            SAPbouiCOM.DataTable objDT = null;
            //SAPbouiCOM.Item objItem = null;
            SAPbouiCOM.Grid objGrid = null;
            SAPbouiCOM.GridColumn objGridColumn = null;
            SAPbouiCOM.EditTextColumn oEditTExt = null;


            try
            {
                objParams = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                objParams.XmlData = ExpensesRes.BYBEXP_Payment;
                objParams.FormType = Settings._Main.PaymentFormType.Trim();
                objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objDT = objForm.DataSources.DataTables.Item("DT_DOCS");


                objRS = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName))
                {
                    strSQL = Settings._HANA.getValidPaymentDocuments;
                }
                else
                {
                    strSQL = Settings._SQL.getValidPaymentDocuments;
                }

                strSQL = strSQL.Replace("[--ValidStatus--]", Settings._Main.PaymentValidValues);

                objDT.ExecuteQuery(strSQL);

                #region Format Grid
                objGrid = objForm.Items.Item("gridDocs").Specific;

                objGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;


                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;

                objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                oEditTExt.LinkedObjectType = "BYB_T1EXP600";

                
                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;


                objGridColumn = objGrid.Columns.Item(2);
                objGridColumn.Editable = false;

                objGridColumn = objGrid.Columns.Item(3);
                objGridColumn.Editable = false;

                objGridColumn = objGrid.Columns.Item(4);
                objGridColumn.Editable = false;

                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(5);
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(6);
                objGridColumn.RightJustified = true;

                objGridColumn = objGrid.Columns.Item(7);
                objGridColumn.Editable = false;



                objGrid.AutoResizeColumns();


                #endregion


                objForm.Visible = true;

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

        public static void filterAccountPaymentForm(SAPbouiCOM.ItemEvent pVal, bool clearFilter)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.UserDataSource oDS = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_ACCT");
                if (clearFilter)
                {
                    objCFL.SetConditions(null);
                    SAPbouiCOM.ChooseFromListEvent oEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    if(oEvent.SelectedObjects != null && oEvent.SelectedObjects.Rows.Count == 1)
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

        public static void addPaymentDocument(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable objDocs = null;
            SAPbouiCOM.Form objForm = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objRequest = null;
            SAPbobsCOM.GeneralService objRequestType = null;
            SAPbobsCOM.GeneralService objRelatedParties = null;
            SAPbobsCOM.GeneralData objRequestData = null;
            SAPbobsCOM.GeneralData objRequestTypeData = null;
            SAPbobsCOM.GeneralData objRelatedPartiesData = null;
            SAPbobsCOM.GeneralData objRequestRelatedPartiesData = null;
            SAPbobsCOM.GeneralDataParams objFilterParams = null;
            SAPbobsCOM.GeneralDataCollection objRequestRelatedParties = null;

            SAPbouiCOM.UserDataSource objPaymentDate = null;
            SAPbouiCOM.UserDataSource objCashAccount = null;
            SAPbouiCOM.UserDataSource objPaymentRemarks = null;

            SAPbobsCOM.Payments objPayment = null;
            bool blResultPayment = false;

            #region DocumentValues
            int intRequestDocEntry = -1;
            string strChecked = "N";
            string strProject = "";
            string strDIM1 = "";
            string strDIM2 = "";
            string strDIM3 = "";
            string strDIM4 = "";
            string strDIM5 = "";
            string strRelatedPartiesCode = "";
            string strCardCode = "";
            string strExpenseType = "";
            string strRequestRemark = "";
            double dbRequestedValue = 0;
            string strCashAccount = "";
            string strRequestTypeAccount = "";
            DateTime dtPaymentDate;
            string strPaymentRemarks = "";
            int intResult = -1;
            string strMessage = "";
            string strCardName = "";
            int intPayNumber = -1;
            
            
            #endregion

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objDocs = objForm.DataSources.DataTables.Item("DT_DOCS");
                if(objDocs != null)
                {
                    #region Get Form information
                    objPaymentDate = objForm.DataSources.UserDataSources.Item("UD_DATE");
                    SAPbobsCOM.SBObob objTemp = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                    SAPbobsCOM.Recordset objRS = objTemp.Format_StringToDate(objPaymentDate.Value);
                    dtPaymentDate = objRS.Fields.Item(0).Value;// Convert.ToDateTime(objPaymentDate.Value,CultureInfo.InvariantCulture);
                    objCashAccount = objForm.DataSources.UserDataSources.Item("UD_ACCT");
                    strCashAccount = objCashAccount.Value.Trim();
                    objPaymentRemarks = objForm.DataSources.UserDataSources.Item("UD_MEMO");
                    strPaymentRemarks = objPaymentRemarks.Value.Trim();
                    #endregion


                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    objRequest = objCompanyService.GetGeneralService(Settings._Main.ExpenseUDO);
                    for (int i=0; i < objDocs.Rows.Count;i++)
                    {
                        strChecked = objDocs.GetValue(0, i);
                        intRequestDocEntry = objDocs.GetValue(1, i);
                        if(strChecked.Trim() == "Y")
                        {
                            objFilterParams = objRequest.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            objFilterParams.SetProperty("DocEntry", intRequestDocEntry);
                            objRequestData = objRequest.GetByParams(objFilterParams);
                            if(objRequestData != null)
                            {
                                strExpenseType = objRequestData.GetProperty("U_EXPTYPE");
                                strProject = objRequestData.GetProperty("U_PROJECT");
                                strDIM1 = objRequestData.GetProperty("U_DIM1");
                                strDIM2 = objRequestData.GetProperty("U_DIM2");
                                strDIM3 = objRequestData.GetProperty("U_DIM3");
                                strDIM4 = objRequestData.GetProperty("U_DIM4");
                                strDIM5 = objRequestData.GetProperty("U_DIM5");
                                strRequestRemark = objRequestData.GetProperty("Remark");
                                dbRequestedValue = objRequestData.GetProperty("U_VALUE");
                                int strNumber = objDocs.GetValue(6, i);
                                if (strNumber > 0)
                                {
                                    intPayNumber = strNumber;
                                }
                                else
                                {
                                    intPayNumber = -1;
                                }
                                DateTime objPayDate = objDocs.GetValue(5, i) == null? new DateTime(2010,1,1): objDocs.GetValue(5, i);
                                if (objPayDate > new DateTime(2015, 1, 1))
                                {
                                    dtPaymentDate = objPayDate;
                                }

                                objRequestRelatedParties = objRequestData.Child(Settings._Main.ExpenseRequestRelatedPartiesChild);
                                #region Retrieve First RelatedParty Code
                                for (int j=0; j < objRequestRelatedParties.Count; j++)
                                {
                                    objRequestRelatedPartiesData = objRequestRelatedParties.Item(j);
                                    strRelatedPartiesCode = objRequestRelatedPartiesData.GetProperty("U_TERRELA");
                                    if(strRelatedPartiesCode.Trim().Length > 0)
                                    {
                                        break;
                                    }
                                }
                                #endregion

                                #region Get Related Party CardCode

                                if(strRelatedPartiesCode.Trim().Length > 0)
                                {
                                    objRelatedParties = objCompanyService.GetGeneralService(Settings._Main.RelatedPartyUDO);
                                    objFilterParams = objRelatedParties.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                    objFilterParams.SetProperty("Code", strRelatedPartiesCode);
                                    objRelatedPartiesData = objRelatedParties.GetByParams(objFilterParams);
                                    if(objRelatedPartiesData != null)
                                    {
                                        strCardCode = objRelatedPartiesData.GetProperty("U_CARDCODE");
                                        strCardName = objRelatedPartiesData.GetProperty("U_LEGALNAME");
                                    }
                                }
                                #endregion

                                objRequestType = objCompanyService.GetGeneralService(Settings._Main.ExpenseTypeUDO);
                                objFilterParams = objRequestType.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                objFilterParams.SetProperty("Code", strExpenseType);
                                objRequestTypeData = objRequestType.GetByParams(objFilterParams);
                                if(objRequestTypeData != null)
                                {
                                    strRequestTypeAccount = objRequestTypeData.GetProperty("U_MAINACCT");
                                }
                                else
                                {
                                    _Logger.Error("BYB: Could not retrieve Expense Type Data: " + strExpenseType);
                                }

                                bool blIsAssociated = T1.B1.Base.DIOperations.Operations.isAccountAsociated(strRequestTypeAccount);
                                if (blIsAssociated)
                                {

                                    objPayment = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                                    objPayment.CardCode = strCardCode;
                                    objPayment.DocDate = dtPaymentDate;
                                    objPayment.ControlAccount = strRequestTypeAccount;
                                    objPayment.TransferAccount = strCashAccount;
                                    objPayment.TransferSum = dbRequestedValue;
                                    objPayment.TransferDate = dtPaymentDate;
                                    objPayment.TaxDate = dtPaymentDate;
                                    objPayment.DocType = SAPbobsCOM.BoRcptTypes.rSupplier;
                                    objPayment.ProjectCode = strProject;
                                    objPayment.Remarks = strRequestRemark + ". " + strPaymentRemarks;
                                    if(intPayNumber > 0)
                                    {
                                        objPayment.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES;
                                        objPayment.DocNum = intPayNumber;
                                    }
                                    
                                    
                                }
                                else
                                {
                                    objPayment = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                                    objPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
                                    objPayment.CardCode = strCardCode;
                                    objPayment.CardName = strCardName;
                                    objPayment.ProjectCode = strProject;
                                    objPayment.DocDate = dtPaymentDate;
                                    objPayment.AccountPayments.AccountCode = strRequestTypeAccount;
                                    objPayment.AccountPayments.Decription = "Desembolso legalizacion ";
                                    objPayment.AccountPayments.SumPaid = dbRequestedValue;
                                    objPayment.AccountPayments.ProfitCenter = strDIM1;
                                    objPayment.AccountPayments.ProfitCenter2 = strDIM2;
                                    objPayment.AccountPayments.ProfitCenter3 = strDIM3;
                                    objPayment.AccountPayments.ProfitCenter4 = strDIM4;
                                    objPayment.AccountPayments.ProfitCenter5 = strDIM5;
                                    objPayment.TransferAccount = strCashAccount;
                                    objPayment.TransferSum = dbRequestedValue;
                                    objPayment.TransferDate = dtPaymentDate;
                                    objPayment.TaxDate = dtPaymentDate;
                                    objPayment.Remarks = strRequestRemark + ". " + strPaymentRemarks;
                                    if (intPayNumber > 0)
                                    {
                                        objPayment.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES;
                                        objPayment.DocNum = intPayNumber;
                                        
                                    }
                                }

                                intResult = objPayment.Add();
                                if(intResult == 0)
                                {
                                    blResultPayment = true;
                                    if (blIsAssociated)
                                    {

                                        if (strDIM1.Trim().Length > 0
                                            || strDIM2.Trim().Length > 0
                                            || strDIM3.Trim().Length > 0
                                            || strDIM4.Trim().Length > 0
                                            || strDIM5.Trim().Length > 0
                                            )
                                        {
                                            string strLastKey = MainObject.Instance.B1Company.GetNewObjectKey();
                                            if (objPayment.GetByKey(Convert.ToInt32(strLastKey)))
                                            {
                                                string strSQL = "";
                                                objRS = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                if (CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName))
                                                {
                                                    strSQL = Settings._HANA.getJEFromPayment;
                                                }
                                                else
                                                {
                                                    strSQL = Settings._SQL.getJEFromPayment;
                                                }

                                                strSQL = strSQL.Replace("[--DocEntry--]", strLastKey);
                                                objRS.DoQuery(strSQL);
                                                if (objRS.RecordCount > 0)
                                                {
                                                    int TransId = objRS.Fields.Item(0).Value;
                                                    SAPbobsCOM.JournalEntries objJE = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                                                    if (objJE.GetByKey(TransId))
                                                    {
                                                        for (int k = 0; k < objJE.Lines.Count; k++)
                                                        {
                                                            objJE.Lines.SetCurrentLine(k);
                                                            objJE.Lines.CostingCode = strDIM1;
                                                            objJE.Lines.CostingCode2 = strDIM2;
                                                            objJE.Lines.CostingCode3 = strDIM3;
                                                            objJE.Lines.CostingCode4 = strDIM4;
                                                            objJE.Lines.CostingCode5 = strDIM5;

                                                        }
                                                        intResult = objJE.Update();
                                                        if (intResult != 0)
                                                        {
                                                            _Logger.Error(MainObject.Instance.B1Company.GetLastErrorDescription());
                                                        }

    ;
                                                    }
                                                }

                                            }
                                            else
                                            {
                                                _Logger.Error("Could not retrieve Payment document. The payment was created but the JE was not updated with Dimensions Information");
                                            }
                                        }
                                    }
                                    objRequestData.SetProperty("U_STATUS", Settings._Main.PaymentStatusValue);
                                    objRequest.Update(objRequestData);
                                    strMessage = "El desembolso se creó con éxito";
                                }
                                else
                                {
                                    strMessage = MainObject.Instance.B1Company.GetLastErrorDescription();
                                }


                            }
                            else
                            {
                                _Logger.Error("Could not get RequestData info DocEntry: " + intRequestDocEntry.ToString());
                            }
                        }
                    }
                    if(blResultPayment)
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage(strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        objForm.Close();
                    }
                    else
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage(strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                    
                }

            }
            catch (COMException comEx)
            {


                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage(comEx.InnerException.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage(er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Legalizar 

        public static void changeRequestStatus(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            SAPbobsCOM.CompanyService objCompanyService = null;


            SAPbobsCOM.GeneralService objRequestObject = null;
            SAPbobsCOM.GeneralData objRequestInfo = null;

            SAPbobsCOM.GeneralService objLegalizationObject = null;
            SAPbobsCOM.GeneralData objLegalizationInfo = null;

            SAPbobsCOM.GeneralDataParams objFilter = null;

            SAPbouiCOM.DBDataSource objDS = null;
            SAPbouiCOM.Form objForm = null;

            int intDocEntry = -1;
            int intRequestEntry = -1;

            string strKey = "";
            XmlDocument objXML = null;
            try
            {
                strKey = BusinessObjectInfo.ObjectKey;
                int intIndexStart = strKey.IndexOf("<DocEntry>");
                int intIndexEnd = strKey.IndexOf("</DocEntry>");
                intDocEntry = Convert.ToInt32(strKey.Substring(intIndexStart + 10, intIndexEnd - intIndexStart - 10));
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objLegalizationObject = objCompanyService.GetGeneralService(Settings._Main.LegalizationUDO);
                objFilter = objLegalizationObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("DocEntry", intDocEntry);
                objLegalizationInfo = objLegalizationObject.GetByParams(objFilter);
                intRequestEntry = Convert.ToInt32(objLegalizationInfo.GetProperty("U_EXPENSECODE"));

                objRequestObject = objCompanyService.GetGeneralService("BYB_T1EXP600");
                objFilter = objRequestObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("DocEntry", intRequestEntry);
                objRequestInfo = objRequestObject.GetByParams(objFilter);
                objRequestInfo.SetProperty("U_STATUS", "EN PROCESO");
                objRequestObject.Update(objRequestInfo);

            }
            catch(COMException comEx)
            {
                string strMessage = "";
                if (comEx.Message.Trim().Length == 0)
                {
                    strMessage = Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.InnerException.Message + "::" + comEx.StackTrace);
                }
                else
                {
                    strMessage = Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace);
                }
                Exception er = new Exception(strMessage);
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public static void deleteRowAfter(string FormUID)
        {
            
            SAPbouiCOM.DBDataSource oDBDS = null;
            SAPbouiCOM.DBDataSource oDBDSLines = null;
            SAPbouiCOM.Matrix objMatrix = null;
            SAPbouiCOM.Form objForm = null;
            int intDocEntry = -1;
            int intExpenseDocEntry = -1;
            LegalizationFormCache objFormCache = null;
            int intIndex = -1;
            int intRowNumber = -1;

            try
            {
                   
                objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                oDBDSLines = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP401");
                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("DocEntry", oDBDS.Offset));
                intExpenseDocEntry = Convert.ToInt32(oDBDS.GetValue("U_EXPENSECODE", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("U_EXPENSECODE", oDBDS.Offset));
                //MainObject.Instance.B1Application.SetStatusBarMessage(intExpenseDocEntry.ToString() + DateTime.Now.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                 objFormCache = getLegalizationInformation(intDocEntry, intExpenseDocEntry, objForm.UniqueID, false);

                if (objFormCache != null && objFormCache.DocumentLines != null)
                {
                    string lastRightClick = CacheManager.CacheManager.Instance.getFromCache("RightClickLastRow") == null ? "" : CacheManager.CacheManager.Instance.getFromCache("RightClickLastRow");
                    string[] strValues = lastRightClick.Split('#');
                    if (strValues[0] == FormUID)
                    {
                        if (strValues.Length == 2)
                        {
                            intRowNumber = Convert.ToInt32(strValues[1].Trim().Length == 0 ? -1 : Convert.ToInt32(strValues[1].Trim()));
                        }
                        if (intRowNumber < 0)
                        {
                            objMatrix = objForm.Items.Item("0_U_G").Specific;
                            SAPbouiCOM.CellPosition objCelPos = objMatrix.GetCellFocus();
                            if (objCelPos != null)
                            {
                                intRowNumber = objCelPos.rowIndex;
                            }
                            else
                            {
                                intRowNumber = -1;
                            }
                        }

                        #region oldCode

                        //objFormCache.DocumentLines = new List<ConceptLines>();
                        //MainObject.Instance.B1Application.SetStatusBarMessage(oDBDSLines.Size.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        //for (int i =0; i < oDBDSLines.Size; i++)
                        //{
                        //    ConceptLines objConcept = new ConceptLines();
                        //    double dbValue = 0;
                        //    string strConceptCode = oDBDSLines.GetValue("U_CONCEPT", i);
                        //    if (strConceptCode.Trim().Length > 0)
                        //    {
                        //        objConcept.ConceptCode = strConceptCode.Trim();
                        //        objConcept.Concept = getConceptInformation(objConcept.ConceptCode);
                        //        string strDate = oDBDSLines.GetValue("U_DATE", i);
                        //        if (strDate.Trim().Length > 0)
                        //        {
                        //            objConcept.Date = Convert.ToDateTime(DateTime.ParseExact(strDate.Trim(),"yyyyMMdd", CultureInfo.InvariantCulture));
                        //        }
                        //        objConcept.Description = oDBDSLines.GetValue("U_DESCRIPTION", i);
                        //        objConcept.DIM1 = oDBDSLines.GetValue("U_DIM1", i);
                        //        objConcept.DIM2 = oDBDSLines.GetValue("U_DIM2", i);
                        //        objConcept.DIM3 = oDBDSLines.GetValue("U_DIM3", i);
                        //        objConcept.DIM4 = oDBDSLines.GetValue("U_DIM4", i);
                        //        objConcept.DIM5 = oDBDSLines.GetValue("U_DIM5", i);
                        //        objConcept.FormLineNum = i+1;
                        //        dbValue = oDBDSLines.GetValue("U_LINETOT", i).Trim().Length == 0 ? 0 : Double.Parse(oDBDSLines.GetValue("U_LINETOT", i).Trim(), CultureInfo.InvariantCulture);
                        //        objConcept.LineTotal = dbValue;
                        //        objConcept.ProfitCenter = oDBDSLines.GetValue("U_DIM1", i);
                        //        objConcept.Project = oDBDSLines.GetValue("U_PROJECT", i);
                        //        objConcept.ThirdParty = oDBDSLines.GetValue("U_THIRDPARTY", i);
                        //        dbValue = oDBDSLines.GetValue("U_VALUE", i).Trim().Length == 0 ? 0 : Double.Parse(oDBDSLines.GetValue("U_VALUE", i).Trim(), CultureInfo.InvariantCulture);
                        //        objConcept.TotalBeforeTaxes = dbValue;
                        //        dbValue = oDBDSLines.GetValue("U_VAT", i).Trim().Length == 0 ? 0 : Double.Parse(oDBDSLines.GetValue("U_VAT", i).Trim(), CultureInfo.InvariantCulture);
                        //        objConcept.VAT = dbValue;
                        //        dbValue = oDBDSLines.GetValue("U_WHTAX", i).Trim().Length == 0 ? 0 : Double.Parse(oDBDSLines.GetValue("U_WHTAX", i).Trim(), CultureInfo.InvariantCulture);
                        //        objConcept.WHTax = dbValue;
                        //        objFormCache.DocumentLines.Add(objConcept);
                        //    }


                        //}

                        //recalcTotalValue(oDBDS, objFormCache, objForm.UniqueID, intDocEntry);
                        #endregion


                        if (intRowNumber > 0)
                        {



                            ConceptLines objLine = null;
                            for (int i = 0; i < objFormCache.DocumentLines.Count; i++)
                            {
                                objLine = objFormCache.DocumentLines[i];
                                if (objLine.FormLineNum == intRowNumber)
                                {
                                    intIndex = i;
                                }
                                else
                                {
                                    if (intRowNumber < objLine.FormLineNum)
                                    {
                                        objLine.FormLineNum--;
                                    }
                                }
                            }

                            if (intIndex >= 0)
                            {
                                objFormCache.DocumentLines.RemoveAt(intIndex);
                            }


                            recalcTotalValue(oDBDS, objFormCache, objForm.UniqueID, intDocEntry);
                        }
                        else
                        {
                            MainObject.Instance.B1Application.SetStatusBarMessage("BYB Error al eleimnar la fila de la legalización. Cierre el formulario sin guardar e intentelo de nuevo. ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        }

                    }
                    else
                    {
                        MainObject.Instance.B1Application.SetStatusBarMessage("BYB: No se pudo recuperar la información de la legalización. Por favor ciere el formulario e intentelo de nuevo. ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }
            }
            catch(Exception er)
            {
                _Logger.Error("",er);
            }
        }

        public static void openLegalizationForm(SAPbouiCOM.MenuEvent pVal)
        {
            if (Expenses.objExpenses == null)
            {
                Expenses.objExpenses = new T1.B1.Expenses.Expenses();
            }

            try
            {
                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "BYB_T1EXP400", "");
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

        public static void filterValidRequestUDO(SAPbouiCOM.ItemEvent pVal, bool clearFilter)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.ChooseFromListEvent oEvent = null;
            SAPbouiCOM.DBDataSource oDBDS = null;
            SAPbouiCOM.UserDataSource oUDS = null;
            int intDocEntry = -1;
            int intExpenseDocEntry = -1;
            LegalizationFormCache objLegalizationFormCache = null;
            double dbValue = 0;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objCFL = objForm.ChooseFromLists.Item("CFL_EXP");
                if (clearFilter)
                {
                    objCFL.SetConditions(null);
                    if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (oEvent.SelectedObjects != null)
                        {
                            intExpenseDocEntry = oEvent.SelectedObjects.GetValue("DocEntry", 0);
                            if (intExpenseDocEntry > 0)
                            {
                                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("DocEntry", oDBDS.Offset));
                                objLegalizationFormCache = new LegalizationFormCache();
                                objLegalizationFormCache.DocEntry = intDocEntry;
                                objLegalizationFormCache.expense = getExpenseInfo(intExpenseDocEntry);
                                CacheManager.CacheManager.Instance.addToCache(pVal.FormUID + "_" + intDocEntry.ToString(), objLegalizationFormCache,CacheManager.CacheManager.objCachePriority.Default);
                                
                                    oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                                    if (oDBDS != null)
                                    {
                                        oDBDS.SetValue("U_VALUE", oDBDS.Offset, objLegalizationFormCache.expense.expectedValue.ToString(CultureInfo.InvariantCulture));
                                    }
                                
                            }
                        }
                    }
                }
                else
                {
                    objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    objCondition = objConditions.Add();
                    objCondition.Alias = "U_STATUS";
                    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    objCondition.CondVal = Settings._Main.PaymentStatusValue;
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

        public static void filterValidConceptsUDO(SAPbouiCOM.ItemEvent pVal, bool clearFilter, out bool BubbleEvent)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            
            string strType = "";
            SAPbobsCOM.Recordset objRS = null;
            LegalizationFormCache objLegalizationFormCache = null;
            string strSQL = "";
            SAPbouiCOM.DBDataSource oDBDS = null;
            int intDocEntry = -1;
            int intExpenseDocEntry = -1;
            string strConceptCode = "";
            BubbleEvent = true;

            SAPbouiCOM.ChooseFromListEvent oEvent = null;
            SAPbouiCOM.DBDataSource oDBLinesDS = null;
            SAPbouiCOM.Matrix objMatrix = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset) == "" ? "0": oDBDS.GetValue("DocEntry", oDBDS.Offset));
                intExpenseDocEntry = Convert.ToInt32(oDBDS.GetValue("U_EXPENSECODE", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("U_EXPENSECODE", oDBDS.Offset));
                if (intExpenseDocEntry > 0)
                {
                    objLegalizationFormCache = getLegalizationInformation(intDocEntry,intExpenseDocEntry,pVal.FormUID, false);
                }
                else
                {
                    MainObject.Instance.B1Application.MessageBox("Por favor seleccione la solicitud de gastos que desea legalizar antes de selccionar los conceptos.");
                    BubbleEvent = false;
                    objLegalizationFormCache = null;
                }

                if (objLegalizationFormCache != null)
                {
                    objCFL = objForm.ChooseFromLists.Item("CFL_CON");
                    if (clearFilter)
                    {
                        objCFL.SetConditions(null);
                        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            oEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                            if (oEvent.SelectedObjects != null)
                            {
                                objMatrix = objForm.Items.Item("0_U_G").Specific;
                                objMatrix.FlushToDataSource();
                                strConceptCode = oEvent.SelectedObjects.GetValue("Code", 0);
                                Concept objConcept = getConceptInformation(strConceptCode);
                                if (objConcept != null)
                                {

                                    ConceptLines objConceptLine = new ConceptLines();
                                    objConceptLine.Concept = objConcept;
                                    objConceptLine.ConceptCode = objConcept.Code;
                                    objConceptLine.Description = objConcept.Description;
                                    objConceptLine.FormLineNum = pVal.Row;


                                    if (objLegalizationFormCache.DocumentLines == null)
                                    {
                                        objLegalizationFormCache.DocumentLines = new List<ConceptLines>();
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
                                                ConceptLines tempLine = objLegalizationFormCache.DocumentLines[i];
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
                                    int CurrentRow = pVal.Row - 1;
                                    oDBLinesDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP401");
                                    oDBLinesDS.SetValue("U_DESCRIPTION", CurrentRow, objConcept.Description);
                                    oDBLinesDS.SetValue("U_VALUE", CurrentRow, "0");
                                    oDBLinesDS.SetValue("U_DATE", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_THIRDPARTY", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_WHTAX", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_VAT", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_LINETOT", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_Project", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_DIM1", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_DIM2", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_DIM3", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_DIM4", CurrentRow, "");
                                    oDBLinesDS.SetValue("U_DIM5", CurrentRow, "");

                                    recalcTotalValue(oDBDS, objLegalizationFormCache, pVal.FormUID, intDocEntry);

                                } 
                            }
                        }
                    }
                    else
                    {
                        strType = objLegalizationFormCache.expense.expnseType.expenseClass;
                        if (strType.Trim().Length > 0)
                        {
                            objRS = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            if (CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName))
                            {
                                strSQL = Settings._HANA.getConceptsCode;
                            }
                            else
                            {
                                strSQL = Settings._SQL.getConceptsCode;
                            }
                            strSQL = strSQL.Replace("[--ExpType--]", strType);
                            objRS.DoQuery(strSQL);
                            if (objRS.RecordCount > 0)
                            {
                                objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

                                int intTotal = objRS.RecordCount;
                                int intCounter = intTotal;
                                while (!objRS.EoF)
                                {
                                    string steCode = objRS.Fields.Item("Code").Value;
                                    objCondition = objConditions.Add();
                                    if (intTotal > 1)
                                    {
                                        objCondition.BracketOpenNum = 1;
                                    }

                                    objCondition.Alias = "Code";
                                    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                    objCondition.CondVal = steCode;
                                    intCounter--;
                                    if (intTotal > 1)
                                    {
                                        objCondition.BracketCloseNum = 1;
                                        if (intCounter != 0)
                                        {
                                            objCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                                        }
                                    }

                                    objRS.MoveNext();

                                }
                                objCFL.SetConditions(objConditions);
                            }
                            else
                            {
                                objConditions = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

                                objCondition = objConditions.Add();

                                objCondition.Alias = "Code";
                                objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                objCondition.CondVal = "NULL-1-1-1";

                                objCFL.SetConditions(objConditions);

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
            LegalizationFormCache objLegalizationFormCache = null;
            string strSQL = "";
            SAPbouiCOM.DBDataSource oDBDS = null;
            int intDocEntry = -1;
            int intExpenseDocEntry = -1;
            string strConceptCode = "";
            BubbleEvent = true;

            SAPbouiCOM.ChooseFromListEvent oEvent = null;
            SAPbouiCOM.DBDataSource oDBLinesDS = null;

            SAPbouiCOM.Matrix objMatrix = null;
            SAPbouiCOM.CellPosition objCelPos = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("DocEntry", oDBDS.Offset));
                intExpenseDocEntry = Convert.ToInt32(oDBDS.GetValue("U_EXPENSECODE", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("U_EXPENSECODE", oDBDS.Offset));
                if (intExpenseDocEntry > 0)
                {
                    objLegalizationFormCache = getLegalizationInformation(intDocEntry, intExpenseDocEntry, pVal.FormUID, false);
                }

                if (objLegalizationFormCache != null)
                {
                    objCFL = objForm.ChooseFromLists.Item("CFL_RP");
                    if (clearFilter)
                    {
                        objCFL.SetConditions(null);
                        objMatrix = objForm.Items.Item(pVal.ItemUID).Specific;
                        objMatrix.FlushToDataSource();
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
                                foreach (ConceptThirdParty oTP in objLegalizationFormCache.DocumentLines[i].Concept.ValidThirdparty)
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

        public static void configureLegalizationForm(string FormUID, bool ReloadCombos)
        {
            if (Expenses.objExpenses == null)
            {
                Expenses.objExpenses = new T1.B1.Expenses.Expenses();
            }

            bool isMultiDim = false;
            List<dimensionInfo> objDimList = null;
            List<costCenterInfo> objProfitCenterList = null;
            List<projectInfo> objProjectList = null;
            SAPbouiCOM.ComboBox objCombo = null;
            SAPbouiCOM.EditText objEdit = null;
            SAPbouiCOM.StaticText objStatic = null;
            SAPbouiCOM.Item objItem = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.DBDataSource oDS = null;
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
                        objCombo = objForm.Items.Item("Item_2").Specific;
                        objCombo.ValidValues.LoadSeries("BYB_T1EXP400", SAPbouiCOM.BoSeriesMode.sf_Add);


                        objEdit = objForm.Items.Item("1_U_E").Specific;
                        strNextNumber = objForm.BusinessObject.GetNextSerialNumber(objCombo.Value).ToString();
                        objEdit.Value = strNextNumber;

                        loadMatrixComboBoxes(FormUID);
                    }

                    #endregion

                    
                    //if (ReloadCombos)
                    //{
                    //    isMultiDim = isMultipleDimension();
                    //    if (isMultiDim)
                    //    {
                    //        objDimList = getDimensionsList();

                    //    }


                    //    objProfitCenterList = getProfitCenterList();

                    //    objProjectList = getProjectList();


                    //    #region load DropDowns Dimensions and ProfitCenter

                    //    #region ProfitCenter
                    //    if (!isMultiDim)
                    //    {
                    //        objCombo = objForm.Items.Item("Item_6").Specific;
                    //        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                    //        {
                    //            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    //        }

                    //        foreach (costCenterInfo oPC in objProfitCenterList)
                    //        {
                    //            objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                    //        }
                    //        objItem = objForm.Items.Item("Item_6");
                    //        objItem.Visible = true;
                    //        objItem = objForm.Items.Item("25_U_S");
                    //        objItem.Visible = true;


                    //    }
                    //    #endregion

                    //    #region Dimensions
                    //    if (isMultiDim)
                    //    {

                    //        foreach (dimensionInfo oDIM in objDimList)
                    //        {
                    //            switch (oDIM.DimentionCode)
                    //            {
                    //                case 1:
                    //                    if (oDIM.isActive)
                    //                    {
                    //                        objCombo = objForm.Items.Item("Item_7").Specific;
                    //                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                    //                        {
                    //                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    //                        }

                    //                        foreach (costCenterInfo oPC in objProfitCenterList)
                    //                        {
                    //                            if (oPC.DimensionCode == 1)
                    //                            {
                    //                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                    //                            }
                    //                        }
                    //                        objStatic = objForm.Items.Item("27_U_S").Specific;
                    //                        objStatic.Caption = oDIM.DimensionName;
                    //                        objItem = objForm.Items.Item("Item_7");
                    //                        objItem.Visible = true;
                    //                        objItem = objForm.Items.Item("27_U_S");
                    //                        objItem.Visible = true;
                    //                    }
                    //                    break;
                    //                case 2:
                    //                    if (oDIM.isActive)
                    //                    {
                    //                        objCombo = objForm.Items.Item("Item_8").Specific;
                    //                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                    //                        {
                    //                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    //                        }

                    //                        foreach (costCenterInfo oPC in objProfitCenterList)
                    //                        {
                    //                            if (oPC.DimensionCode == 2)
                    //                            {
                    //                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                    //                            }
                    //                        }
                    //                        objStatic = objForm.Items.Item("28_U_S").Specific;
                    //                        objStatic.Caption = oDIM.DimensionName;
                    //                        objItem = objForm.Items.Item("Item_8");
                    //                        objItem.Visible = true;
                    //                        objItem = objForm.Items.Item("28_U_S");
                    //                        objItem.Visible = true;
                    //                    }
                    //                    break;
                    //                case 3:
                    //                    if (oDIM.isActive)
                    //                    {
                    //                        objCombo = objForm.Items.Item("Item_9").Specific;
                    //                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                    //                        {
                    //                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    //                        }

                    //                        foreach (costCenterInfo oPC in objProfitCenterList)
                    //                        {
                    //                            if (oPC.DimensionCode == 3)
                    //                            {
                    //                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                    //                            }
                    //                        }
                    //                        objStatic = objForm.Items.Item("29_U_S").Specific;
                    //                        objStatic.Caption = oDIM.DimensionName;
                    //                        objItem = objForm.Items.Item("Item_9");
                    //                        objItem.Visible = true;
                    //                        objItem = objForm.Items.Item("29_U_S");
                    //                        objItem.Visible = true;
                    //                    }
                    //                    break;
                    //                case 4:
                    //                    if (oDIM.isActive)
                    //                    {
                    //                        objCombo = objForm.Items.Item("Item_10").Specific;
                    //                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                    //                        {
                    //                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    //                        }

                    //                        foreach (costCenterInfo oPC in objProfitCenterList)
                    //                        {
                    //                            if (oPC.DimensionCode == 4)
                    //                            {
                    //                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                    //                            }
                    //                        }
                    //                        objStatic = objForm.Items.Item("30_U_S").Specific;
                    //                        objStatic.Caption = oDIM.DimensionName;
                    //                        objItem = objForm.Items.Item("Item_10");
                    //                        objItem.Visible = true;
                    //                        objItem = objForm.Items.Item("30_U_S");
                    //                        objItem.Visible = true;
                    //                    }
                    //                    break;
                    //                case 5:
                    //                    if (oDIM.isActive)
                    //                    {
                    //                        objCombo = objForm.Items.Item("Item_11").Specific;
                    //                        for (int i = 0; i < objCombo.ValidValues.Count; i++)
                    //                        {
                    //                            objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    //                        }

                    //                        foreach (costCenterInfo oPC in objProfitCenterList)
                    //                        {
                    //                            if (oPC.DimensionCode == 5)
                    //                            {
                    //                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                    //                            }
                    //                        }
                    //                        objStatic = objForm.Items.Item("31_U_S").Specific;
                    //                        objStatic.Caption = oDIM.DimensionName;
                    //                        objItem = objForm.Items.Item("Item_11");
                    //                        objItem.Visible = true;
                    //                        objItem = objForm.Items.Item("31_U_S");
                    //                        objItem.Visible = true;
                    //                    }
                    //                    break;
                    //            }
                    //        }








                    //    }
                    //    #endregion

                    //    #endregion
                    //    #region Projects
                    //    objCombo = objForm.Items.Item("Item_0").Specific;
                    //    for (int i = 0; i < objCombo.ValidValues.Count; i++)
                    //    {
                    //        objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    //    }

                    //    foreach (projectInfo oPC in objProjectList)
                    //    {
                    //        objCombo.ValidValues.Add(oPC.ProjectCode, oPC.ProjectName);

                    //    }
                    //    #endregion
                    //}
                    //#region Assign Default Values
                    //if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    //{
                    //    if (Settings._Main.DefaultCreateStatus.Length > 0)
                    //    {
                    //        oDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP600");
                    //        oDS.SetValue("U_STATUS", oDS.Offset, Settings._Main.DefaultCreateStatus);
                    //    }
                    //}
                    //#endregion

                    //objForm.PaneLevel = objForm.PaneLevel;
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

        public static double getExpenseValue(int DocEntry)
        {
            double dbTotal = 0;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralData objExpenseInfo = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;
            SAPbobsCOM.GeneralService objExpenseObject = null;
            try
            {
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objExpenseObject = objCompanyService.GetGeneralService(Settings._Main.ExpenseUDO);
                objFilter = objExpenseObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("DocEntry", DocEntry);
                objExpenseInfo = objExpenseObject.GetByParams(objFilter);
                dbTotal = objExpenseInfo.GetProperty("U_VALUE");


            }
            catch(COMException comEx)
            {
                _Logger.Error("", comEx);
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
            return dbTotal;
        }

        public static Tuple<string,string> getExpenseTypeInfo(int DocEntry)
        {
            Tuple<string, string> expInfo = new Tuple<string, string>("", "");
            string strExpenseType = "";
            string strConceptType = "";
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralData objExpenseInfo = null;
            SAPbobsCOM.GeneralData objExpenseTypeInfo = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;
            SAPbobsCOM.GeneralService objExpenseObject = null;
            SAPbobsCOM.GeneralService objExpenseTypeObject = null;
            try
            {
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objExpenseObject = objCompanyService.GetGeneralService(Settings._Main.ExpenseUDO);
                objFilter = objExpenseObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("DocEntry", DocEntry);
                objExpenseInfo = objExpenseObject.GetByParams(objFilter);
                strExpenseType = objExpenseInfo.GetProperty("U_EXPTYPE");

                objExpenseTypeObject = objCompanyService.GetGeneralService(Settings._Main.ExpenseTypeUDO);
                objFilter = objExpenseTypeObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("Code", strExpenseType);
                objExpenseTypeInfo = objExpenseTypeObject.GetByParams(objFilter);
                strConceptType = objExpenseTypeInfo.GetProperty("U_EXPTYPE");
                expInfo = new Tuple<string, string>(strExpenseType, strConceptType);







            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            return expInfo;
        }

        public static void getLegalizationDocEntryOnLoad(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            SAPbouiCOM.Form objForm = null;
            int intDocEntry = -1;
            int intExpenseDocEntry = -1;
            SAPbouiCOM.DBDataSource oDBDS = null;
            SAPbouiCOM.UserDataSource oUDS = null;


            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                intExpenseDocEntry = Convert.ToInt32(oDBDS.GetValue("U_EXPENSECODE", oDBDS.Offset));
                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset));

                LegalizationFormCache objFormCacheInfo = getLegalizationInformation(intDocEntry, intExpenseDocEntry, BusinessObjectInfo.FormUID, false);

            }
            catch(COMException comEX)
            {
                _Logger.Error("", comEX);
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }

        private static LegalizationFormCache getLegalizationInfo(int intLegalizationDocEntry)
        {
            LegalizationFormCache objLegalizationFormCache = new LegalizationFormCache();
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;

            SAPbobsCOM.GeneralData objLegalizationInfo = null;
            SAPbobsCOM.GeneralService objLegalizationObject = null;
            SAPbobsCOM.GeneralDataCollection objLegalizationLines = null;
            SAPbobsCOM.GeneralData objLegalizationLineInfo = null;

            


            try
            {
                objLegalizationFormCache.DocEntry = intLegalizationDocEntry;
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objLegalizationObject = objCompanyService.GetGeneralService(Settings._Main.LegalizationUDO);
                objFilter = objLegalizationObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("DocEntry", intLegalizationDocEntry);
                objLegalizationInfo = objLegalizationObject.GetByParams(objFilter);
                objLegalizationFormCache.DocNum = objLegalizationInfo.GetProperty("DocNum");
                objLegalizationFormCache.Series = objLegalizationInfo.GetProperty("Series");
                objLegalizationFormCache.ConceptList = new List<Concept>();

                int intExpenseDocEntry = Convert.ToInt32(objLegalizationInfo.GetProperty("U_EXPENSECODE"));


                objLegalizationFormCache.expense = getExpenseInfo(intExpenseDocEntry);
                objLegalizationFormCache.ExpenseValue = objLegalizationInfo.GetProperty("U_VALUE");
                objLegalizationFormCache.DocumentLines = new List<ConceptLines>();
                objLegalizationFormCache.isPosted = objLegalizationInfo.GetProperty("U_ISPOSTED") == "Y" ? true : false;
                objLegalizationFormCache.TotalValue = objLegalizationInfo.GetProperty("U_TOTVALUE");
                objLegalizationFormCache.PostingDate = objLegalizationInfo.GetProperty("U_POSTDATE");
                objLegalizationFormCache.JournalEntry = objLegalizationInfo.GetProperty("U_JEENTRY");
                objLegalizationFormCache.remarks = objLegalizationInfo.GetProperty("Remark");


                objLegalizationLines = objLegalizationInfo.Child("BYB_T1EXP401");
                for(int i=0; i < objLegalizationLines.Count; i++)
                {
                    objLegalizationLineInfo = objLegalizationLines.Item(i);
                    ConceptLines objConceptLine = new ConceptLines();
                    string strConcept = objLegalizationLineInfo.GetProperty("U_CONCEPT");
                    if (strConcept.Trim().Length > 0)
                    {
                        objConceptLine.ConceptCode = objLegalizationLineInfo.GetProperty("U_CONCEPT");
                        objConceptLine.Date = objLegalizationLineInfo.GetProperty("U_DATE");
                        objConceptLine.Description = objLegalizationLineInfo.GetProperty("U_DESCRIPTION");
                        objConceptLine.DIM1 = objLegalizationLineInfo.GetProperty("U_DIM1");
                        objConceptLine.DIM2 = objLegalizationLineInfo.GetProperty("U_DIM2");
                        objConceptLine.DIM3 = objLegalizationLineInfo.GetProperty("U_DIM3");
                        objConceptLine.DIM4 = objLegalizationLineInfo.GetProperty("U_DIM4");
                        objConceptLine.DIM5 = objLegalizationLineInfo.GetProperty("U_DIM5");
                        objConceptLine.LineTotal = objLegalizationLineInfo.GetProperty("U_LINETOT");
                        objConceptLine.ProfitCenter = objLegalizationLineInfo.GetProperty("U_DIM1");
                        objConceptLine.Project = objLegalizationLineInfo.GetProperty("U_PROJECT");
                        objConceptLine.ThirdParty = objLegalizationLineInfo.GetProperty("U_THIRDPARTY");
                        objConceptLine.TotalBeforeTaxes = objLegalizationLineInfo.GetProperty("U_VALUE");
                        objConceptLine.VAT = objLegalizationLineInfo.GetProperty("U_VAT");
                        objConceptLine.WHTax = objLegalizationLineInfo.GetProperty("U_WHTAX");
                        objConceptLine.Concept = getConceptInformation(objConceptLine.ConceptCode);
                        objConceptLine.FormLineNum = i+1;

                        objLegalizationFormCache.DocumentLines.Add(objConceptLine);
                    }
                }

                
                
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                objLegalizationFormCache = null;
            }
            return objLegalizationFormCache;
        }
        private static Expense getExpenseInfo(int intExpenseDocEntry)
        {
            Expense objExpense = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;
            SAPbobsCOM.GeneralData objExpenseInfo = null;
            SAPbobsCOM.GeneralService objExpenseObject = null;
            SAPbobsCOM.GeneralDataCollection objThirdPartyLines = null;
            SAPbobsCOM.GeneralData objExpenseThirdPartyInfo = null;

            try
            {
                objExpense = new Expense();

                objExpense.docEntry = intExpenseDocEntry;
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objExpenseObject = objCompanyService.GetGeneralService(Settings._Main.ExpenseUDO);
                objFilter = objExpenseObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("DocEntry", intExpenseDocEntry);
                objExpenseInfo = objExpenseObject.GetByParams(objFilter);
                objExpense.docNum = objExpenseInfo.GetProperty("DocNum");
                objExpense.endDate = objExpenseInfo.GetProperty("U_ENDDATE");
                objExpense.expectedValue = objExpenseInfo.GetProperty("U_VALUE");

                ExpenseCostAccounting oCostAcc = new ExpenseCostAccounting();
                oCostAcc.CostCenter = objExpenseInfo.GetProperty("U_COSTCENTER");
                oCostAcc.DIM1 = objExpenseInfo.GetProperty("U_DIM1");
                oCostAcc.DIM2 = objExpenseInfo.GetProperty("U_DIM2");
                oCostAcc.DIM3 = objExpenseInfo.GetProperty("U_DIM3");
                oCostAcc.DIM4 = objExpenseInfo.GetProperty("U_DIM4");
                oCostAcc.DIM5 = objExpenseInfo.GetProperty("U_DIM5");
                oCostAcc.Project = objExpenseInfo.GetProperty("U_PROJECT");
                
                objExpense.expenseCostAccounting = oCostAcc;
                objExpense.expenseResponsableThirdParty = new List<ExpenseResponsableThirdParty>();

                string strExpenseType = objExpenseInfo.GetProperty("U_EXPTYPE");

                objExpense.expnseType = getExpenseType(strExpenseType);
                objExpense.legalizedValue= objExpenseInfo.GetProperty("U_REALVALUE");
                objExpense.remark = objExpenseInfo.GetProperty("Remark");
                objExpense.series = objExpenseInfo.GetProperty("Series");
                objExpense.startDate = objExpenseInfo.GetProperty("U_STARTDATE");
                objExpense.status = objExpenseInfo.GetProperty("U_STATUS");


                objThirdPartyLines = objExpenseInfo.Child("BYB_T1EXP601");
                for (int i = 0; i < objThirdPartyLines.Count; i++)
                {
                    objExpenseThirdPartyInfo = objThirdPartyLines.Item(i);
                    ExpenseResponsableThirdParty objThirdParty = new ExpenseResponsableThirdParty();
                    objThirdParty.CardCode = objExpenseThirdPartyInfo.GetProperty("U_TERRELA");
                    objExpense.expenseResponsableThirdParty.Add(objThirdParty);
                    
                }

                

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objExpense = null;
            }
            return objExpense;
        }

        private static ExpenseType getExpenseType(string strExpenseType)
        {
            ExpenseType objExpenseType = new ExpenseType();
            
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;
            SAPbobsCOM.GeneralData objExpenseTypeInfo = null;
            SAPbobsCOM.GeneralService objExpenseTypeObject = null;
            

            try
            {
                objExpenseType = new ExpenseType();

                objExpenseType.code = strExpenseType;
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objExpenseTypeObject = objCompanyService.GetGeneralService(Settings._Main.ExpenseTypeUDO);
                objFilter = objExpenseTypeObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("Code", strExpenseType);
                objExpenseTypeInfo = objExpenseTypeObject.GetByParams(objFilter);
                objExpenseType.account = objExpenseTypeInfo.GetProperty("U_MAINACCT");
                objExpenseType.expenseClass = objExpenseTypeInfo.GetProperty("U_EXPTYPE");
                objExpenseType.name = objExpenseTypeInfo.GetProperty("Name");
                objExpenseType.remark = objExpenseTypeInfo.GetProperty("U_COMMENT");
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objExpenseType = null;
            }
            return objExpenseType;
        }
        
        private static LegalizationFormCache getLegalizationInformation(int intLegalizationDocEntry, int intExpenseDocEntry, string FormUID, bool ForceLoad)
        {
            LegalizationFormCache objLegalizationFormCache = null;

            bool blLoadLegalizationInformation = false;
            
            try
            {
                objLegalizationFormCache = CacheManager.CacheManager.Instance.getFromCache(FormUID + "_" + intLegalizationDocEntry.ToString());
                if (objLegalizationFormCache != null)
                {
                    if (objLegalizationFormCache.expense.docEntry != intExpenseDocEntry)
                    {
                        blLoadLegalizationInformation = true;
                    }
                }
                else
                {
                    blLoadLegalizationInformation = true;
                }

                if(ForceLoad)
                {
                    blLoadLegalizationInformation = ForceLoad;
                }
                if(blLoadLegalizationInformation)
                {
                    objLegalizationFormCache = getLegalizationInfo(intLegalizationDocEntry);
                    CacheManager.CacheManager.Instance.addToCache(FormUID + "_" + intLegalizationDocEntry.ToString(), objLegalizationFormCache,CacheManager.CacheManager.objCachePriority.Default);
                }
                
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                objLegalizationFormCache = null;
            }
            return objLegalizationFormCache;
        }

        public static void serializeConceptInformation(string strCode)
        {
            SAPbobsCOM.CompanyService objService = null;
            SAPbobsCOM.GeneralService objConceptObject = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;
            SAPbobsCOM.GeneralData objConceptInfo = null;
            SAPbobsCOM.GeneralData objConceptThirdPartyInfo = null;
            SAPbobsCOM.GeneralData objConceptWtaxInfo = null;
            SAPbobsCOM.GeneralData objConceptVATInfo = null;
            SAPbobsCOM.GeneralDataCollection objChildCollection = null;

            string strDescription = "";
            List<string> WHTaxCodeList = null;
            List<string> VATCodeList = null;

            try
            {
                objService = MainObject.Instance.B1Company.GetCompanyService();
                objConceptObject = objService.GetGeneralService(Settings._Main.ConceptUDO);
                objFilter = objConceptObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("Code", strCode);
                objConceptInfo = objConceptObject.GetByParams(objFilter);



            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
            }
            catch (Exception ex)
            {
                _Logger.Error("", ex);
            }
        }

        public static Concept getConceptInformation(string strCode)
        {
            Concept objConcept= null;
            SAPbobsCOM.CompanyService objService = null;
            SAPbobsCOM.GeneralDataParams objFilter = null;

            SAPbobsCOM.GeneralService objConceptObject = null;
            SAPbobsCOM.GeneralData objConceptInfo = null;

            SAPbobsCOM.GeneralDataCollection objExpenseClassChild = null;
            SAPbobsCOM.GeneralDataCollection objThirdpartyChild = null;
            SAPbobsCOM.GeneralDataCollection objVATChild = null;
            SAPbobsCOM.GeneralDataCollection objWHTaxChild = null;


            string strDescription = "";
            List<string> WHTaxCodeList = null;
            List<string> VATCodeList = null;

            try
            {
                objService = MainObject.Instance.B1Company.GetCompanyService();
                objConceptObject = objService.GetGeneralService(Settings._Main.ConceptUDO);
                objFilter = objConceptObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                objFilter.SetProperty("Code", strCode);
                objConceptInfo = objConceptObject.GetByParams(objFilter);

                objConcept = new Concept();
                objConcept.Account = objConceptInfo.GetProperty("U_ACCTCODE");
                objConcept.Code = objConceptInfo.GetProperty("Code");
                objConcept.Description = objConceptInfo.GetProperty("U_DESCRIP");
                objConcept.Name = objConceptInfo.GetProperty("Name");
                objConcept.validExpenseType = new List<string>();
                objConcept.ValidThirdparty = new List<ConceptThirdParty>();
                objConcept.ValidVAT = new List<ConceptVAT>();
                objConcept.ValidWHTax = new List<ConceptWHTax>();

                objThirdpartyChild = objConceptInfo.Child("BYB_T1EXP101");
                for(int i=0; i < objThirdpartyChild.Count; i++)
                {
                    ConceptThirdParty objTP = new ConceptThirdParty();
                    objTP.Code = objThirdpartyChild.Item(i).GetProperty("U_RELPARCODE");
                    objTP.Default = Convert.ToString(objThirdpartyChild.Item(i).GetProperty("U_ISDEFAULT")) == "Y" ? true : false;
                    if (objTP.Code.Trim().Length > 0)
                    {
                        objConcept.ValidThirdparty.Add(objTP);
                    }
                }

                objWHTaxChild = objConceptInfo.Child("BYB_T1EXP102");
                for (int i = 0; i < objWHTaxChild.Count; i++)
                {
                    ConceptWHTax objWHtax = new ConceptWHTax();
                    objWHtax.Code = objWHTaxChild.Item(i).GetProperty("U_WTCODE");
                    if (objWHtax.Code.Trim().Length > 0)
                    {
                        objConcept.ValidWHTax.Add(objWHtax);
                    }
                }

                objVATChild = objConceptInfo.Child("BYB_T1EXP103");
                for (int i = 0; i < objVATChild.Count; i++)
                {
                    ConceptVAT objVAT = new ConceptVAT();
                    objVAT.Code = objVATChild.Item(i).GetProperty("U_TAXCODE");
                    if (objVAT.Code.Trim().Length > 0)
                    {
                        objConcept.ValidVAT.Add(objVAT);
                    }
                }

                objExpenseClassChild = objConceptInfo.Child("BYB_T1EXP104");
                for (int i = 0; i < objExpenseClassChild.Count; i++)
                {
                    string strExpenseClass = objExpenseClassChild.Item(i).GetProperty("U_EXPTYPE");
                    if (strExpenseClass.Trim().Length > 0)
                    {
                        objConcept.validExpenseType.Add(strExpenseClass);
                    }
                }
                
            }
            catch(COMException comEx)
            {
                _Logger.Error("", comEx);
                objConcept = null;
            }
            catch(Exception ex)
            {
                _Logger.Error("", ex);
                objConcept = null;
            }
            return objConcept;
        }

        public static void refreshLineInfo(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Conditions objConditions = null;
            SAPbouiCOM.Condition objCondition = null;
            SAPbouiCOM.ChooseFromList objCFL = null;
            SAPbouiCOM.Form objForm = null;
            

            string strType = "";
            SAPbobsCOM.Recordset objRS = null;
            LegalizationFormCache objLegalizationFormCache = null;
            string strSQL = "";
            SAPbouiCOM.DBDataSource oDBDS = null;
            int intDocEntry = -1;
            int intExpenseDocEntry = -1;
            string strConceptCode = "";
            

            SAPbouiCOM.ChooseFromListEvent oEvent = null;
            SAPbouiCOM.DBDataSource oDBLinesDS = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                oDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                intDocEntry = Convert.ToInt32(oDBDS.GetValue("DocEntry", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("DocEntry", oDBDS.Offset));
                intExpenseDocEntry = Convert.ToInt32(oDBDS.GetValue("U_EXPENSECODE", oDBDS.Offset) == "" ? "0" : oDBDS.GetValue("U_EXPENSECODE", oDBDS.Offset));
                if (intExpenseDocEntry > 0)
                {
                    objLegalizationFormCache = getLegalizationInformation(intDocEntry, intExpenseDocEntry, pVal.FormUID, false);
                }
                

                if (objLegalizationFormCache != null)
                {

                    if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {

                        for (int i = 0; i < objLegalizationFormCache.DocumentLines.Count; i++)
                        {
                            ConceptLines objConceptLine = objLegalizationFormCache.DocumentLines[i];

                            if (objConceptLine.FormLineNum == pVal.Row)
                            {
                                oDBLinesDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP401");
                                string CurrentCode = oDBLinesDS.GetValue("U_CONCEPT", oDBLinesDS.Offset);
                                if (CurrentCode.Trim() != objConceptLine.ConceptCode.Trim())
                                {

                                    Concept objConcept = getConceptInformation(CurrentCode);
                                    if (objConcept != null)
                                    {

                                        ConceptLines newConceptLine = new ConceptLines();
                                        newConceptLine.Concept = objConcept;
                                        newConceptLine.ConceptCode = objConcept.Code;
                                        newConceptLine.Description = objConcept.Description;
                                        newConceptLine.FormLineNum = pVal.Row;
                                        objLegalizationFormCache.DocumentLines[i] = newConceptLine;


                                        CacheManager.CacheManager.Instance.addToCache(pVal.FormUID + "_" + intDocEntry.ToString(), objLegalizationFormCache, CacheManager.CacheManager.objCachePriority.Default);


                                        oDBLinesDS.SetValue("U_DESCRIPTION", oDBLinesDS.Offset, newConceptLine.Description);
                                        oDBLinesDS.SetValue("U_VALUE", oDBLinesDS.Offset, "0");
                                        oDBLinesDS.SetValue("U_DATE", oDBLinesDS.Offset, "");
                                        oDBLinesDS.SetValue("U_THIRDPARTY", oDBLinesDS.Offset, "");
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

        public static void refreshFormValues(SAPbouiCOM.ItemEvent pVal, Tuple<string, string, int> tLastGotFocusColumn)
        {
            SAPbouiCOM.DBDataSource objDBDS = null;
            SAPbouiCOM.DBDataSource objDBDSLines = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;
            LegalizationFormCache objLegalizationFormCache = null;
            int intDocEntry = -1;
            int intExpenseDocEntry = -1;
            double dbRowValue = 0;
            SAPbouiCOM.CellPosition objCelPos = null;
            int intRowNumber = -1;
            
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                intDocEntry = Convert.ToInt32(objDBDS.GetValue("DocEntry", objDBDS.Offset) == "" ? "0" : objDBDS.GetValue("DocEntry", objDBDS.Offset));
                intExpenseDocEntry = Convert.ToInt32(objDBDS.GetValue("U_EXPENSECODE", objDBDS.Offset));
                if (intExpenseDocEntry > 0)
                {
                    objLegalizationFormCache = getLegalizationInformation(intDocEntry, intExpenseDocEntry, pVal.FormUID, false);
                }
                
                if (objLegalizationFormCache != null)
                {
                    
                        objMatrix = objForm.Items.Item(pVal.ItemUID).Specific;
                    objMatrix.FlushToDataSource();
                    intRowNumber = tLastGotFocusColumn.Item3-1;
                    
                    
                        objDBDSLines = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP401");
                        if (!Double.TryParse(objDBDSLines.GetValue("U_VALUE", intRowNumber), out dbRowValue))
                        {
                            dbRowValue = 0;
                        }
                        if (dbRowValue > 0)
                        {
                        

                        double dbVATTotal = 0;
                            double dbWTTAX = 0;
                            for (int i = 0; i < objLegalizationFormCache.DocumentLines.Count; i++)
                            {
                                ConceptLines objLines = objLegalizationFormCache.DocumentLines[i];
                                if (objLines.FormLineNum == pVal.Row)
                                {
                                    objLegalizationFormCache.DocumentLines[i].TotalBeforeTaxes = dbRowValue;
                                    #region calculateVAT (128)

                                    foreach (ConceptVAT objConcetpVAT in objLines.Concept.ValidVAT)
                                    {
                                        SAPbobsCOM.SalesTaxCodes objVAT = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesTaxCodes);
                                        if (objVAT.GetByKey(objConcetpVAT.Code))
                                        {
                                            double dbPercent = objVAT.Rate;
                                            dbVATTotal += dbRowValue * (dbPercent / 100);

                                        }
                                    }
                                    objLegalizationFormCache.DocumentLines[i].VAT = dbVATTotal;

                                    #endregion


                                    #region calculateWHT (178)

                                    foreach (ConceptWHTax objConcetpWHT in objLines.Concept.ValidWHTax)
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



                                    objDBDSLines.SetValue("U_VALUE", intRowNumber, Convert.ToString(objLegalizationFormCache.DocumentLines[i].TotalBeforeTaxes.ToString(CultureInfo.InvariantCulture)));

                                    objDBDSLines.SetValue("U_WHTAX", intRowNumber, Convert.ToString(objLegalizationFormCache.DocumentLines[i].WHTax.ToString(CultureInfo.InvariantCulture)));
                                    objDBDSLines.SetValue("U_VAT", intRowNumber, Convert.ToString(objLegalizationFormCache.DocumentLines[i].VAT.ToString(CultureInfo.InvariantCulture)));
                                    objDBDSLines.SetValue("U_LINETOT", intRowNumber, Convert.ToString(objLegalizationFormCache.DocumentLines[i].LineTotal.ToString(CultureInfo.InvariantCulture)));
                                recalcTotalValue(objDBDS, objLegalizationFormCache, pVal.FormUID, intDocEntry);

                                //objDBDS.SetValue("U_TOTVALUE", objDBDS.Offset,Convert.ToString(objLegalizationFormCache.TotalValue, CultureInfo.InvariantCulture));






                                    break;
                                }

                            }
                        }
                        objMatrix.LoadFromDataSource();
                    
                
                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void loadMatrixComboBoxes(string FormUID)
        {
            if (Expenses.objExpenses == null)
            {
                Expenses.objExpenses = new T1.B1.Expenses.Expenses();
            }

            bool isMultiDim = false;
            List<dimensionInfo> objDimList = null;
            List<costCenterInfo> objProfitCenterList = null;
            List<projectInfo> objProjectList = null;
            SAPbouiCOM.ComboBox objCombo = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Matrix objMatrix = null;
            SAPbouiCOM.Column objColumn = null;
            LegalizationFormCache objLegalizationFormCache = null;
            int intNumberOfRows = -1;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
                objForm.Freeze(true);


                if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    //objLegalizationFormCache = getLegalizationInformation(intDocEntry, intExpenseDocEntry, pVal.FormUID);

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
                        objMatrix = objForm.Items.Item("0_U_G").Specific;
                        intNumberOfRows = objMatrix.RowCount;
                        if (intNumberOfRows > 0)
                        {
                            objColumn = objMatrix.Columns.Item("C_0_6");
                            objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
                            for (int i = 0; i < objCombo.ValidValues.Count; i++)
                            {
                                objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }

                            foreach (costCenterInfo oPC in objProfitCenterList)
                            {
                                objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
                            }
                            updateVisualizarionString(objForm.TypeEx, "C_0_6", "0_U_G", "Centro de Costo");
                            objColumn = objMatrix.Columns.Item("C_0_7");
                            objColumn.Visible = false;
                            objColumn = objMatrix.Columns.Item("C_0_8");
                            objColumn.Visible = false;
                            objColumn = objMatrix.Columns.Item("C_0_9");
                            objColumn.Visible = false;
                            objColumn = objMatrix.Columns.Item("Col_1");
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
                                        objColumn = objMatrix.Columns.Item("C_0_6");
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
                                        updateVisualizarionString(objForm.TypeEx, "C_0_6", "0_U_G", oDIM.DimensionName);

                                    }
                                    
                                    break;
                                case 2:
                                    if (oDIM.isActive)
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_7");
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
                                        updateVisualizarionString(objForm.TypeEx, "C_0_7", "0_U_G", oDIM.DimensionName);

                                    }
                                    else
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_7");
                                        objColumn.Visible = false;
                                    }
                                    break;
                                case 3:
                                    if (oDIM.isActive)
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_8");
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
                                        updateVisualizarionString(objForm.TypeEx, "C_0_8", "0_U_G", oDIM.DimensionName);

                                    }
                                    else
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_8");
                                        objColumn.Visible = false;
                                    }
                                    break;
                                case 4:
                                    if (oDIM.isActive)
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_9");
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
                                        updateVisualizarionString(objForm.TypeEx, "C_0_9", "0_U_G", oDIM.DimensionName);

                                    }
                                    else
                                    {
                                        objColumn = objMatrix.Columns.Item("C_0_9");
                                        objColumn.Visible = false;
                                    }
                                    break;
                                case 5:
                                    if (oDIM.isActive)
                                    {
                                        objColumn = objMatrix.Columns.Item("Col_1");
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
                                        updateVisualizarionString(objForm.TypeEx, "Col_1", "0_U_G", oDIM.DimensionName);

                                    }
                                    else
                                    {
                                        objColumn = objMatrix.Columns.Item("Col_1");
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
                        objColumn = objMatrix.Columns.Item("C_0_10");
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
                if(!objDynamicStrings.GetByKey(FormUID,ColumnID,ItemID))
                {
                    objDynamicStrings.FormID = FormUID;
                    objDynamicStrings.ItemID = ItemID;
                    if (ColumnID.Trim().Length > 0)
                    {
                        objDynamicStrings.ColumnID = ColumnID;
                    }
                    objDynamicStrings.ItemString = NewValue;
                    intResult = objDynamicStrings.Add();
                    if(intResult != 0)
                    {
                        _Logger.Error(MainObject.Instance.B1Company.GetLastErrorDescription());
                    }

                }
            }
            catch(Exception er)
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

            LegalizationFormCache objFormCache = null;

            SAPbobsCOM.GeneralDataParams oFilter = null;

            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objLegalizationObject = null;
            SAPbobsCOM.GeneralData objLegalizationInfo = null;

            SAPbobsCOM.GeneralData objLegalizationJEInfo = null;
            SAPbobsCOM.GeneralDataCollection objLegalizationJELines = null;


            SAPbobsCOM.JournalEntries objAdditionalJE = null;
            SAPbobsCOM.JournalEntries objMainJE = null;

            List<SAPbobsCOM.JournalEntries> objAllJE = null;
            List<int> JETransIdListE = null;

            SAPbobsCOM.GeneralService objRequestObject = null;
            SAPbobsCOM.GeneralData objRequestInfo = null;

            //Incluir un campo para la fecha de contabilizacion general diferente a la fecha actual
            DateTime dtPostingDate = DateTime.Now;

            //Double totalCreditSCMain = 0;
            //Double totalDebitSCMain = 0;

            int intExpenseDocEntry = -1;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    objDBDS = objForm.DataSources.DBDataSources.Item("@BYB_T1EXP400");
                    intLegalizationDocEntry = Convert.ToInt32(objDBDS.GetValue("DocEntry", objDBDS.Offset));
                    intExpenseDocEntry = Convert.ToInt32(objDBDS.GetValue("U_EXPENSECODE", objDBDS.Offset));

                    objFormCache = getLegalizationInformation(intLegalizationDocEntry, intExpenseDocEntry, pVal.FormUID, true);


                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    objLegalizationObject = objCompanyService.GetGeneralService(Settings._Main.LegalizationUDO);
                    oFilter = objLegalizationObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oFilter.SetProperty("DocEntry", intLegalizationDocEntry);
                    objLegalizationInfo = objLegalizationObject.GetByParams(oFilter);

                    if (objLegalizationInfo != null && objFormCache != null)
                    {
                        
                        if (objFormCache.PostingDate == null || objFormCache.PostingDate < new DateTime(2013, 1, 1))
                        {
                            objFormCache.PostingDate = DateTime.Now;
                        }
                        else
                        {
                            dtPostingDate = objFormCache.PostingDate;
                        }

                        objMainJE = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        objMainJE.ReferenceDate = dtPostingDate;
                        objMainJE.TaxDate = dtPostingDate;
                        objMainJE.DueDate = dtPostingDate;
                        objMainJE.Reference = intLegalizationDocEntry.ToString();
                        objMainJE.Memo = "Contabilización de legalizaciones No. " + intExpenseDocEntry.ToString();
                        if (Settings._Main.LegalizationTransactionCode.Trim().Length > 0)
                        {
                            objMainJE.TransactionCode = Settings._MainPettyCash.PCLegalizationTransactionCode.Trim();
                        }
                        objMainJE.AutomaticWT = SAPbobsCOM.BoYesNoEnum.tYES;
                        objMainJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES;
                        objMainJE.ProjectCode = objFormCache.expense.expenseCostAccounting.Project;


                        bool blFirstMain = true;
                        double dbPostingTotal = 0;
                        string strOutMessage = "";

                        //Dictionary<string, double> objWT = new Dictionary<string, double>();
                        foreach (ConceptLines objLines in objFormCache.DocumentLines)
                        {
                            if (objLines.ConceptCode.Trim().Length > 0)
                            {
                                if (objLines.Concept.ValidWHTax != null && objLines.Concept.ValidWHTax.Count > 0)
                                {
                                    Double totalCreditSCAdditional = 0;
                                    Double totalDebitSCAdditional = 0;

                                    //double dbTotal = objLines.TotalBeforeTaxes + objLines.VAT;

                                    objAdditionalJE = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                                    objAdditionalJE.ReferenceDate = dtPostingDate;
                                    objAdditionalJE.TaxDate = dtPostingDate;
                                    objAdditionalJE.DueDate = dtPostingDate;
                                    objAdditionalJE.Reference = intLegalizationDocEntry.ToString();
                                    objAdditionalJE.Memo = "Contabilización de legalizaciones No. " + intExpenseDocEntry.ToString();
                                    if (Settings._Main.LegalizationTransactionCode.Trim().Length > 0)
                                    {
                                        objAdditionalJE.TransactionCode = Settings._MainPettyCash.PCLegalizationTransactionCode.Trim();
                                    }
                                    objAdditionalJE.AutomaticWT = SAPbobsCOM.BoYesNoEnum.tYES;
                                    objAdditionalJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES;
                                    objAdditionalJE.ProjectCode = objFormCache.expense.expenseCostAccounting.Project;

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
                                    if (objLines.Concept.ValidVAT != null && objLines.Concept.ValidVAT.Count > 0)
                                    {
                                        foreach (ConceptVAT objVat in objLines.Concept.ValidVAT)
                                        {
                                            if (objVat.Code.Trim().Length > 0)
                                            {
                                                objAdditionalJE.Lines.TaxPostAccount = SAPbobsCOM.BoTaxPostAccEnum.tpa_PurchaseTaxAccount;
                                                objAdditionalJE.Lines.TaxCode = objVat.Code.Trim();
                                                totalDebitSCAdditional += T1.B1.Base.DIOperations.Operations.getSCValue(T1.B1.Base.DIOperations.Operations.getTaxAmountLC(objVat.Code.Trim(), objLines.TotalBeforeTaxes), dtPostingDate, out strOutMessage, SAPbobsCOM.RoundingContextEnum.rcTax);
                                                break;
                                            }
                                        }
                                    }
                                    objAdditionalJE.Lines.WTLiable = SAPbobsCOM.BoYesNoEnum.tYES;
                                    objAdditionalJE.Lines.ProjectCode = objLines.Project;
                                    objAdditionalJE.Lines.CostingCode = objLines.DIM1;
                                    objAdditionalJE.Lines.CostingCode2 = objLines.DIM2;
                                    objAdditionalJE.Lines.CostingCode3 = objLines.DIM3;
                                    objAdditionalJE.Lines.CostingCode4 = objLines.DIM4;
                                    objAdditionalJE.Lines.CostingCode5 = objLines.DIM5;

                                    bool blFirstWT = true;
                                    foreach (ConceptWHTax oWTCode in objLines.Concept.ValidWHTax)
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
                                            objAdditionalJE.WithholdingTaxData.WTCode = oWTCode.Code;
                                            blFirstWT = false;
                                        }
                                    }
                                    double dbDifference = totalCreditSCAdditional - totalDebitSCAdditional;

                                    if (dbDifference > 0)
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

                                    if (objAllJE == null)
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
                                        if (objFormCache.expense.expenseResponsableThirdParty != null && objFormCache.expense.expenseResponsableThirdParty.Count > 0)
                                        {
                                            foreach (ExpenseResponsableThirdParty oTP in objFormCache.expense.expenseResponsableThirdParty)
                                            {
                                                string strTP = oTP.CardCode.Trim();
                                                if (strTP.Length > 0)
                                                {
                                                    objMainJE.Lines.ShortName = strTP;
                                                    objMainJE.Lines.ControlAccount = objFormCache.expense.expnseType.account;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            objMainJE.Lines.AccountCode = objFormCache.expense.expnseType.account;
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

                                    if (objLines.Concept.ValidVAT != null && objLines.Concept.ValidVAT.Count > 0)
                                    {
                                        foreach (ConceptVAT objVat in objLines.Concept.ValidVAT)
                                        {
                                            if (objVat.Code.Trim().Length > 0)
                                            {
                                                objMainJE.Lines.TaxPostAccount = SAPbobsCOM.BoTaxPostAccEnum.tpa_PurchaseTaxAccount;
                                                objMainJE.Lines.TaxCode = objVat.Code.Trim();
                                                dbPostingTotal += objLines.VAT;
                                                //dbPostingTotal += T1.B1.Base.DIOperations.Operations.getSCValue(T1.B1.Base.DIOperations.Operations.getTaxAmountLC(objVat.Code.Trim(), objLines.TotalBeforeTaxes), dtPostingDate, out strOutMessage, SAPbobsCOM.RoundingContextEnum.rcTax);
                                                break;
                                            }
                                        }
                                    }

                                }
                            }

                        }

                        objMainJE.Lines.SetCurrentLine(0);
                        objMainJE.Lines.Credit = dbPostingTotal;
                        if (dbPostingTotal > 0)
                        {

                            if (objAllJE == null)
                            {
                                objAllJE = new List<SAPbobsCOM.JournalEntries>();
                            }
                            objAllJE.Add(objMainJE);
                        }

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

                                objLegalizationInfo.SetProperty("U_ISPOSTED", "Y");
                                objLegalizationJELines = objLegalizationInfo.Child("BYB_T1EXP403");
                                
                                foreach(int intJE in JETransIdListE)
                                {
                                    SAPbobsCOM.JournalEntries objJEItem = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                                    if(objJEItem.GetByKey(intJE))
                                    {
                                        objLegalizationJEInfo = objLegalizationJELines.Add();
                                        double dbValue = 0;
                                        objJEItem.Lines.SetCurrentLine(0);
                                        dbValue += Math.Abs(objJEItem.Lines.Debit - objJEItem.Lines.Credit);
                                        objLegalizationJEInfo.SetProperty("U_VALUE", dbValue);
                                        objLegalizationJEInfo.SetProperty("U_JEENTRY", intJE);
                                        objLegalizationJEInfo.SetProperty("U_POSTDATE", objJEItem.ReferenceDate);
                                    }
                                }
                                objLegalizationObject.Update(objLegalizationInfo);

                                objRequestObject = objCompanyService.GetGeneralService("BYB_T1EXP600");
                                oFilter = objRequestObject.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                oFilter.SetProperty("DocEntry", intExpenseDocEntry);
                                objRequestInfo = objRequestObject.GetByParams(oFilter);
                                objRequestInfo.SetProperty("U_STATUS", "CERRADA");
                                objRequestInfo.SetProperty("U_REALVALUE", objFormCache.TotalValue);
                                objRequestObject.Update(objRequestInfo);

                                MainObject.Instance.B1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            }
                            objForm.Items.Item("btnContab").Enabled = false;
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
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);

            }
        }

        private static void recalcTotalValue(SAPbouiCOM.DBDataSource objDBDS, LegalizationFormCache objLegalizationFormCache, string FormUID, int intDocEntry)
        {
            double dbTotal = 0;
            try
            {
                
                foreach (ConceptLines oLines in objLegalizationFormCache.DocumentLines)
                {
                    dbTotal += oLines.LineTotal;
                }
                objLegalizationFormCache.TotalValue = dbTotal;
                objDBDS.SetValue("U_TOTVALUE", objDBDS.Offset, Convert.ToString(objLegalizationFormCache.TotalValue, CultureInfo.InvariantCulture));
                CacheManager.CacheManager.Instance.addToCache(FormUID + "_" + intDocEntry.ToString(), objLegalizationFormCache, CacheManager.CacheManager.objCachePriority.Default);
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }

        }

        #endregion

    }
}






