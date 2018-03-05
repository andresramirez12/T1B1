using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;
using System.Xml;
using SAPbobsCOM;
using System.Globalization;
using System.ComponentModel;
using Newtonsoft.Json;
using System.Drawing;

namespace T1.B1.WithholdingTax
{
    public class WithholdingTax
    {
        private static WithholdingTax objWithHoldingTax;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private WithholdingTax()
        {
            if (objWithHoldingTax == null)
            {
                objWithHoldingTax = new WithholdingTax();
            }
        }


        static public void getSelectedBPInformation(SAPbouiCOM.ItemEvent pVal, bool useCFL)
        {
            SAPbouiCOM.ChooseFromListEvent oCFLEvent = null;
            SAPbouiCOM.DBDataSource oDB = null;
            SAPbouiCOM.Form objForm = null;
            bool blReadWTConfig = false;
            try
            {
                if (useCFL)
                {
                    oCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    if (oCFLEvent != null && oCFLEvent.SelectedObjects != null)
                    {
                        bool isDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + pVal.FormUID) == null ? false : true;
                        if (!isDisabled)
                        {
                            string strLastCardCode = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID) == null ? "" : CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID);

                            bool blAddOnCalc = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + pVal.FormUID) == null ? false : true;
                            if (strLastCardCode.Trim().Length == 0)
                            {
                                blAddOnCalc = true;
                            }
                            if (blAddOnCalc)
                            {

                                string strPickedCardCode = oCFLEvent.SelectedObjects.GetValue("CardCode", 0);
                                if (strLastCardCode.Trim().Length == 0)
                                {
                                    blReadWTConfig = true;

                                }
                                else
                                {
                                    if (strPickedCardCode.Trim() != strLastCardCode)
                                    {
                                        blReadWTConfig = true;
                                    }
                                }
                                if (blReadWTConfig)
                                {
                                    CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID, strPickedCardCode, CacheManager.CacheManager.objCachePriority.Default);
                                    WithholdingTax.getWTforBP(pVal,true);
                                }
                            }
                        }
                    }
                }
                else
                {
                    

                    bool isDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + pVal.FormUID) == null ? false : true;
                    if (!isDisabled)
                    {
                        string strLastCardCode = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID) == null ? "" : CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID);

                        bool blAddOnCalc = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + pVal.FormUID) == null ? false : true;
                        if (strLastCardCode.Trim().Length == 0)
                        {
                            blAddOnCalc = true;
                        }
                        if (blAddOnCalc)
                        {
                            objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                            string strPickedCardCode = (objForm.TypeEx != "133" ? objForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim() : objForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim());



                            if (strLastCardCode.Trim().Length == 0)
                            {
                                blReadWTConfig = true;

                            }
                            else
                            {
                                if (strPickedCardCode.Trim() != strLastCardCode)
                                {
                                    blReadWTConfig = true;
                                }
                            }
                            if (blReadWTConfig)
                            {
                                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID, strPickedCardCode, CacheManager.CacheManager.objCachePriority.Default);
                                WithholdingTax.getWTforBP(pVal, useCFL);
                            }
                        }
                    }
                        
                    
                }
            }
            catch(COMException comEx)
            {
                _Logger.Error("", comEx);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + comEx.InnerException.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        

        public static void getWTforBP(SAPbouiCOM.ItemEvent pVal, bool useCFL)
        {
            SAPbouiCOM.Form objForm = null;
            string strCardCode = "";
            SAPbobsCOM.BusinessPartners objBP = null;
            SAPbobsCOM.WithholdingTaxCodes objWTINfo = null;
            XmlDocument xmlDocument = null;
            XmlNodeList xmlNodes = null;
            List<WithholdingTaxConfigDetail> objWithHoldingTaxInfo = null;
            SAPbouiCOM.ChooseFromListEvent oCFLEvent = null;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if (useCFL)
                {


                    oCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    if (oCFLEvent != null)
                    {
                        if (oCFLEvent.SelectedObjects != null)
                        {
                            strCardCode = oCFLEvent.SelectedObjects.GetValue("CardCode", 0);
                            //strCardCode = (objForm.TypeEx != "133" ? objForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim() : objForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim());

                            objBP = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                            if (objBP.GetByKey(strCardCode))
                            {
                                xmlDocument = new XmlDocument();
                                xmlDocument.LoadXml(objBP.GetAsXML());
                                xmlNodes = xmlDocument.SelectNodes("/BOM/BO/BPWithholdingTax/row/WTCode");
                                if (xmlNodes != null)
                                {
                                    objWithHoldingTaxInfo = new List<WithholdingTaxConfigDetail>();
                                    T1.B1.Base.UIOperations.Operations.startProgressBar("Cargando retenciones asignadas al Socio de negocio", 2);
                                    foreach (XmlNode xn in xmlNodes)
                                    {
                                        string WTCode = xn.InnerText;
                                        objWTINfo = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                                        if (objWTINfo.GetByKey(WTCode))
                                        {
                                            WithholdingTaxConfigDetail objDet = new WithholdingTaxConfigDetail();
                                            objDet.WTCode = WTCode;
                                            objDet.U_BYBCOMM = objWTINfo.UserFields.Fields.Item("U_BYB_COMM").Value;
                                            objDet.U_BYB_AFEC = objWTINfo.UserFields.Fields.Item("U_BYB_AFEC").Value;
                                            objDet.U_BYB_MIN = objWTINfo.UserFields.Fields.Item("U_BYB_MIN").Value;
                                            objDet.U_BYB_MUNI = objWTINfo.UserFields.Fields.Item("U_BYB_MUNI").Value;
                                            objDet.U_BYB_TIPO = objWTINfo.UserFields.Fields.Item("U_BYB_TIPO").Value;
                                            objDet.MUNI = getWTMuniInfo(objDet.U_BYB_MUNI);
                                            objWithHoldingTaxInfo.Add(objDet);
                                        }
                                    }
                                    T1.B1.Base.UIOperations.Operations.stopProgressBar();
                                }
                                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + objForm.UniqueID, objWithHoldingTaxInfo, CacheManager.CacheManager.objCachePriority.Default);
                                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + objForm.UniqueID, true, CacheManager.CacheManager.objCachePriority.Default);
                                //activateWTMenu(objForm.UniqueID);

                            }
                        }
                    }
                }
                else
                {
                    
                           // strCardCode = oCFLEvent.SelectedObjects.GetValue("CardCode", 0);
                            strCardCode = (objForm.TypeEx != "133" ? objForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim() : objForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim());

                            objBP = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                            if (objBP.GetByKey(strCardCode))
                            {
                                xmlDocument = new XmlDocument();
                                xmlDocument.LoadXml(objBP.GetAsXML());
                                xmlNodes = xmlDocument.SelectNodes("/BOM/BO/BPWithholdingTax/row/WTCode");
                                if (xmlNodes != null)
                                {
                                    objWithHoldingTaxInfo = new List<WithholdingTaxConfigDetail>();
                                    T1.B1.Base.UIOperations.Operations.startProgressBar("Cargando retenciones asignadas al Socio de negocio", 2);
                                    foreach (XmlNode xn in xmlNodes)
                                    {
                                        string WTCode = xn.InnerText;
                                        objWTINfo = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                                        if (objWTINfo.GetByKey(WTCode))
                                        {
                                            WithholdingTaxConfigDetail objDet = new WithholdingTaxConfigDetail();
                                            objDet.WTCode = WTCode;
                                            objDet.U_BYBCOMM = objWTINfo.UserFields.Fields.Item("U_BYB_COMM").Value;
                                            objDet.U_BYB_AFEC = objWTINfo.UserFields.Fields.Item("U_BYB_AFEC").Value;
                                            objDet.U_BYB_MIN = objWTINfo.UserFields.Fields.Item("U_BYB_MIN").Value;
                                            objDet.U_BYB_MUNI = objWTINfo.UserFields.Fields.Item("U_BYB_MUNI").Value;
                                            objDet.U_BYB_TIPO = objWTINfo.UserFields.Fields.Item("U_BYB_TIPO").Value;
                                            objDet.MUNI = getWTMuniInfo(objDet.U_BYB_MUNI);
                                            objWithHoldingTaxInfo.Add(objDet);
                                        }
                                    }
                                    T1.B1.Base.UIOperations.Operations.stopProgressBar();
                                }
                                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + objForm.UniqueID, objWithHoldingTaxInfo, CacheManager.CacheManager.objCachePriority.Default);
                                CacheManager.CacheManager.Instance.addToCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + objForm.UniqueID, true, CacheManager.CacheManager.objCachePriority.Default);
                                //activateWTMenu(objForm.UniqueID);

                            }
                        
                }
            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
                T1.B1.Base.UIOperations.Operations.stopProgressBar();

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
            }
        }

        public static void activateWTMenu(string FormUID)
        {
            SAPbouiCOM.MenuItem objMenuItem = null;

            try
            {
                objMenuItem = MainObject.Instance.B1Application.Menus.Item("5897");
                if (objMenuItem.Enabled)
                {
                    CacheManager.CacheManager.Instance.addToCache("WTAutoActivate", FormUID, CacheManager.CacheManager.objCachePriority.Default);
                    MainObject.Instance.B1Application.ActivateMenuItem("5897");
                    CacheManager.CacheManager.Instance.removeFromCache("WTAutoActivate");
                }

            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);


            }
            catch (Exception er)
            {
                _Logger.Error("", er);

            }
        }

        public static void setBPWT(string FormUID, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Form objDocument = null;
            List<WithholdingTaxConfigDetail> objWithHoldingTaxInfo = null;
            bool blAutoGenerated = false;
            SAPbouiCOM.Matrix objMatrix = null;
            SAPbouiCOM.EditText objEdit = null;
            string strWTCodeValue = "";
            int intNum = -1;
            string strMuniCOdeInAddress = "";
            bool blFirstTime = false;

            try
            {
                blAutoGenerated = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + FormUID) != null ? true : false;
                if (blAutoGenerated)
                {
                    objWithHoldingTaxInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + FormUID) != null ? CacheManager.CacheManager.Instance.getFromCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + FormUID) : new List<WithholdingTaxConfigDetail>();

                }

                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objMatrix = objForm.Items.Item("6").Specific;
                
                    

                if (objWithHoldingTaxInfo != null && objWithHoldingTaxInfo.Count > 0)
                {

                    objDocument = MainObject.Instance.B1Application.Forms.Item(FormUID);
                    strMuniCOdeInAddress = getMuniFromDocument(objDocument);

                    objMatrix.Clear();
                    objMatrix.AddRow(1, -1);

                    #region Fill Matrix
                    intNum = 1;
                    bool blFirst = true;
                    bool blAddressMuni = false;//There is an Muni in the Bp Address
                    if (strMuniCOdeInAddress.Trim().Length > 0)
                    {
                        blAddressMuni = true;
                    }
                    foreach (WithholdingTaxConfigDetail oDetail in objWithHoldingTaxInfo)
                    {
                        
                        bool blCheckMuni = false;
                        //The WT has Muni specific
                        if (blAddressMuni && oDetail.MUNI.Count > 0)
                        {
                            blCheckMuni = true;
                        }
                        //Get the first column in the matrix first row
                        objEdit = objMatrix.GetCellSpecific("1", intNum);
                        string strWTCode = oDetail.WTCode;
                        bool blFound = false;
                        #region Muni Check
                        if (blCheckMuni)
                        {

                            foreach (WithholdingTaxConfigMun oMun in oDetail.MUNI)
                            {
                                if (oMun.U_MUNCODE == strMuniCOdeInAddress)
                                {
                                    blFound = true;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            blFound = true;
                        }
                        #endregion
                        if (!blFound)
                        {
                            strWTCode = "";
                        }
                        if (strWTCode.Trim().Length > 0)
                        {
                            if (!blFirst)
                            {
                                objMatrix.AddRow(1, -1);
                            }
                            objEdit.Value = strWTCode;
                            blFirst = false;
                            intNum++;
                        }
                    }
                    #endregion

                    if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    objForm.Close();
                }
                else
                {
                    objMatrix.Clear();
                    objMatrix.AddRow(1, -1);

                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    if (MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID == pVal.FormUID)
                    {
                        objForm.Close();
                    }
                }
                CacheManager.CacheManager.Instance.addToCache("WTLogicDone_" + FormUID, true, CacheManager.CacheManager.objCachePriority.NotRemovable);
            }
            catch (COMException cOMException1)
            {

                _Logger.Error("", cOMException1);
            }
            catch (Exception exception2)
            {
                _Logger.Error("", exception2);
            }
        }

        private static List<WithholdingTaxConfigMun> getWTMuniInfo(string strCode)
        {
            SAPbobsCOM.CompanyService objCompany = null;
            SAPbobsCOM.GeneralService objWTInfoService = null;
            SAPbobsCOM.GeneralDataParams objGetParams = null;
            SAPbobsCOM.GeneralData objGeneralData = null;
            SAPbobsCOM.GeneralDataCollection objMuniINfo = null;
            List < WithholdingTaxConfigMun >objList = null;

            try
            {
                if (strCode.Trim().Length > 0)
                {
                    objCompany = MainObject.Instance.B1Company.GetCompanyService();
                    objWTInfoService = objCompany.GetGeneralService(Settings._WithHoldingTax.WTMuniInfoUDO);
                    objGetParams = objWTInfoService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    objGetParams.SetProperty("Code", strCode);
                    objGeneralData = objWTInfoService.GetByParams(objGetParams);
                    objMuniINfo = objGeneralData.Child(Settings._WithHoldingTax.WTMuniInfoChildUDO);
                    if (objMuniINfo.Count > 0)
                    {
                        objList = new List<WithholdingTaxConfigMun>();
                        foreach (GeneralData oDef in objMuniINfo)
                        {
                            try
                            {
                                string strMunCode = oDef.GetProperty("U_MUNCODE");
                                if (strMunCode != null && strMunCode.Trim().Length > 0)
                                {
                                    WithholdingTaxConfigMun oMun = new WithholdingTaxConfigMun();
                                    oMun.U_MUNCODE = strMunCode;
                                    objList.Add(oMun);


                                }
                            }
                            catch (Exception er)
                            {
                                _Logger.Error("", er);
                            }
                        }
                    }
                }
                else
                {
                    objList = new List<WithholdingTaxConfigMun>();
                }

            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
                objList = new List<WithholdingTaxConfigMun>();

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objList = new List<WithholdingTaxConfigMun>();
            }
            return objList;
        }
        
        
        private static string getMuniFromDocument(SAPbouiCOM.Form objForm)
        {
            string strCode = "";
            string strCardCode = "";
            string strAddressCode = "";
            SAPbobsCOM.BusinessPartners objBP = null;
            SAPbobsCOM.BPAddresses objBpAddress = null;

            try
            {
                strCardCode = (objForm.TypeEx != "133" ? objForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim() : objForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim());
                strAddressCode = (objForm.TypeEx != "133" ? objForm.DataSources.DBDataSources.Item("OPCH").GetValue("PayToCode", 0).Trim() : objForm.DataSources.DBDataSources.Item("OINV").GetValue("PayToCode", 0).Trim());
                objBP = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                if(objBP.GetByKey(strCardCode))
                {
                    objBpAddress = objBP.Addresses;
                    for(int i=0; i < objBpAddress.Count; i++)
                    {
                        objBpAddress.SetCurrentLine(i);
                        if(objBpAddress.AddressType == BoAddressType.bo_BillTo && objBpAddress.AddressName == strAddressCode)
                        {
                            strCode = objBpAddress.UserFields.Fields.Item("U_BYB_MUNI").Value;
                            break;
                        }
                    }
                    
                }
            }
            catch (COMException comEx)
            {
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", comEx.ErrorCode, "::", comEx.Message, "::", comEx.StackTrace })));
                _Logger.Error("", exception);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: No se pudo recuperar el municipio de la direccion de pago. Las retenciones no se filtraran por municipio", SAPbouiCOM.BoMessageTime.bmt_Short, true);

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                MainObject.Instance.B1Application.SetStatusBarMessage("BYB: No se pudo recuperar el municipio de la direccion de pago. Las retenciones no se filtraran por municipio", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            return strCode;
        }

        static public bool formModeAdd(SAPbouiCOM.ItemEvent pVal)
        {
            bool blResult = false;
            try
            {
                SAPbouiCOM.Form objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                if(objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    blResult = true;
                }
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objForm);
                objForm = null;
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                blResult = false;
            }
            return blResult;
        }



        #region Internal Entries

        private static double getBaseAmount(SAPbobsCOM.Documents objDoc)
        {
            double dbBase = 0;
            try
            {
                //dbBase = objDoc.BaseAmount;
                if (dbBase == 0)
                {
                    //double dbExpenses = 0;
                    //SAPbobsCOM.DocumentsAdditionalExpenses objExp = objDoc.Expenses;
                    //for(int i=0; i < objExp.Count; i++)
                    //{
                    //  objExp.SetCurrentLine(i);
                    //dbExpenses += objExp.LineTotal;
                    //}
                    double dbTest = objDoc.RoundingDiffAmount > 0.5 ? objDoc.RoundingDiffAmount : objDoc.RoundingDiffAmount * -1;
                    dbBase = objDoc.DocTotal - objDoc.VatSum + objDoc.WTAmount + dbTest;
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                dbBase = -1;
            }
            return dbBase;
        }

        static public void addDocumentInfo(AddDocumentInfoArgs BusinessObjectInfo)
        {
            BackgroundWorker addDocumentInfoWorker = new BackgroundWorker();
            addDocumentInfoWorker.WorkerSupportsCancellation = false;
            addDocumentInfoWorker.WorkerReportsProgress = false;
            addDocumentInfoWorker.DoWork += AddDocumentInfoWorker_DoWork;
            addDocumentInfoWorker.RunWorkerCompleted += AddDocumentInfoWorker_RunWorkerCompleted;
            addDocumentInfoWorker.RunWorkerAsync(BusinessObjectInfo);
            
        }

        static private void AddDocumentInfoWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!(e.Error == null))
            {
                _Logger.Error(e.Error.Message);

            }

            else
            {
                //this.tbProgress.Text = "Done!";
            }
        }

        static private void AddDocumentInfoWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            AddDocumentInfoArgs oInfo = null;
            SAPbobsCOM.Documents objDoc = null;
           List<string> WHPurchaseDocuments = new List<string>();
        List<string> WHSalesDocuments = new List<string>();
            SAPbobsCOM.CompanyService objCompanyService = null;

            SAPbobsCOM.GeneralDataParams oFilter = null;

            SAPbobsCOM.GeneralService objEntryObject = null;
            SAPbobsCOM.GeneralData objEntryInfo = null;
            SAPbobsCOM.GeneralData objEntryLinesInfo = null;
            SAPbobsCOM.GeneralDataCollection objEntryLinesObject = null;

            SAPbobsCOM.GeneralService objRelatedPartyObject = null;
            SAPbobsCOM.GeneralData objRelatedPartyInfo = null;

            int intDocEntry = -1;
            int intDocNum = -1;
            string strDocType = "";
            string strCardCode = "";
            string strRelatedParty = "";
            double dbDocTotal = 0;
            double dbBaseAmnt = 0;

            int intJE = -1;
            string strOerationType = "";
            string strOpeation = "";
            string strCode = "";
            double dbPercent = 0;
            double dbBase = 0;
            double dbValue = 0;
            int intDocLine = 0;


            try
            {
                oInfo = (AddDocumentInfoArgs)e.Argument;
                WHPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTPurchaseObjects);
                WHSalesDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTSalesObjects);

                var B1Object = MainObject.Instance.B1Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), oInfo.ObjectType));

                if (B1Object.Browser.GetByKeys(oInfo.ObjectKey))
                {
                    if (WHPurchaseDocuments.Contains(oInfo.FormtTypeEx) || WHSalesDocuments.Contains(oInfo.FormtTypeEx))
                    {
                        objDoc = (SAPbobsCOM.Documents)B1Object;
                        objCompanyService = MainObject.Instance.B1Company.GetCompanyService();

                        strCardCode = objDoc.CardCode;
                        intDocEntry = objDoc.DocEntry;
                        intDocNum = objDoc.DocNum;
                        strDocType = oInfo.ObjectType.Trim();
                        dbDocTotal = objDoc.DocTotal;
                        dbBaseAmnt = getBaseAmount(objDoc);
                        #region Get ThirdParty Info
                        objRelatedPartyObject = objCompanyService.GetGeneralService("BYB_T1RPA100");
                        oFilter = objRelatedPartyObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oFilter.SetProperty("Code", objDoc.CardCode);
                        try
                        {
                            objRelatedPartyInfo = objRelatedPartyObject.GetByParams(oFilter);
                            strRelatedParty = objRelatedPartyInfo.GetProperty("Code");
                        }
                        catch(Exception er)
                        {
                            strRelatedParty = "";
                        }
                        #endregion

                        objEntryObject = objCompanyService.GetGeneralService("BYB_T1WHT400");
                        objEntryInfo = objEntryObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                        objEntryInfo.SetProperty("U_DOCNUM", intDocNum);
                        objEntryInfo.SetProperty("U_DOCENTRY", intDocEntry);
                        objEntryInfo.SetProperty("U_CARDCODE", strCardCode);
                        objEntryInfo.SetProperty("U_RELPARTY", strRelatedParty);
                        objEntryInfo.SetProperty("U_DOCTYPE", strDocType);
                        objEntryInfo.SetProperty("U_DOCTOTAL", dbDocTotal);
                        objEntryInfo.SetProperty("U_BASEAMNT", dbBaseAmnt);

                        intJE = objDoc.TransNum;

                        objEntryLinesObject = objEntryInfo.Child("BYB_T1WHT401");

                        SAPbobsCOM.WithholdingTaxData oWHTData = objDoc.WithholdingTaxData;
                        for(int i =0; i < oWHTData.Count; i++)
                        {
                            oWHTData.SetCurrentLine(i);
                            SAPbobsCOM.WithholdingTaxCodes oWT = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                            if (oWT.GetByKey(oWHTData.WTCode))
                            {

                                objEntryLinesInfo = objEntryLinesObject.Add();
                                objEntryLinesInfo.SetProperty("U_JE", intJE);
                                objEntryLinesInfo.SetProperty("U_OPERTYPE", "WHT");
                                objEntryLinesInfo.SetProperty("U_OPER", "INI");
                                objEntryLinesInfo.SetProperty("U_CODE", oWHTData.WTCode);
                                objEntryLinesInfo.SetProperty("U_PERCENT", oWT.BaseAmount);

                                if(oWT.BaseType == WithholdingTaxCodeBaseTypeEnum.wtcbt_VAT)
                                {
                                    objEntryLinesInfo.SetProperty("U_BASEAMNT", objDoc.VatSum);
                                }
                                else
                                {
                                    objEntryLinesInfo.SetProperty("U_BASEAMNT", dbBaseAmnt);
                                }
                                objEntryLinesInfo.SetProperty("U_AMNT", oWHTData.TaxableAmount);
                                objEntryLinesInfo.SetProperty("U_DOCLINE", -1);
                            }


                        }

                        for (int i = 0; i < objDoc.Lines.Count; i++)
                        {
                            objDoc.Lines.SetCurrentLine(i);
                            SAPbobsCOM.SalesTaxCodes oSTC = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oSalesTaxCodes);
                            if (oSTC.GetByKey(objDoc.Lines.TaxCode))
                            {

                                objEntryLinesInfo = objEntryLinesObject.Add();
                                objEntryLinesInfo.SetProperty("U_JE", intJE);
                                objEntryLinesInfo.SetProperty("U_OPERTYPE", "TAX");
                                objEntryLinesInfo.SetProperty("U_OPER", "INI");
                                objEntryLinesInfo.SetProperty("U_CODE", objDoc.Lines.TaxCode);
                                objEntryLinesInfo.SetProperty("U_PERCENT", oSTC.Rate);
                                objEntryLinesInfo.SetProperty("U_BASEAMNT", objDoc.Lines.LineTotal);
                                
                                objEntryLinesInfo.SetProperty("U_AMNT", objDoc.Lines.NetTaxAmount);
                                objEntryLinesInfo.SetProperty("U_DOCLINE", objDoc.Lines.LineNum);
                            }


                        }

                        objEntryObject.Add(objEntryInfo);

                    }
                }

                        



            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
               
            }
        }




        #endregion

        #region Operaciones Faltantes

        static public void loadMissingOperationsForm()
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
                objParams.XmlData = SWTaxResources.OperacionesFaltantesRetencion;
                objParams.FormType = "BYB_FMWT01";
                objForm = MainObject.Instance.B1Application.Forms.AddEx(objParams);
                objDT = objForm.DataSources.DataTables.Item("DT_TRA");


                objRS = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName))
                {
                    strSQL = Settings._HANA.getMissingOperationsQuery;
                }
                else
                {
                    strSQL = Settings._SQL.getMissingOperationsQuery;
                }

                objDT.ExecuteQuery(strSQL);

                #region Format Grid
                objGrid = objForm.Items.Item("grTRA").Specific;

                //objGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;


                objGridColumn = objGrid.Columns.Item(0);
                objGridColumn.Editable = false;

                objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                oEditTExt.LinkedObjectType = "18";


                objGridColumn = objGrid.Columns.Item(1);
                objGridColumn.Editable = false;


                objGridColumn = objGrid.Columns.Item(2);
                objGridColumn.Editable = false;

                objGridColumn = objGrid.Columns.Item(3);
                objGridColumn.Editable = false;

                objGridColumn = objGrid.Columns.Item(4);
                objGridColumn.RightJustified = true;
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

        static public void createMissingOperations(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.DataTable objDT = null;
            SAPbouiCOM.DataTable objDtRes = null;
            SAPbouiCOM.Form objForm = null;
            SAPbouiCOM.Grid objGrid = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.GeneralService objGeneralService = null;
            SAPbobsCOM.GeneralData objGeneralData = null;
            SAPbobsCOM.GeneralDataParams objResult = null;
            SAPbouiCOM.Item objItem = null;
            string strInternalCode = "";
            string strLegalName = "";
            string strID = "";
            SAPbouiCOM.GridColumn objGridColumn;
            SAPbouiCOM.EditTextColumn oEditTExt;



            SAPbobsCOM.BusinessPartners objBP = null;

            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                objDT = objForm.DataSources.DataTables.Item("DT_TRA");
                objDtRes = objForm.DataSources.DataTables.Item("DT_RES");
                if (objDT.Rows.Count > 0)
                {
                    T1.B1.Base.UIOperations.Operations.startProgressBar("Procesando...", objDT.Rows.Count);
                    objForm.Freeze(true);
                    for (int i = 0; i < objDT.Rows.Count; i++)
                    {
                        string strDocumentType = objDT.GetValue(3, i);
                        int intDocEntry = objDT.GetValue(0, i);
                        int intDocNum = objDT.GetValue(1, i);
                        string strResult = "";
                        switch (strDocumentType)
                        {
                            case "Factura Proveedor":
                                strResult = addPurchaseInvoiceWTInternalRegistry(intDocEntry);
                                break;
                            case "NC Factura Proveedor":
                                strResult = addCNPurchaseInvoiceWTInternalRegistry(intDocEntry);
                                break;
                        }

                        objDtRes.Rows.Add(1);
                        objDtRes.SetValue(0, objDtRes.Rows.Count - 1, objDT.GetValue(0, i));
                        objDtRes.SetValue(1, objDtRes.Rows.Count - 1, objDT.GetValue(1, i));
                        objDtRes.SetValue(2, objDtRes.Rows.Count - 1, objDT.GetValue(2, i));
                        objDtRes.SetValue(3, objDtRes.Rows.Count - 1, objDT.GetValue(3, i));
                        objDtRes.SetValue(4, objDtRes.Rows.Count - 1, objDT.GetValue(4, i));
                        objDtRes.SetValue(5, objDtRes.Rows.Count - 1, strResult);

                        T1.B1.Base.UIOperations.Operations.setProgressBarMessage(strDocumentType + " " + intDocNum + " procesada.", i + 1);

                    }

                    objGrid = objForm.Items.Item("grTRA").Specific;

                    objGrid.DataTable = objDtRes;
                 

                    objGridColumn = objGrid.Columns.Item(0);
                    objGridColumn.Editable = false;

                    objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                    oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                    oEditTExt.LinkedObjectType = "18";


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
                    objGridColumn.Editable = false;


                    for (int i = 0; i < objDtRes.Rows.Count; i++)
                    {
                        string strResult = objDtRes.GetValue(5, i);
                        strResult = strResult.Substring(0, 2);
                        if (strResult == "OK")
                        {
                            objGrid.CommonSetting.SetCellBackColor(i + 1, 4, Color.Green.R | (Color.Green.G << 8) | (Color.Green.B << 16));

                        }
                        else
                        {
                            objGrid.CommonSetting.SetCellBackColor(i + 1, 4, Color.Red.R | (Color.Red.G << 8) | (Color.Red.B << 16));

                        }
                    }


                    //objGrid.DataTable = objDtRes;

                    //objGrid = objForm.Items.Item("grTRA").Specific;

                    ////objGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;


                    //objGridColumn = objGrid.Columns.Item(0);
                    //objGridColumn.Editable = false;

                    //objGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                    //oEditTExt = (SAPbouiCOM.EditTextColumn)objGridColumn;
                    //oEditTExt.LinkedObjectType = "2";


                    //objGridColumn = objGrid.Columns.Item(1);
                    //objGridColumn.Editable = false;


                    //objGridColumn = objGrid.Columns.Item(2);
                    //objGridColumn.Editable = false;


                    //objGridColumn = objGrid.Columns.Item(3);
                    //objGridColumn.Editable = false;


                    //for (int i = 0; i < objDtRes.Rows.Count; i++)
                    //{
                    //    string strResult = objDtRes.GetValue("Result", i);
                    //    if (strResult == "OK")
                    //    {
                    //        objGrid.CommonSetting.SetCellBackColor(i + 1, 4, Color.Green.R | (Color.Green.G << 8) | (Color.Green.B << 16));

                    //    }
                    //    else
                    //    {
                    //        objGrid.CommonSetting.SetCellBackColor(i + 1, 4, Color.Red.R | (Color.Red.G << 8) | (Color.Red.B << 16));

                    //    }
                    //}






                    objGrid.AutoResizeColumns();




                    T1.B1.Base.UIOperations.Operations.stopProgressBar();

                    objItem = objForm.Items.Item("btnAdd");
                    objItem.Enabled = false;
                    objForm.Freeze(false);
                }
                else
                {
                    objItem = objForm.Items.Item("btnAdd");
                    objItem.Enabled = false;
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

        public static double getPercentFromCode(string WTCode)
        {
            double dblPrctBsAmnt = -1;
            SAPbobsCOM.WithholdingTaxCodes objWT = null;
            try
            {
                objWT = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                if(objWT.GetByKey(WTCode))
                {
                    dblPrctBsAmnt = objWT.BaseAmount;
                }
                else
                {
                    dblPrctBsAmnt = - 1;
                    _Logger.Error("Could not get Percent info");
                }

            }
            catch (COMException cOMException1)
            {
                _Logger.Error("", cOMException1);
                dblPrctBsAmnt = -1;
            }
            catch (Exception exception2)
            {
                _Logger.Error("", exception2);
                dblPrctBsAmnt = -1;
            }
            return dblPrctBsAmnt;
        }

        private static string addPurchaseInvoiceWTInternalRegistry(int DocEntry)
        {

            string objResult = "";
            Documents objDoc = null;
            CompanyService objCompanyService = null;
            GeneralService objEntryObject = null;
            GeneralData objEntryInfo = null;
            GeneralData objEntryLinesInfo = null;
            GeneralDataCollection objEntryLinesObject = null;
            GeneralDataParams objResultAdd = null;

            GeneralService objRelatedPartyObject = null;
            GeneralDataParams oFilter = null;
            GeneralData objRelatedPartyInfo = null;


            string strCardCode = "";
            int intDocEntry = -1;
            int intDocNum = -1;
            double dbDocTotal = 0;
            double dbBaseAmnt = 0;
            string strRelatedParty = "";
            string strDocType = "";


            int intJE = -1;
            try
            {
                objDoc = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
                if (objDoc.GetByKey(DocEntry))
                {

                    #region Temp


                    //objDoc = (SAPbobsCOM.Documents)B1Object;
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();

                    strCardCode = objDoc.CardCode;
                    intDocEntry = objDoc.DocEntry;
                    intDocNum = objDoc.DocNum;
                    strDocType = "18";
                    dbDocTotal = objDoc.DocTotal;
                    dbBaseAmnt = getBaseAmount(objDoc);
                    #region Get ThirdParty Info
                    objRelatedPartyObject = objCompanyService.GetGeneralService("BYB_T1RPAU01");//100
                    oFilter = objRelatedPartyObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oFilter.SetProperty("Code", objDoc.CardCode);
                    try
                    {
                        objRelatedPartyInfo = objRelatedPartyObject.GetByParams(oFilter);
                        strRelatedParty = objRelatedPartyInfo.GetProperty("Code");
                    }
                    catch (Exception er)
                    {
                        strRelatedParty = "";
                    }
                    #endregion

                    objEntryObject = objCompanyService.GetGeneralService("BYB_T1WHT400");
                    objEntryInfo = objEntryObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                    objEntryInfo.SetProperty("U_DOCNUM", intDocNum);
                    objEntryInfo.SetProperty("U_DOCENTRY", intDocEntry);
                    objEntryInfo.SetProperty("U_CARDCODE", strCardCode);
                    objEntryInfo.SetProperty("U_RELPARTY", strRelatedParty);
                    objEntryInfo.SetProperty("U_DOCTYPE", strDocType);
                    objEntryInfo.SetProperty("U_DOCTOTAL", dbDocTotal);
                    objEntryInfo.SetProperty("U_BASEAMNT", dbBaseAmnt);

                    intJE = objDoc.TransNum;

                    objEntryLinesObject = objEntryInfo.Child("BYB_T1WHT401");

                    SAPbobsCOM.WithholdingTaxData oWHTData = objDoc.WithholdingTaxData;
                    for (int i = 0; i < oWHTData.Count; i++)
                    {
                        oWHTData.SetCurrentLine(i);
                        SAPbobsCOM.WithholdingTaxCodes oWT = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                        if (oWT.GetByKey(oWHTData.WTCode))
                        {

                            objEntryLinesInfo = objEntryLinesObject.Add();
                            objEntryLinesInfo.SetProperty("U_JE", intJE);
                            objEntryLinesInfo.SetProperty("U_OPERTYPE", "WHT");
                            objEntryLinesInfo.SetProperty("U_OPER", "INI");
                            objEntryLinesInfo.SetProperty("U_CODE", oWHTData.WTCode);
                            objEntryLinesInfo.SetProperty("U_PERCENT", oWT.BaseAmount);

                            if (oWT.BaseType == WithholdingTaxCodeBaseTypeEnum.wtcbt_VAT)
                            {
                                objEntryLinesInfo.SetProperty("U_BASEAMNT", objDoc.VatSum);
                            }
                            else
                            {
                                objEntryLinesInfo.SetProperty("U_BASEAMNT", dbBaseAmnt);
                            }
                            objEntryLinesInfo.SetProperty("U_AMNT", oWHTData.TaxableAmount);
                            objEntryLinesInfo.SetProperty("U_DOCLINE", -1);
                        }
                    }

                    objResultAdd = objEntryObject.Add(objEntryInfo);
                    objResult = "OK. El registro se creo satisfactoriamente";

                    
                    #endregion Temp


                }


                else
                {
                    _Logger.Error("No se pudo recuperar la información del documento " + DocEntry.ToString());
                    objResult = "No se pudo recuperar la información del documento";
                }
            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception);
                objResult = exception.Message;
            }
            catch (Exception ex)
            {
                _Logger.Error("", ex);
                objResult = ex.Message;
            }
            return objResult;
        }

        private static string addCNPurchaseInvoiceWTInternalRegistry(int DocEntry)
        {

            string objResult = "";
            Documents objDoc = null;
            CompanyService objCompanyService = null;
            GeneralService objEntryObject = null;
            GeneralData objEntryInfo = null;
            GeneralData objEntryLinesInfo = null;
            GeneralDataCollection objEntryLinesObject = null;
            GeneralDataParams objResultAdd = null;

            GeneralService objRelatedPartyObject = null;
            GeneralDataParams oFilter = null;
            GeneralData objRelatedPartyInfo = null;


            string strCardCode = "";
            int intDocEntry = -1;
            int intDocNum = -1;
            double dbDocTotal = 0;
            double dbBaseAmnt = 0;
            string strRelatedParty = "";
            string strDocType = "";


            int intJE = -1;
            try
            {
                objDoc = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oPurchaseCreditNotes);
                if (objDoc.GetByKey(DocEntry))
                {

                    #region Temp


                    //objDoc = (SAPbobsCOM.Documents)B1Object;
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();

                    strCardCode = objDoc.CardCode;
                    intDocEntry = objDoc.DocEntry;
                    intDocNum = objDoc.DocNum;
                    strDocType = "19";
                    dbDocTotal = objDoc.DocTotal;
                    dbBaseAmnt = getBaseAmount(objDoc);
                    #region Get ThirdParty Info
                    objRelatedPartyObject = objCompanyService.GetGeneralService("BYB_T1RPAU01");//100
                    oFilter = objRelatedPartyObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oFilter.SetProperty("Code", objDoc.CardCode);
                    try
                    {
                        objRelatedPartyInfo = objRelatedPartyObject.GetByParams(oFilter);
                        strRelatedParty = objRelatedPartyInfo.GetProperty("Code");
                    }
                    catch (Exception er)
                    {
                        strRelatedParty = "";
                    }
                    #endregion

                    objEntryObject = objCompanyService.GetGeneralService("BYB_T1WHT400");
                    objEntryInfo = objEntryObject.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                    objEntryInfo.SetProperty("U_DOCNUM", intDocNum);
                    objEntryInfo.SetProperty("U_DOCENTRY", intDocEntry);
                    objEntryInfo.SetProperty("U_CARDCODE", strCardCode);
                    objEntryInfo.SetProperty("U_RELPARTY", strRelatedParty);
                    objEntryInfo.SetProperty("U_DOCTYPE", strDocType);
                    objEntryInfo.SetProperty("U_DOCTOTAL", dbDocTotal);
                    objEntryInfo.SetProperty("U_BASEAMNT", dbBaseAmnt);

                    intJE = objDoc.TransNum;

                    objEntryLinesObject = objEntryInfo.Child("BYB_T1WHT401");

                    SAPbobsCOM.WithholdingTaxData oWHTData = objDoc.WithholdingTaxData;
                    for (int i = 0; i < oWHTData.Count; i++)
                    {
                        oWHTData.SetCurrentLine(i);
                        SAPbobsCOM.WithholdingTaxCodes oWT = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oWithholdingTaxCodes);
                        if (oWT.GetByKey(oWHTData.WTCode))
                        {

                            objEntryLinesInfo = objEntryLinesObject.Add();
                            objEntryLinesInfo.SetProperty("U_JE", intJE);
                            objEntryLinesInfo.SetProperty("U_OPERTYPE", "WHT");
                            objEntryLinesInfo.SetProperty("U_OPER", "INI");
                            objEntryLinesInfo.SetProperty("U_CODE", oWHTData.WTCode);
                            objEntryLinesInfo.SetProperty("U_PERCENT", oWT.BaseAmount);

                            if (oWT.BaseType == WithholdingTaxCodeBaseTypeEnum.wtcbt_VAT)
                            {
                                objEntryLinesInfo.SetProperty("U_BASEAMNT", objDoc.VatSum);
                            }
                            else
                            {
                                objEntryLinesInfo.SetProperty("U_BASEAMNT", dbBaseAmnt);
                            }
                            objEntryLinesInfo.SetProperty("U_AMNT", oWHTData.TaxableAmount);
                            objEntryLinesInfo.SetProperty("U_DOCLINE", -1);
                        }
                    }

                    objResultAdd = objEntryObject.Add(objEntryInfo);
                    objResult = "OK. El registro se creo satisfactoriamente";


                    #endregion Temp


                }


                else
                {
                    _Logger.Error("No se pudo recuperar la información del documento " + DocEntry.ToString());
                    objResult = "No se pudo recuperar la información del documento";
                }
            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception);
                objResult = exception.Message;
            }
            catch (Exception ex)
            {
                _Logger.Error("", ex);
                objResult = ex.Message;
            }
            return objResult;
        }

        #endregion

    }
}
