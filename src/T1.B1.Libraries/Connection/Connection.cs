using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using log4net;

namespace T1.B1.Connection
{
    public class Class
    {

        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private bool blConnected = false;
        private SAPbouiCOM._IApplicationEvents_ItemEventEventHandler objHandler1;
        private SAPbouiCOM._IApplicationEvents_MenuEventEventHandler objHandler2;
        SAPbouiCOM.Application objApplication = null;
        

        public bool Connected
        {
            get
            {
                return blConnected;
            }
        }

        public void clearStopEvents()
        {
            try
            {
                objApplication.ItemEvent -= objHandler1;
                objApplication.MenuEvent -= objHandler2;
            }
            catch(Exception er)
            {
                _Logger.Error(er);
            }
        }

        
        public void B1Connect(bool stopEvents)
        {
            SAPbouiCOM.SboGuiApi objGUIApi = null;

            SAPbobsCOM.Company objCompany = null;

            try
            {
                _Logger.Debug("Starting SAP Connection");
                objGUIApi = new SAPbouiCOM.SboGuiApi();
                _Logger.Debug("Retrieving connection string from Cache");
                objGUIApi.Connect((string)T1.CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.connStringCacheName));
                _Logger.Debug("Connectong to current company");
                objApplication = objGUIApi.GetApplication(-1);
                objApplication.EventLevel = SAPbouiCOM.BoEventLevelType.elf_Both;
                if (Settings._Main.useCompatibilityConnection)
                {
                    _Logger.Info("Connecting to DI API using Compatibility mode (cookies)");
                    objCompany = new SAPbobsCOM.Company();
                    string strContextCookie = objCompany.GetContextCookie();
                    string strConnectionString = objApplication.Company.GetConnectionContext(strContextCookie);
                    _Logger.Debug("Setting Login Context");
                    if (objCompany.SetSboLoginContext(strConnectionString) == 0)
                    {
                        if (objCompany.Connect() != 0)
                        {
                            string strError = objCompany.GetLastErrorCode() + ":" + objCompany.GetLastErrorDescription();
                            _Logger.Error(strError);
                        }
                        else
                        {
                            bool isHANA = objCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB ? true : false;
                            T1.CacheManager.CacheManager.Instance.addToCache(T1.CacheManager.Settings._Main.isHANACacheName, isHANA, CacheManager.CacheManager.objCachePriority.NotRemovable);
                        }
                    }
                    else
                    {
                        string strError = objCompany.GetLastErrorCode() + ":" + objCompany.GetLastErrorDescription();
                        _Logger.Error(strError);
                    }


                }
                else
                {
                    _Logger.Info("Connecting to DI API using shared memory library");
                    objCompany = new SAPbobsCOM.Company();
                    if (Settings._Main.useCompanyApplication)
                    {
                        objCompany.Application = objApplication;

                        if (objCompany.Connect() != 0)
                        {
                            string strError = objCompany.GetLastErrorCode() + ":" + objCompany.GetLastErrorDescription();
                            _Logger.Error(strError);
                        }
                        else
                        {
                            bool isHANA = objCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB ? true : false;
                            T1.CacheManager.CacheManager.Instance.addToCache(T1.CacheManager.Settings._Main.isHANACacheName, isHANA, CacheManager.CacheManager.objCachePriority.NotRemovable);
                        }
                    }
                    else
                    {
                        objCompany = objApplication.Company.GetDICompany();
                        bool isHANA = objCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB ? true : false;
                        T1.CacheManager.CacheManager.Instance.addToCache(T1.CacheManager.Settings._Main.isHANACacheName, isHANA, CacheManager.CacheManager.objCachePriority.NotRemovable);
                    }
                }
                _Logger.Debug("Checking connection status");
                blConnected = objCompany.Connected;
                if (blConnected)
                {
                    _Logger.Debug("Connected to company " + objCompany.CompanyName);
                    objCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                    objCompany.XMLAsString = true;
                    T1.B1.MainObject.Instance.B1Application = objApplication;
                    T1.B1.MainObject.Instance.B1Company = objCompany;

                    SAPbobsCOM.CompanyService objServ = objCompany.GetCompanyService();
                    SAPbobsCOM.AdminInfo objAdmInfo = objServ.GetAdminInfo();

                    T1.B1.InternalClasses.AdminInfo B1AdmInfo = new InternalClasses.AdminInfo();

                    B1AdmInfo.DecimalSeparator = objAdmInfo.DecimalSeparator;
                    B1AdmInfo.ThousandsSeparator = objAdmInfo.ThousandsSeparator;
                    B1AdmInfo.DateSeparator = objAdmInfo.DateSeparator;
                    B1AdmInfo.SystemCurrency = objAdmInfo.SystemCurrency;
                    B1AdmInfo.RateAccuracy = objAdmInfo.RateAccuracy;
                    B1AdmInfo.QueryAccuracy = objAdmInfo.QueryAccuracy;
                    B1AdmInfo.AccuracyofQuantities = objAdmInfo.AccuracyofQuantities;
                    B1AdmInfo.PercentageAccuracy = objAdmInfo.PercentageAccuracy;
                    B1AdmInfo.PriceAccuracy = objAdmInfo.PriceAccuracy;
                    B1AdmInfo.TotalsAccuracy = objAdmInfo.TotalsAccuracy;
                    B1AdmInfo.LocalCurrency = objAdmInfo.LocalCurrency;

                    B1AdmInfo.DateTemplate = objAdmInfo.DateTemplate;
                    B1AdmInfo.DisplayCurrencyontheRight = objAdmInfo.DisplayCurrencyontheRight == SAPbobsCOM.BoYesNoEnum.tYES ? true : false;
                    B1AdmInfo.FederalTaxID = objAdmInfo.FederalTaxID;
                    B1AdmInfo.MeasuringAccuracy = objAdmInfo.MeasuringAccuracy;


                    T1.B1.MainObject.Instance.B1AdminInfo = B1AdmInfo;

                }
                if (stopEvents)
                {


                    objHandler1 = new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(objApplication_ItemEvent);
                    objHandler2 = new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(objApplication_MenuEvent);

                    objApplication.ItemEvent += objHandler1;
                    objApplication.MenuEvent += objHandler2;
                }

            }
            catch (COMException comEx)
            {

                //Exception er = new Exception(Convert.ToString(comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("", comEx);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }






        }

        
        void objApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = false;
        }

        void objApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = false;
        }




    }
}
