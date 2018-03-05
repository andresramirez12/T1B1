using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;

namespace T1.Classes
{
    class BYBConnection
    {

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
            objApplication.ItemEvent -= objHandler1;
            objApplication.MenuEvent -= objHandler2;
        }

        public void B1Connect(bool stopEvents)
        {
            SAPbouiCOM.SboGuiApi objGUIApi = null;
            
            SAPbobsCOM.Company objCompany = null;

            try
            {
                objGUIApi = new SAPbouiCOM.SboGuiApi();
                objGUIApi.Connect((string)BYBCache.Instance.getFromCache(T1.Settings.CacheItemNames.Default.ConnectionString));
                objApplication = objGUIApi.GetApplication(-1);
                if (T1.Settings.ApplicationConfiguration.Default.useCompatibilityConnection)
                {
                    objCompany = new SAPbobsCOM.Company();
                    string strContextCookie = objCompany.GetContextCookie();
                    string strConnectionString = objApplication.Company.GetConnectionContext(strContextCookie);
                    if (objCompany.SetSboLoginContext(strConnectionString) == 0)
                    {
                        if (objCompany.Connect() != 0)
                        {
                            BYBExceptionHandling.reportException("", "B1ConnectionClass.B1Connect", new Exception(objCompany.GetLastErrorCode() + "::" + objCompany.GetLastErrorDescription()), 2, System.Diagnostics.EventLogEntryType.Error);
                        }
                    }
                    else
                    {
                        BYBExceptionHandling.reportException("", "B1ConnectionClass.B1Connect", new Exception(objCompany.GetLastErrorCode() + "::" + objCompany.GetLastErrorDescription()), 3, System.Diagnostics.EventLogEntryType.Error);
                    }

                
                }
                else
                {

                    objApplication.EventLevel = SAPbouiCOM.BoEventLevelType.elf_Both;
                    objCompany = new SAPbobsCOM.Company();
                    objCompany.Application = objApplication;

                    objCompany.Connect();
                    
                        //= objApplication.Company.GetDICompany();
                    
                }

                blConnected = objCompany.Connected;
                if (blConnected)
                {
                    BYBB1MainObject.Instance.B1Application = objApplication;
                    BYBB1MainObject.Instance.B1Company = objCompany;

                    SAPbobsCOM.CompanyService objServ =  objCompany.GetCompanyService();
                    SAPbobsCOM.AdminInfo objAdmInfo = objServ.GetAdminInfo();
                    string strDecimalSeparator = objAdmInfo.DecimalSeparator;
                    string strThousendSeparator = objAdmInfo.ThousandsSeparator;

                    BYBCache.Instance.addToCache("decimales", strDecimalSeparator ,BYBCache.objCachePriority.NotRemovable);
                    BYBCache.Instance.addToCache("miles", strThousendSeparator, BYBCache.objCachePriority.NotRemovable);

                        

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
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "B1ConnectionClass.B1Connect", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "B1ConnectionClass.B1Connect", er, 1, System.Diagnostics.EventLogEntryType.Error);
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
