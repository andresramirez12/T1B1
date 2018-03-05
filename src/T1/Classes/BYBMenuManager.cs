using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace T1.Classes
{
    class BYBMenuManager
    {

        private SAPbouiCOM.Application objApplication = null;
        public bool objStatus = false;
        
        public bool Status
        {
            get
            {
                return objStatus;
            }
        }

        public BYBMenuManager()
        {
            try
            {
                objApplication = BYBB1MainObject.Instance.B1Application;
                addMenu();
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "MenuManagerClass.MenuManagerClass", er, 11, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.MenuManagerClass", er, 11, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        public void addMenu()
        {
            
            
            try
            {
                T1.B1.Shared.Shared.addSharedMenu();
                B1.Expenses.Expenses.addMainMenu();
                B1.RelatedParties.RelatedParties.addMenu();
                B1.WithholdingTax.WithholdingTax.addMenu();

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "MenuManagerClass.addMenu", er, 12, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.addMenu", er, 12, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        public void removeMenu()
        {
            

            try
            {


                
             

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 13, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 13, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        


        
    }
}
