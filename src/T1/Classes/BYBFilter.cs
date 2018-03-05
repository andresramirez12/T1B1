using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace T1.Classes
{
    class BYBFilter
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

        public BYBFilter()
        {
            SAPbouiCOM.EventFilters objFilterCollection = null;
            SAPbouiCOM.EventFilter objFilter = null;
            try
            {
                objApplication = BYBB1MainObject.Instance.B1Application;
                objFilterCollection = objApplication.GetFilter();
                objFilterCollection.Reset();
                
                //objFilter = objFilterCollection.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
                
                //objFilter = objFilterCollection.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
                //objFilter.AddEx(B1.BPImpairment.InteractionId.Default.frmBPFilteringFormType);
                //objFilter.AddEx(B1.BPImpairment.InteractionId.Default.frmBPResultFormType);
                //objFilter.AddEx(B1.InventoryImpairment.InteractionId.Default.frmItemFilteringFormType);
                //objFilter.AddEx(B1.InventoryImpairment.InteractionId.Default.frmItemResultFormType);
                
                //objFilter = objFilterCollection.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
                //objFilter.AddEx(B1.FSNotes.InteractionId.Default.fsNotesFormType);

                //objFilter = objFilterCollection.Add(SAPbouiCOM.BoEventTypes.et_PRINT_LAYOUT_KEY);
                //objFilter.AddEx(B1.FSNotes.InteractionId.Default.fsNotesFormType);

                objFilter = objFilterCollection.Add(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS);

               
                
                
                objStatus = true;

                BYBB1MainObject.Instance.B1Application = objApplication;
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "FilterClass.FilterClass", er, 10, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "FilterClass.FilterClass", er, 10, System.Diagnostics.EventLogEntryType.Error);
            }


        }

    }
}
