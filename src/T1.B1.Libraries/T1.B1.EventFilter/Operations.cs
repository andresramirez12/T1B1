using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using log4net;

namespace T1.B1.EventFilter
{
    public class Operations
    {
        private SAPbouiCOM.Application objApplication = null;
        public bool objStatus = false;
        private static readonly ILog _Logger = LogManager.GetLogger("T1.B1.EventFilter");


        public bool Status
        {
            get
            {

                return objStatus;
            }
        }

        public Operations()
        {
            SAPbouiCOM.EventFilters objFilterCollection = null;
            SAPbouiCOM.EventFilter objFilter = null;
            try
            {
                objApplication = T1.B1.MainObject.Instance.B1Application;
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

                T1.B1.MainObject.Instance.B1Application = objApplication;
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

    }
}
