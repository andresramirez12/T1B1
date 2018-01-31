using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;

using System.Collections;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;



namespace T1.B1.MediosMagneticos
{
    public class Operations
    {
        private static Operations objMediosMagneticos;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private Operations()
        {

        }
        public static void formDataAddEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool blBubbleEvent)
        {
            if(objMediosMagneticos == null)
            {
                objMediosMagneticos = new Operations();
            }

            blBubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == "133" &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    //SelfWithholdingTax.addSelfWithHoldingTax(BusinessObjectInfo);
                }
            }
            catch (COMException COMException)
            {
                _Logger.Error("", COMException);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if(pVal.MenuUID == "BYB_M005"
                    && !pVal.BeforeAction)
                {
                    //SelfWithholdingTax.loadCancelSWTaxForm();
                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {


            BubbleEvent = true;
            try
            {
                if(pVal.FormTypeEx == Settings._SelfWithHoldingTax.CancelFormUID
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "btnGet"
                    )
                {
                    //SelfWithholdingTax.getPostedSWTaxDocuments(FormUID, pVal);
                }

                if (!pVal.BeforeAction
                    && pVal.EventType != SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    && pVal.ItemUID == "grdSWT"
                    && pVal.FormTypeEx == Settings._SelfWithHoldingTax.CancelFormUID)
                {
                    T1.B1.Base.UIOperations.Operations.toggleSelectCheckBox(pVal, "dtSelfWT", "1");
                }

                if (!pVal.BeforeAction
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    
                    && pVal.ItemUID == "btnCalc"
                    && pVal.FormTypeEx == Settings._SelfWithHoldingTax.CancelFormUID)
                {
                     //WithholdingTax.SelfWithholdingTax.cancelPostedTaxDocuments(FormUID, pVal);
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



    }
}
