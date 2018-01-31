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

namespace T1.B1.InformesTerceros
{
    public class Operations
    {
        private static Operations InformesTerceros;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private Operations()
        {

        }


        public static void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            if (InformesTerceros == null)
            {
                InformesTerceros = new Operations();
            }

            BubbleEvent = true;
            try
            {
                if (pVal.MenuUID == "BYB_MITR02"
                    && !pVal.BeforeAction)
                {
                    BalanceTerceros.getTransactionList();
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

    }
}
