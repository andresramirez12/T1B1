using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;
using System.Collections;
using System.Globalization;
using System.Windows.Forms;

namespace T1.B1.Shared
{
    public class Shared
    {
        static private Shared objShared = null;
    
        private Shared()
        {

        }

        static public void addSharedMenu()
        {
            SAPbouiCOM.MenuCreationParams objMenu = null;
            SAPbouiCOM.MenuItem objMenuItem = null;

            try
            {
                objMenu = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = B1.Shared.LocalizactionStrings.Default.T1MainFolderMenuDescription;
                objMenu.UniqueID = B1.Shared.InteractionId.Default.T1MainFolderMenuId;
                objMenu.Position = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Shared.InteractionId.Default.T1MainFolderMenuParentId).SubMenus.Count + 1;
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(B1.Shared.InteractionId.Default.T1MainFolderMenuId))
                {
                    objMenuItem = BYBB1MainObject.Instance.B1Application.Menus.Item(B1.Shared.InteractionId.Default.T1MainFolderMenuParentId).SubMenus.AddEx(objMenu);
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "Shared.addSharedMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "Shared.addSharedMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }
    }
}
