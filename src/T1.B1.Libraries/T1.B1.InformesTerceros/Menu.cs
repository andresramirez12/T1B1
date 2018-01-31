using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.InformesTerceros
{
    public class Menu
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Menu objMenuObject;

        private Menu()
        {
            objMenuObject = new Menu();
        }


        public static void addITRMenu()
        {
            try
            {
                SAPbouiCOM.MenuCreationParams objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Informes Terceros";
                objMenu.UniqueID = "BYB_MITR01";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                int count = MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MITR01"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Balance por Terceros";
                objMenu.UniqueID = "BYB_MITR02";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MITR01").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MITR02"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MITR01").SubMenus.AddEx(objMenu);
                }

                

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
