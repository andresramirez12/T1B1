using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.ReletadParties
{
    public class Menu
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Menu objMenuObject;

        private Menu()
        {
            objMenuObject = new Menu();
        }

        

        public static void addThirdParitesMenu()
        {
            try
            {
                SAPbouiCOM.MenuCreationParams objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Terceros relacionados";
                objMenu.UniqueID = "BYB_MRP01";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                int count = MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MRP01"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Utilidades";
                objMenu.UniqueID = "BYB_MRP02";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MRP01").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MRP02"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MRP01").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Crear faltantes";
                objMenu.UniqueID = "BYB_MRP03";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MRP02").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MRP03"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MRP02").SubMenus.AddEx(objMenu);
                }

                

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Configuración";
                objMenu.UniqueID = "BYB_MRP04";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MRP01").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MRP04"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MRP01").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Tereceros Relacionados";
                objMenu.UniqueID = "BYB_MRP05";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MRP04").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MRP05"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MRP04").SubMenus.AddEx(objMenu);
                }

                

                
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
