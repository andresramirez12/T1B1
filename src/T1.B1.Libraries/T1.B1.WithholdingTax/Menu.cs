using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.WithholdingTax
{
    public class Menu
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Menu objMenuObject;

        private Menu()
        {
            objMenuObject = new Menu();
        }

        
        public static void addWTMenu()
        {
            try
            {
                SAPbouiCOM.MenuCreationParams objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Autoretenciones";
                objMenu.UniqueID = "BYB_MWT01";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                int count = MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT01"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Configuración";
                objMenu.UniqueID = "BYB_MWT02";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MWT01").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT02"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MWT01").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Autoretenciones Faltantes";
                objMenu.UniqueID = "BYB_MWT03";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MWT01").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT03"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MWT01").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Cancelar Autoretenciones";
                objMenu.UniqueID = "BYB_MWT04";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MWT01").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT04"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MWT01").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Retenciones";
                objMenu.UniqueID = "BYB_MWT05";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT05"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Grupo de Municipios";
                objMenu.UniqueID = "BYB_MWT06";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MWT05").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT06"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MWT05").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Registro de operaciones faltantes";
                objMenu.UniqueID = "BYB_MWT07";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MWT05").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT07"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MWT05").SubMenus.AddEx(objMenu);
                }


                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Utilidades";
                //objMenu.UniqueID = "BYB_MWT02";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                //count = MainObject.Instance.B1Application.Menus.Item("BYB_MWT01").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT02"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_MWT02").SubMenus.AddEx(objMenu);
                //}

                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Calcular faltantes";
                //objMenu.UniqueID = "BYB_MWT03";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //count = MainObject.Instance.B1Application.Menus.Item("BYB_MWT02").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT03"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_MWT03").SubMenus.AddEx(objMenu);
                //}

                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Reversar";
                //objMenu.UniqueID = "BYB_MWT04";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //count = MainObject.Instance.B1Application.Menus.Item("BYB_MWT02").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT04"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_MWT04").SubMenus.AddEx(objMenu);
                //}



                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Autoretenciones";
                //objMenu.UniqueID = "BYB_MWT06";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //count = MainObject.Instance.B1Application.Menus.Item("BYB_MWT05").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MWT06"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_MWT06").SubMenus.AddEx(objMenu);
                //}

                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Retenciones";
                //objMenu.UniqueID = "BYB_M008";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                //count = MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_M008"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.AddEx(objMenu);
                //}

                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Retención por artículo";
                //objMenu.UniqueID = "BYB_M009";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //count = MainObject.Instance.B1Application.Menus.Item("BYB_M008").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_M009"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_M008").SubMenus.AddEx(objMenu);
                //}




            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
