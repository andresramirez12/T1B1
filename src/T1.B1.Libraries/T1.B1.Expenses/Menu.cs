using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.Expenses
{
    public class Menu
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Menu objMenuObject;

        private Menu()
        {
            objMenuObject = new Menu();
        }

        public static void addMenu(ref StringBuilder sr)
        {
            try
            {
                addExpensesMenu();
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

        public static void removeMenu(string MenuId)
        {


            try
            {
                MainObject.Instance.B1Application.Menus.RemoveEx(MenuId);
                

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

        public static void addExpensesMenu()
        {
            try
            {
                SAPbouiCOM.MenuCreationParams objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Cajas Menores";
                objMenu.UniqueID = "BYB_MCM001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;


                int count = MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MCM001"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Creación Caja Menor";
                objMenu.UniqueID = "BYB_MCM007";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MCM007"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Apertura Caja Menor";
                objMenu.UniqueID = "BYB_MCM002";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MCM002"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Cierre Caja Menor";
                objMenu.UniqueID = "BYB_MCM003";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MCM003"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.AddEx(objMenu);
                }


                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Registro Gasto Caja Menor";
                objMenu.UniqueID = "BYB_MCM004";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MCM004"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.AddEx(objMenu);
                }

                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Arqueo Caja Menor";
                //objMenu.UniqueID = "BYB_MCM005";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                //count = MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MCM005"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.AddEx(objMenu);
                //}

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Conceptos Caja Menor";
                objMenu.UniqueID = "BYB_MCM006";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MCM006"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MCM001").SubMenus.AddEx(objMenu);
                }






                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Legalizaciones";
                objMenu.UniqueID = "BYB_MEX001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                

                count = MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX001"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_M001").SubMenus.AddEx(objMenu);
                }
                

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Configuración";
                objMenu.UniqueID = "BYB_MEX002";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX002"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.AddEx(objMenu);
                }
                

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Herramientas";
                objMenu.UniqueID = "BYB_MEX003";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX003"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.AddEx(objMenu);
                }
                

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Informes";
                objMenu.UniqueID = "BYB_MEX004";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX004"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Clasificación Legalizaciones";
                objMenu.UniqueID = "BYB_MEX008";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX002").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX008"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX002").SubMenus.AddEx(objMenu);
                }


                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Conceptos";
                objMenu.UniqueID = "BYB_MEX005";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                
                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX002").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX005"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX002").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Tipos de Legalizaciones";
                objMenu.UniqueID = "BYB_MEX006";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX002").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX006"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX002").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Solicitud de Legalización";
                objMenu.UniqueID = "BYB_MEX007";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX007"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Aprobación de Solicitud";
                objMenu.UniqueID = "BYB_MEX009";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX009"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Desembolsos";
                objMenu.UniqueID = "BYB_MEX010";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX010"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.AddEx(objMenu);
                }

                objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "Legalización";
                objMenu.UniqueID = "BYB_MEX011";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                count = MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEX011"))
                {
                    MainObject.Instance.B1Application.Menus.Item("BYB_MEX001").SubMenus.AddEx(objMenu);
                }


                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Tipos de legalización";
                //objMenu.UniqueID = "BYB_MEXF";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //if (Settings._Main.recreateMenu)
                //{
                //    if (MainObject.Instance.B1Application.Menus.Exists("BYB_MEXF"))
                //    {
                //        removeMenu("BYB_MEXF");
                //    }
                //}
                //count = MainObject.Instance.B1Application.Menus.Item("BYB_MEXB").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEXF"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_MEXB").SubMenus.AddEx(objMenu);
                //}


                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Solicitud de legalización";
                //objMenu.UniqueID = "BYB_MEXG";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //if (Settings._Main.recreateMenu)
                //{
                //    if (MainObject.Instance.B1Application.Menus.Exists("BYB_MEXG"))
                //    {
                //        removeMenu("BYB_MEXG");
                //    }
                //}
                //count = MainObject.Instance.B1Application.Menus.Item("BYB_MEXA").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEXG"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_MEXA").SubMenus.AddEx(objMenu);
                //}


                //objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                //objMenu.String = "Registro de legalización";
                //objMenu.UniqueID = "BYB_MEXH";
                //objMenu.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //if (Settings._Main.recreateMenu)
                //{
                //    if (MainObject.Instance.B1Application.Menus.Exists("BYB_MEXH"))
                //    {
                //        removeMenu("BYB_MEXH");
                //    }
                //}
                //count = MainObject.Instance.B1Application.Menus.Item("BYB_MEXA").SubMenus.Count + 1;
                //objMenu.Position = count;
                //if (!MainObject.Instance.B1Application.Menus.Exists("BYB_MEXH"))
                //{
                //    MainObject.Instance.B1Application.Menus.Item("BYB_MEXA").SubMenus.AddEx(objMenu);
                //}











            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}
