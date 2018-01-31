using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.MenuManager
{
    public class Operations
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static Operations objMenuManager;
        

        

        private Operations()
        {
            objMenuManager = new Operations();
        }

        public static void addMenu()
        {


            try
            {
                //StringBuilder sbMenu = new StringBuilder();
                //sbMenu.Append(Properties.Resources.OpenMenuString);
                //addMainMenu(ref sbMenu);
                addMainMenu();
                T1.B1.WithholdingTax.Menu.addWTMenu();
                T1.B1.ReletadParties.Menu.addThirdParitesMenu();
                T1.B1.Expenses.Menu.addExpensesMenu();
                T1.B1.InformesTerceros.Menu.addITRMenu();

                //T1.B1.WithholdingTax.Menu.addMenu(ref sbMenu);
                
                //sbMenu.Append(Properties.Resources.CloseMenuString);
                //string strFinalString = sbMenu.ToString();
                //strFinalString = strFinalString.Replace("[--BasePath--]", AppDomain.CurrentDomain.BaseDirectory);
                
                //MainObject.Instance.B1Application.LoadBatchActions(ref strFinalString);
                //string strResult = MainObject.Instance.B1Application.GetLastBatchResults();
                //sbMenu = null;

                //T1.B1.Expenses.Menu.addMenu(ref sbMenu);
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

        public static void removeMenu()
        {


            try
            {





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

        private static void addMainMenu(ref StringBuilder sbMenu)
        {

            sbMenu.Append(Properties.Resources.Menu);
        }

        private static void addMainMenu()
        {
            try
            {



                SAPbouiCOM.MenuCreationParams objMenu = MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                objMenu.String = "T1";
                objMenu.Image = AppDomain.CurrentDomain.BaseDirectory + "Original\\T1.png";
                objMenu.UniqueID = "BYB_M001";
                objMenu.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                int count = MainObject.Instance.B1Application.Menus.Item("43520").SubMenus.Count + 1;
                objMenu.Position = count;
                if (!MainObject.Instance.B1Application.Menus.Exists("BYB_M001"))
                {
                    MainObject.Instance.B1Application.Menus.Item("43520").SubMenus.AddEx(objMenu);
                }
                SAPbouiCOM.IMenuItem objM = MainObject.Instance.B1Application.Menus.Item("43520");
                string strTest = objM.SubMenus.GetAsXML();
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }

    }
}
