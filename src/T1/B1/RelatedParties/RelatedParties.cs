using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using T1.Classes;
using System.Xml;
using System.Runtime.InteropServices;
using System.Collections;
using System.Globalization;
using System.Windows.Forms;

namespace T1.B1.RelatedParties
{
    public class RelatedParties
    {

         static private RelatedParties objRelatedParties = null;

        private RelatedParties()
        {

        }

        static public void addMenu()
        {
            string strMenuDescription = "";
            string strMenuId = "";

            if (objRelatedParties == null)
                objRelatedParties = new RelatedParties();

            try
            {


                strMenuId = "BYBRPMN01";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Terceros Relacionados";
                    
                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                    int intTotal = BYBB1MainObject.Instance.B1Application.Menus.Item("8704").SubMenus.Count;
                    objMenuCreationParams.Position = intTotal;

                    BYBB1MainObject.Instance.B1Application.Menus.Item("8704").SubMenus.AddEx(objMenuCreationParams);
                }

                strMenuId = "BYBRPMN02";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Crear Terceros Faltantes";

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intTotal = BYBB1MainObject.Instance.B1Application.Menus.Item("BYBRPMN01").SubMenus.Count;
                    objMenuCreationParams.Position = intTotal;

                    BYBB1MainObject.Instance.B1Application.Menus.Item("BYBRPMN01").SubMenus.AddEx(objMenuCreationParams);
                }


                strMenuId = "BYBRPMN03";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Terceros Relacionados";


                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                    int intTotal = BYBB1MainObject.Instance.B1Application.Menus.Item("43526").SubMenus.Count + 1;
                    objMenuCreationParams.Position = intTotal;


                    BYBB1MainObject.Instance.B1Application.Menus.Item("43526").SubMenus.AddEx(objMenuCreationParams);
                }

                strMenuId = "BYBRPMN04";
                if (!BYBB1MainObject.Instance.B1Application.Menus.Exists(strMenuId))
                {
                    strMenuDescription = "Terceros Relacionados";

                    SAPbouiCOM.MenuCreationParams objMenuCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objMenuCreationParams.String = strMenuDescription;
                    objMenuCreationParams.UniqueID = strMenuId;
                    objMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    int intTotal = BYBB1MainObject.Instance.B1Application.Menus.Item("BYBRPMN03").SubMenus.Count;
                    objMenuCreationParams.Position = intTotal;

                    BYBB1MainObject.Instance.B1Application.Menus.Item("BYBRPMN03").SubMenus.AddEx(objMenuCreationParams);
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.addMainMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.addMainMenu", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

        }

        static public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form objForm = null;


            try
            {
                if (objRelatedParties == null)
                    objRelatedParties = new RelatedParties();

                if (!pVal.BeforeAction)
                {
                    if (pVal.MenuUID == "BYBRPMN04")
                    {
                        SAPbouiCOM.FormCreationParams objFormCreationParams = null;
                        objFormCreationParams = BYBB1MainObject.Instance.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                        objFormCreationParams.XmlData = localForm("BYBRP001");
                        objFormCreationParams.FormType = "BYB_T1RP001";
                        objFormCreationParams.UniqueID = Guid.NewGuid().ToString().Substring(1, 20);
                        objForm = BYBB1MainObject.Instance.B1Application.Forms.AddEx(objFormCreationParams);
                        objForm.Visible = true;
                    }

                    if (pVal.MenuUID == "BYBRPMN02")
                    {
                        createRelatedParties();
                    }

                    

                }
           
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static private string localForm(string strFormId)
        {
            string strResult = "";

            if (objRelatedParties == null)
                objRelatedParties = new RelatedParties();

            try
            {


                if (strFormId == "BYBRP001")
                {
                    strResult = B1.RelatedParties.Resources.RelatedParties.BYBRP001;

                }



            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return strResult;

        }

        static private void createRelatedParties()
        {

            SAPbobsCOM.Recordset objRecordset = null;
            string strSql = "";
            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oRelatedPartiesService = null;
            SAPbobsCOM.GeneralData oSReletadPartiesData = null;
            SAPbobsCOM.GeneralDataParams oRelatedPartiesParams = null;


            string strCardCode = "";
            string strCardName = "";
            string strCardType = "";
            string strLicTradNum = "";

            int intCounter = 0;
            string strMessage = "";

            try
            {
                objRecordset = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                strSql = T1.B1.RelatedParties.Resources.dbQueries.missingBPQuery;
                objRecordset.DoQuery(strSql);


                if (objRecordset.RecordCount > 0)
                {
                    oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                    objRecordset.MoveFirst();
                    while (!objRecordset.EoF)
                    {
                        strCardCode = objRecordset.Fields.Item("CardCode").Value;
                        strCardName = objRecordset.Fields.Item("CardName").Value;
                        strCardType = objRecordset.Fields.Item("CardType").Value;
                        strLicTradNum = objRecordset.Fields.Item("LicTradNum").Value;



                        oRelatedPartiesService = oCompanyService.GetGeneralService("BYB_T1RPAU01");
                        oSReletadPartiesData = oRelatedPartiesService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                        oSReletadPartiesData.SetProperty("Code", strCardCode);
                        oSReletadPartiesData.SetProperty("U_CardCode", strCardCode);
                        oSReletadPartiesData.SetProperty("Name", strCardName);
                        oSReletadPartiesData.SetProperty("U_LegalName", strCardName);
                        oSReletadPartiesData.SetProperty("U_IdNum", strLicTradNum);
                        if (strCardType == "C")
                        {
                            oSReletadPartiesData.SetProperty("U_RelType2", "Y");
                        }
                        else if (strCardType == "S")
                        {
                            oSReletadPartiesData.SetProperty("U_RelType5", "Y");
                        }

                        oSReletadPartiesData.SetProperty("U_TipDoc", "NT");

                        oRelatedPartiesParams = oRelatedPartiesService.Add(oSReletadPartiesData);
                        if (oRelatedPartiesParams != null && (oRelatedPartiesParams.GetProperty("Code") == strCardCode))
                        {
                            intCounter++;
                        }

                        objRecordset.MoveNext();
                    }

                    strMessage = "Se crearon " + intCounter.ToString() + " Terceros Relacionados correctamente";




                }
                else
                {
                    strMessage = "Todos los Socios de Negocio estan creados como Terceros Relacionados";
                }
                BYBB1MainObject.Instance.B1Application.MessageBox(strMessage);
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "BPFiltering.localForm", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }


            /*
            static public void formDataAddEvent(SAPbouiCOM.BusinessObjectInfo pVal, out bool blBubbleEvent)
            {
                bool blOpenForm = false;
                Hashtable objWTHash = new Hashtable();
                blBubbleEvent = true;
                SAPbouiCOM.Form objForm = null;
                SAPbobsCOM.Documents objPurch = null;

                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                SAPbobsCOM.GeneralData oChild = null;
                SAPbobsCOM.GeneralDataCollection oChildren = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                SAPbobsCOM.CompanyService oCompanyService = null;

                try
                {
                    objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);

                    WTCalculation(objForm, out blOpenForm);
                    objWTHash = BYBCache.Instance.getFromCache("WTCalc" + pVal.FormUID);
                    if(objWTHash != null)
                    {
                        objPurch = BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);


                        oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                        oGeneralService = oCompanyService.GetGeneralService("BYB_T1WHT200");
                        oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                        oChildren = oGeneralData.Child("BYB_T1WHT201");

                        foreach (string strKey in objWTHash.Keys)
                        {
                            oChild = oChildren.Add();
                            ArrayList objArray = (ArrayList)objWTHash[strKey];
                            oChild.SetProperty("U_WTCode", strKey);
                            oChild.SetProperty("U_WTName", objArray[0]);
                            oChild.SetProperty("U_DbValue", objArray[1]);
                            oChild.SetProperty("U_WTTaxBase", objArray[2]);
                            oChild.SetProperty("U_WTValue", objArray[3]);
                            oChild.SetProperty("U_Account", objArray[4]);
                        }

                        oGeneralData.SetProperty("Code", );
                        oGeneralData.SetProperty("U_Type", "P");
                        oGeneralData.SetProperty("U_DocEntry", dbTotal);

                        oGeneralParams = oGeneralService.Add(oGeneralData);
                        int intNumber = oGeneralParams.GetProperty("DocEntry");
                        if (intNumber > 0)
                        {
                            BYBB1MainObject.Instance.B1Application.MessageBox(MessageStrings.Default.impairmentDoneMessage + "Asiento número: " + BYBB1MainObject.Instance.B1Company.GetNewObjectKey());

                        }
                        else
                        {
                            BYBB1MainObject.Instance.B1Application.MessageBox("Se produjo un error al registrar el detalle de los socios de negocio. El deterioro se contabilizó con el asiento número: " + BYBB1MainObject.Instance.B1Company.GetNewObjectKey());

                        }
                    }




                }
                catch (COMException comEx)
                {
                    Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                    BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);

                }
                catch (Exception er)
                {
                    BYBExceptionHandling.reportException(er.Message, "BPFiltering.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);

                }
            } 
            */


        }
}
