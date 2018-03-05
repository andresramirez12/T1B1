
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.IO;

namespace T1.Classes
{
    class BYBEventManager
    {

        private SAPbouiCOM.Application objApplication = null;
        public bool objStatus = false;
        
        
        public bool Status
        {
            get
            {
                return objStatus;
            }
        }

        public BYBEventManager()
        {
            try
            {
                objApplication = BYBB1MainObject.Instance.B1Application;   
                objApplication.AppEvent += objApplication_AppEvent;

                objApplication.EventLevel = SAPbouiCOM.BoEventLevelType.elf_Both;


                ///RightClick events per Module
                objApplication.RightClickEvent += T1.B1.WithholdingTax.WithholdingTax.RightClickEvent;
                objApplication.RightClickEvent += T1.B1.Expenses.Expenses.RightClickEvent;

                ///Menu Events per Module
                objApplication.MenuEvent += T1.B1.Expenses.Expenses.MenuEvent;
                objApplication.MenuEvent += T1.B1.RelatedParties.RelatedParties.MenuEvent;
                objApplication.MenuEvent += T1.B1.WithholdingTax.WithholdingTax.MenuEvent;


                //ItemEvents per Module
                objApplication.ItemEvent += T1.B1.WithholdingTax.WithholdingTax.ItemEvent;
                objApplication.ItemEvent += T1.B1.Expenses.Expenses.ItemEvent;



                ///UDO Event per Module Can I Inercept the event to change the XML?
                //objApplication.UDOEvent += T1.B1.WithholdingTax.WithholdingTax.UDOEvent;


                ///Move this to each class and then to dll for easy update without versioning

                objApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(objApplication_ItemEvent) ;

                objApplication.FormDataEvent += ObjApplication_FormDataEvent;
                objApplication.FormDataEvent += B1.WithholdingTax.InternalRegistry.InternalRegistry.ObjApplication_FormDataEvent;
                objApplication.FormDataEvent += T1.B1.Expenses.Expenses.LoadDataEvent;
                
                
                
                
                
                objStatus = true;


                BYBB1MainObject.Instance.B1Application = objApplication;

                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }


        }

        

        private void ObjApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {

                

                #region Add SelfWithHolding
                if(BusinessObjectInfo.ActionSuccess && BusinessObjectInfo.FormTypeEx=="133" && !BusinessObjectInfo.BeforeAction && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD )
                {
                    B1.WithholdingTax.WithholdingTax.formDataAddEvent(BusinessObjectInfo, out BubbleEvent);
                }
                #endregion Add SelfWithHolding





                //#region FormDataAdd PurchaseOrder
                //if (BusinessObjectInfo.BeforeAction && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.FormTypeEx=="141")
                //{
                //    B1.WithholdingTax.WithholdingTax.formDataAddEvent(BusinessObjectInfo, out BubbleEvent);
                //}
                //#endregion FormDataAdd PurchaseOrder
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 14, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        void objApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
          
           
            BubbleEvent = true;
            try
            {
                
                
                
                #region Event manager for ChooseFromList
                if(pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction && pVal.FormTypeEx=="9999" && pVal.ItemUID == "5")
                {
                    SAPbouiCOM.Form objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.DBDataSource oDS = objForm.DataSources.DBDataSources.Item(0);
                    string oDSName = oDS.TableName;
                    if(oDSName == "@BYB_T1EXP100")
                    {
                        BubbleEvent = false;
                        BYBB1MainObject.Instance.B1Application.ActivateMenuItem("BYBT1mnu02");
                        objForm.Close();
                    }
                    objForm = null;
                    oDS = null;
                }
                #endregion ChooseFromList Add Concept


                

                #region WithHoldingTax

                //if(pVal.FormTypeEx == "133" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                //{
                //    SAPbouiCOM.Form objForm = null;
                //    objForm = BYBB1MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                //    SAPbouiCOM.DBDataSource oDS = objForm.DataSources.DBDataSources.Item("INV5");
                //    BYBB1MainObject.Instance.B1Application.MessageBox("Size: "+oDS.Size.ToString());
                //    BYBB1MainObject.Instance.B1Application.MessageBox("Offset: " + oDS.Offset.ToString());
                //    oDS.InsertRecord(1);

                //}

                #endregion


                #region Get money on load or on combobox select 
                //if (pVal.FormTypeEx == "141" && !pVal.BeforeAction)
                //{
                //    if(pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                //    {
                //        if (pVal.ItemUID == "70" || pVal.ItemUID == "63")
                //        {

                //            B1.WithholdingTax.WithholdingTax.getMoneyFromForm(pVal);
                //        }
                //    }
                //    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                //    {
                        

                //            B1.WithholdingTax.WithholdingTax.getMoneyFromForm(pVal);
                        
                //    }
                //}
                #endregion Get money on load or on combobox select
                #region Configure Purchase Invoice Withholding Tax

                //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DRAW && pVal.FormTypeEx == "141" )
                //{
                //    if(pVal.BeforeAction)
                //    {
                //        //B1.WithholdingTax.WithholdingTax.initSystemForm(pVal);
                //    }
                //}

                #endregion Configure Purchase Invoice Withholding Tax

                #region Get all Withholding taxes from a BP
                //if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST || pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && pVal.FormTypeEx == "141" && pVal.ItemUID == "4")
                //{
                //    if (!pVal.BeforeAction)
                //    {

                        
                //        B1.WithholdingTax.WithholdingTax.getWTCodesForBP(pVal);
                //    }
                //}
                #endregion Get all Withholding taxes from a BP

                

                //TODO Add other events that may update document total to show

                //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.FormTypeEx == "141")
                //{
                //    if(pVal.BeforeAction)
                //    {
                //        //TODO Add here all controls that need validation to show total of Withholding Tax
                //    }
                //}


                if(pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "BYB_WTWTLB" && pVal.FormTypeEx == "141")
                {
                    //if(pVal.BeforeAction)
                    //{
                    //    BubbleEvent = false;
                    //    B1.WithholdingTax.WithholdingTax.linkedButtonClick(pVal);
                    //}

                }

                //if ((
                //    pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST || 
                //    pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                //    ) && pVal.FormTypeEx == "141" && pVal.ItemUID == "38")
                //{
                //    if (!pVal.BeforeAction)
                //    {
                //        if (pVal.ColUID == "1" ||
                //            pVal.ColUID == "11" ||
                //            pVal.ColUID == "14" ||
                //            pVal.ColUID == "15" ||
                //            pVal.ColUID == "174")
                //        {

                //            bool blOpenWindow = false;
                //            B1.WithholdingTax.WithholdingTax.getMoneyFromForm(pVal);
                //            B1.WithholdingTax.WithholdingTax.WTCalculation(ref pVal, out blOpenWindow);
                //        }
                //    }
                //}




                /*
                #region BPFiltering CFL
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.FormTypeEx == B1.BPImpairment.InteractionId.Default.frmBPFilteringFormType && !pVal.BeforeAction)
              {
                  SAPbouiCOM.IChooseFromListEvent oCFL = null;
                  oCFL = (SAPbouiCOM.IChooseFromListEvent)pVal;
                  if (pVal.ItemUID == "Item_1")
                  {
                      
                     B1.BPImpairment.BPFiltering.objEditBPF_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                  }

                  if (pVal.ItemUID == "Item_4")
                  {
                      
                      B1.BPImpairment.BPFiltering.objEditBPT_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                  }

                  if (pVal.ItemUID == "Item_24")
                  {
                      
                      B1.BPImpairment.BPFiltering.objEditBPG_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                  }
              }
                #endregion BPFiltering CFL

                #region BPImpairment CFL

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.FormTypeEx == B1.BPImpairment.InteractionId.Default.frmBPResultFormType && !pVal.BeforeAction)
                {
                  SAPbouiCOM.IChooseFromListEvent oCFL = null;
                  oCFL = (SAPbouiCOM.IChooseFromListEvent)pVal;
                  if (pVal.ItemUID == "Item_10")
                  {

                      B1.BPImpairment.BPImpairment.objEditBPAcc_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                  }

                  if (pVal.ItemUID == "Item_11")
                  {

                      B1.BPImpairment.BPImpairment.objEditBPDet_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                  }


                }
                #endregion BPImpairment CFL

                #region BPFIltering Query Button
                if(pVal.FormTypeEx == B1.BPImpairment.InteractionId.Default.frmBPFilteringFormType && !pVal.BeforeAction && pVal.ItemUID == B1.BPImpairment.InteractionId.Default.frmBPFilteringQueryButtonId)
                {
                    B1.BPImpairment.BPFiltering.objButton_ClickAfter(pVal.ItemUID, FormUID, pVal);
                }

                #endregion BPFIltering

                #region BPImpair GridFocus

                if(pVal.ItemUID == B1.BPImpairment.InteractionId.Default.frmBPImpairmentGridId && !pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.FormTypeEx == B1.BPImpairment.InteractionId.Default.frmBPResultFormType)
                {
                    if(pVal.ColUID == B1.BPImpairment.InteractionId.Default.frmBPImpairmentColumnIndex || pVal.ColUID == B1.BPImpairment.InteractionId.Default.frmBPImpairmentColumnDeterioro || pVal.ColUID == B1.BPImpairment.InteractionId.Default.frmBPImpairmentColumnPercentage)
                    {
                        B1.BPImpairment.BPImpairment.objGridColumn_LostFocusAfter(FormUID, pVal);
                    }
                }


                #endregion BPImpair GridFocus

                #region BPImpairment Calculate Button
                if (pVal.FormTypeEx == B1.BPImpairment.InteractionId.Default.frmBPResultFormType && !pVal.BeforeAction && pVal.ItemUID == B1.BPImpairment.InteractionId.Default.frmBPImpairmentButtonSumId)
                {
                    B1.BPImpairment.BPImpairment.objButton_ClickAfter(FormUID, pVal);
                }

                if (pVal.FormTypeEx == B1.BPImpairment.InteractionId.Default.frmBPResultFormType && !pVal.BeforeAction && pVal.ItemUID == B1.BPImpairment.InteractionId.Default.frmBPImpairmentButtonContabId)
                {
                    B1.BPImpairment.BPImpairment.objButton_PressedAfter(FormUID, pVal);
                }
                #endregion

                #region FSNotes Button
                if (pVal.FormTypeEx == B1.FSNotes.InteractionId.Default.fsNotesFormType && !pVal.BeforeAction && pVal.ItemUID == B1.FSNotes.InteractionId.Default.formButtonGetId)
                {
                    B1.FSNotes.FSNotes.oGetButton_PressedAfter(FormUID, pVal);
                }

                if (pVal.FormTypeEx == B1.FSNotes.InteractionId.Default.fsNotesFormType && !pVal.BeforeAction && pVal.ItemUID == B1.FSNotes.InteractionId.Default.formButtonUpdateId)
                {
                    B1.FSNotes.FSNotes.oUpdate_PressedAfter(FormUID, pVal);
                }

                #endregion FSNotes Button

                #region FSNotes Combo
                if(pVal.FormTypeEx == B1.FSNotes.InteractionId.Default.fsNotesFormType && !pVal.BeforeAction && pVal.ItemUID == B1.FSNotes.InteractionId.Default.formTypeCmbId && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    B1.FSNotes.FSNotes.oNoteCombo_ComboSelectAfter(FormUID, pVal);
                }

                if (pVal.FormTypeEx == B1.FSNotes.InteractionId.Default.fsNotesFormType && !pVal.BeforeAction && pVal.ItemUID == B1.FSNotes.InteractionId.Default.formTypeCmbNote && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    B1.FSNotes.FSNotes.oTypeCombo_ComboSelectAfter(FormUID, pVal);
                }
                #endregion FSNotes Combo

                #region ItemFiltering CFL
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.FormTypeEx == B1.InventoryImpairment.InteractionId.Default.frmItemFilteringFormType && !pVal.BeforeAction)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFL = null;
                    oCFL = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    if (pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmEditItemFromId)
                    {

                        B1.InventoryImpairment.ItemsFiltering.objEditITF_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                    }

                    if (pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmEditItemToId)
                    {

                        B1.InventoryImpairment.ItemsFiltering.objEditITT_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                    }

                    if (pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmEditItemGroupId)
                    {

                        B1.InventoryImpairment.ItemsFiltering.objEditITG_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                    }
                    if (pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmEditWarehouseFromId)
                    {

                        B1.InventoryImpairment.ItemsFiltering.objEditWF_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                    }
                    if (pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmEditWarehouseToId)
                    {

                        B1.InventoryImpairment.ItemsFiltering.objEditWT_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                    }
                }
                #endregion ItemFiltering CFL

                #region ItemImpairment CFL

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.FormTypeEx == B1.InventoryImpairment.InteractionId.Default.frmItemResultFormType && !pVal.BeforeAction)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFL = null;
                    oCFL = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    if (pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmCmbImpAccId)
                    {

                        B1.InventoryImpairment.ItemImpairment.objEditInvDet_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                    }

                    if (pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmCmbInvAcctId)
                    {

                        B1.InventoryImpairment.ItemImpairment.objEditInvAcc_ChooseFromListAfter(oCFL.SelectedObjects, FormUID);
                    }


                }
                #endregion ItemImpairment CFL

                #region ItemFIltering Query Button
                if (pVal.FormTypeEx == B1.InventoryImpairment.InteractionId.Default.frmItemFilteringFormType && !pVal.BeforeAction && pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmQueryButtonId)
                {
                    B1.InventoryImpairment.ItemsFiltering.objButton_ClickAfter(FormUID, pVal);
                }

                #endregion ItemFIltering Query Button

                #region ItemImpair GridFocus

                if (pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmGridResultId && !pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.FormTypeEx == B1.InventoryImpairment.InteractionId.Default.frmItemResultFormType)
                {
                    if (pVal.ColUID == B1.InventoryImpairment.InteractionId.Default.frmColumnDetId)
                    {
                        B1.InventoryImpairment.ItemImpairment.objGridColumn_LostFocusAfter(FormUID, pVal);
                    }
                }


                #endregion ItemImpair GridFocus

                #region ItemImpairment Calculate Button
                if (pVal.FormTypeEx == B1.InventoryImpairment.InteractionId.Default.frmItemResultFormType && !pVal.BeforeAction && pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmButtonSum)
                {
                    B1.InventoryImpairment.ItemImpairment.objButton_ClickAfter(FormUID, pVal);
                }

                if (pVal.FormTypeEx == B1.InventoryImpairment.InteractionId.Default.frmItemResultFormType && !pVal.BeforeAction && pVal.ItemUID == B1.InventoryImpairment.InteractionId.Default.frmButtonCont)
                {
                    B1.InventoryImpairment.ItemImpairment.objButtonConytab_ClickAfter(FormUID, pVal);
                }
                #endregion ItemImpairment Calculate Button

                #region FixedAssets LostFocus

                if (pVal.ItemUID == B1.FixedAssets.InteractionId.Default.fixedAssetsMasterFormMatrixId && !pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.FormTypeEx == B1.FixedAssets.InteractionId.Default.productionUnitsFormType)
                {
                    if (pVal.ColUID == B1.FixedAssets.InteractionId.Default.fixedAssetsPUColumnId)
                    {
                        B1.FixedAssets.ProductionUnits.objMatrix_LostFocusAfter(FormUID, pVal);
                    }
                }
                #endregion

                */
                /*  

                if (pVal.FormTypeEx == "420" && pVal.ItemUID == "1" && pVal.BeforeAction)
                {
                    oForm = BYBB1MainObject.Instance.B1Application.Forms.Item(FormUID);
                    SAPbouiCOM.ComboBox objCombo = oForm.Items.Item("26").Specific;
                    string strValue = objCombo.Selected.Description;
                    if (strValue.Contains("01"))
                    {
                        BubbleEvent = false;
                        string strPath = AppDomain.CurrentDomain.BaseDirectory + "\\CR\\01.rpt";
                        ReportDocument myDataReport = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                        myDataReport.Load(strPath);
                        Stream returnData = myDataReport.ExportToStream(ExportFormatType.PortableDocFormat);
                        FileStream fileStream = File.Create(AppDomain.CurrentDomain.BaseDirectory+"/temp.pdf", (int)returnData.Length);

                        byte[] bytesInStream = new byte[returnData.Length];
                        returnData.Read(bytesInStream, 0, bytesInStream.Length);
                         
                        fileStream.Write(bytesInStream, 0, bytesInStream.Length);
                        fileStream.Close();

                        System.Diagnostics.Process.Start(AppDomain.CurrentDomain.BaseDirectory + "/temp.pdf");

                        //myDataReport.Close();
                        //return returnData;


                    }

                }
                */




            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "BYBEventManager.objApplication_ItemEvent", er, 15, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "BYBEventManager.objApplication_ItemEvent", er, 15, System.Diagnostics.EventLogEntryType.Error);
            }
        } 

        void objApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            BYBMenuManager objMenuManager = new BYBMenuManager();
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:

                        Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                        Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:

                        
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                        
                        Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        Application.Exit();
                        break;
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 17, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "MenuManagerClass.removeMenu", er, 17, System.Diagnostics.EventLogEntryType.Error);
            }
        }


        
    }
}
