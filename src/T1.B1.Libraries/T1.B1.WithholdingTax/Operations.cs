using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using Newtonsoft.Json;
using System.Collections;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;

namespace T1.B1.WithholdingTax
{
    public class Operations
    {
        private static Operations objWithHoldingTax;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static bool runResizelogic = true;
        private static List<string> WHPurchaseDocuments = new List<string>();
        private static List<string> WHSalesDocuments = new List<string>();
        private Operations()
        {
            WHPurchaseDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTPurchaseObjects);
            WHSalesDocuments = JsonConvert.DeserializeObject<List<string>>(Settings._WithHoldingTax.WTSalesObjects);
        }
        public static void formDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool blBubbleEvent)
        {
            if(objWithHoldingTax == null)
            {
                objWithHoldingTax = new Operations();
            }

            blBubbleEvent = true;
            try
            {
                #region Autoretenciones
                #region Invoice
                if (BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == "133" &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    SelfWithholdingTax.addSelfWithHoldingTax(BusinessObjectInfo);
                }

                if (BusinessObjectInfo.ActionSuccess &&
                    (BusinessObjectInfo.FormTypeEx == "133" || BusinessObjectInfo.FormTypeEx == "179") && 
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
                {
                    SelfWithholdingTax.getSWTaxInfoForDocument(BusinessObjectInfo);
                }
                #endregion

                #region CreditNote
                if (BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == "179" &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    SelfWithholdingTax.addSelfWithHoldingTax(BusinessObjectInfo);
                }
                #endregion


                #endregion

                #region Retenciones

                if (BusinessObjectInfo.ActionSuccess
                    && (WHPurchaseDocuments.Contains(BusinessObjectInfo.FormTypeEx) || WHSalesDocuments.Contains(BusinessObjectInfo.FormTypeEx))
                    && !BusinessObjectInfo.BeforeAction
                    && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    AddDocumentInfoArgs objArgs = new AddDocumentInfoArgs();
                    objArgs.ObjectKey = BusinessObjectInfo.ObjectKey;
                    objArgs.ObjectType = BusinessObjectInfo.Type;
                    objArgs.FormtTypeEx = BusinessObjectInfo.FormTypeEx;
                    objArgs.FormUID = BusinessObjectInfo.FormUID;

                    WithholdingTax.addDocumentInfo(objArgs);
                }
                
                #endregion

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
                #region Autoretention

                if (pVal.MenuUID == "BYB_MWT04"
                    && !pVal.BeforeAction)
                {
                    SelfWithholdingTax.loadCancelSWTaxForm();
                }
                if (pVal.MenuUID == "BYB_MWT03"
                    && !pVal.BeforeAction)
                {
                    SelfWithholdingTax.loadMissingSWTaxForm();
                }

                //COnfiguración Autoretencion
                if (pVal.MenuUID == "BYB_MWT02"
                    && !pVal.BeforeAction)
                {
                    SelfWithholdingTax.loadSWTaxConfigForm();
                }

                //Add Row BP
                if (pVal.MenuUID == "BYB_MWTRU"
                    && !pVal.BeforeAction)
                {
                    EventInfoClass eventInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                    SelfWithholdingTax.relatedPartiedMatrixOperationUDO(eventInfo, "Add");
                }
                //Remove Row BP
                if (pVal.MenuUID == "BYB_MWTDRU"
                    && !pVal.BeforeAction)
                {

                    EventInfoClass eventInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                    SelfWithholdingTax.relatedPartiedMatrixOperationUDO(eventInfo, "Delete");
                }
                #endregion

                #region Retenciones
                if(pVal.MenuUID == "5897"
                    && pVal.BeforeAction)
                {
                    string strLastActiveForm = MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID;
                    CacheManager.CacheManager.Instance.addToCache("WTLastActiveForm", strLastActiveForm, CacheManager.CacheManager.objCachePriority.Default);
                    
                }
                if (pVal.MenuUID == "6005"
                    && pVal.BeforeAction)
                {
                    string strLastActiveForm = MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID;
                    CacheManager.CacheManager.Instance.addToCache("LastActiveForm", strLastActiveForm, CacheManager.CacheManager.objCachePriority.Default);

                }

                if (pVal.MenuUID == "BYB_MWT06"
                    && pVal.BeforeAction)
                {
                    MainObject.Instance.B1Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "BYB_T1WHT200", "");

                }

                //Transacciones faltantes
                if (pVal.MenuUID == "BYB_MWT07"
                    && !pVal.BeforeAction)
                {

                    WithholdingTax.loadMissingOperationsForm();
                }




                #endregion
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            if(objWithHoldingTax == null)
            {
                objWithHoldingTax = new Operations();
            }
            string[] showInFolderList;
            bool blInList = false;
            BubbleEvent = true;
            try
            {
                #region WithHolding Tax

                #region Purchase


                if (WHPurchaseDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && (pVal.ItemUID == "4" || pVal.ItemUID == "54")
                    && !pVal.BeforeAction

                    )
                {
                    if (WithholdingTax.formModeAdd(pVal))
                        {
                        WithholdingTax.getSelectedBPInformation(pVal, true);
                        CacheManager.CacheManager.Instance.addToCache("WTCFLExecuted", true, CacheManager.CacheManager.objCachePriority.Default);
                    }

                }

                if (WHPurchaseDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD
                    && !pVal.BeforeAction

                    )
                {
                    CacheManager.CacheManager.Instance.removeFromCache("Disable_" + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache("WTLogicDone_" + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache("WTCFLExecuted");
                    CacheManager.CacheManager.Instance.removeFromCache("WTLastActiveForm");

                }

                if (WHPurchaseDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    && !pVal.BeforeAction

                    )
                {
                    if (WithholdingTax.formModeAdd(pVal))
                        {
                        bool WTExec = CacheManager.CacheManager.Instance.getFromCache("WTCFLExecuted") == null ? false : true;
                        bool LogicDone = CacheManager.CacheManager.Instance.getFromCache("WTLogicDone_" + pVal.FormUID) == null ? false : true;
                        if (pVal.ItemUID == "4" || pVal.ItemUID == "54")
                        {
                            LogicDone = false;
                        }

                        if (!LogicDone)
                        {
                            if (!WTExec)
                            {
                                WithholdingTax.getSelectedBPInformation(pVal, false);

                            }
                            else
                            {
                                CacheManager.CacheManager.Instance.removeFromCache("WTCFLExecuted");
                            }

                            WithholdingTax.activateWTMenu(pVal.FormUID);
                        }
                    }

                }
                // BillTo Combobox
                if (WHPurchaseDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    && pVal.ItemUID == "226"
                    && !pVal.BeforeAction

                    )
                {
                    if (WithholdingTax.formModeAdd(pVal))
                        {
                        WithholdingTax.getSelectedBPInformation(pVal, false);
                        WithholdingTax.activateWTMenu(pVal.FormUID);
                    }
                    
                }

                //LinkTo WT Table arrow
                if (WHPurchaseDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && pVal.ItemUID == "173"
                    && pVal.BeforeAction

                    )
                {
                    if (WithholdingTax.formModeAdd(pVal))
                        {
                        string strLastActiveForm = MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID;
                        CacheManager.CacheManager.Instance.addToCache("WTLastActiveForm", strLastActiveForm, CacheManager.CacheManager.objCachePriority.Default);
                    }
                }



                #endregion

                #region Sales


                if (WHSalesDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && (pVal.ItemUID == "4" || pVal.ItemUID == "54")
                    && !pVal.BeforeAction

                    )
                {
                    if (WithholdingTax.formModeAdd(pVal))
                        {
                        WithholdingTax.getSelectedBPInformation(pVal, true);
                        CacheManager.CacheManager.Instance.addToCache("WTCFLExecuted", true, CacheManager.CacheManager.objCachePriority.Default);
                    }

                }

                if (WHSalesDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD
                    && !pVal.BeforeAction

                    )
                {
                    CacheManager.CacheManager.Instance.removeFromCache("Disable_" + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTInfoGenCachePrefix + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTFOrmInfoCachePrefix + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache(Settings._WithHoldingTax.WTLastCardCodeCachePrefix + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache("WTLogicDone_" + pVal.FormUID);
                    CacheManager.CacheManager.Instance.removeFromCache("WTCFLExecuted");
                    CacheManager.CacheManager.Instance.removeFromCache("WTLastActiveForm");

                }

                if (WHSalesDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    && !pVal.BeforeAction

                    )
                {
                    if (WithholdingTax.formModeAdd(pVal))
                        {
                        bool WTExec = CacheManager.CacheManager.Instance.getFromCache("WTCFLExecuted") == null ? false : true;
                        bool LogicDone = CacheManager.CacheManager.Instance.getFromCache("WTLogicDone_" + pVal.FormUID) == null ? false : true;
                        if (pVal.ItemUID == "4" || pVal.ItemUID == "54")
                        {
                            LogicDone = false;
                        }

                        if (!LogicDone)
                        {
                            if (!WTExec)
                            {
                                WithholdingTax.getSelectedBPInformation(pVal, false);

                            }
                            else
                            {
                                CacheManager.CacheManager.Instance.removeFromCache("WTCFLExecuted");
                            }

                            WithholdingTax.activateWTMenu(pVal.FormUID);
                        }
                    }

                }
                // BillTo Combobox
                if (WHSalesDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    && pVal.ItemUID == "226"
                    && !pVal.BeforeAction

                    )
                {
                    if (WithholdingTax.formModeAdd(pVal))
                        {
                        WithholdingTax.getSelectedBPInformation(pVal, false);
                        WithholdingTax.activateWTMenu(pVal.FormUID);
                    }

                }

                //LinkTo WT Table arrow
                if (WHSalesDocuments.Contains(pVal.FormTypeEx)
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && pVal.ItemUID == "173"
                    && pVal.BeforeAction

                    )
                {
                    if (WithholdingTax.formModeAdd(pVal))
                        {
                        string strLastActiveForm = MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID;
                        CacheManager.CacheManager.Instance.addToCache("WTLastActiveForm", strLastActiveForm, CacheManager.CacheManager.objCachePriority.Default);
                    }
                }



                #endregion

                #region Missing Operations
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                   && !pVal.BeforeAction
                   && pVal.ActionSuccess
                   && pVal.FormTypeEx == "BYB_FMWT01"
                   && pVal.ItemUID == "btnAdd"
                   )
                {
                    WithholdingTax.createMissingOperations(pVal);



                }
                #endregion

                //if (pVal.FormTypeEx == "133"
                //    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                //    && pVal.ItemUID == "10002122"
                //    && !pVal.BeforeAction)
                //{
                //    bool isDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + pVal.FormUID) == null ? false : true;
                //    if (!isDisabled)
                //    {
                //        WithholdingTax.getWTforBP(pVal,false);
                //    }
                //}

                #region WithHolding Tax Form

                if (pVal.FormTypeEx == "60504"
                    && !pVal.BeforeAction
                        && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    string strLastActiveForm = CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm") == null ? "" : CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm");
                    if (strLastActiveForm.Trim().Length > 0)
                    {
                        bool blDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + strLastActiveForm) != null ? true : false;
                        if (!blDisabled)
                        {
                            string strFormAutoActivate = CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") != null ? CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") : "";
                            if (strFormAutoActivate.Trim() == strLastActiveForm.Trim())
                            {
                                T1.B1.Base.UIOperations.Operations.startProgressBar("Asignando retenciones automáticas...", 2);
                                WithholdingTax.setBPWT(strFormAutoActivate, pVal);
                                //CacheManager.CacheManager.Instance.removeFromCache("WTLastActiveForm");
                                //T1.B1.Base.UIOperations.Operations.stopProgressBar();
                            }
                        }
                    }

                    
                }
                if (pVal.FormTypeEx == "60504"
                    && pVal.BeforeAction
                    && pVal.EventType != SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && pVal.ItemUID == "1")
                {
                    string blAutoActivate = CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") != null ? CacheManager.CacheManager.Instance.getFromCache("WTAutoActivate") : "";
                    SAPbouiCOM.Form objForm = null;
                    string strLastActiveForm = CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm") == null ? "" : CacheManager.CacheManager.Instance.getFromCache("WTLastActiveForm");
                    if (strLastActiveForm.Trim().Length > 0)
                    {
                        bool isDisabled = CacheManager.CacheManager.Instance.getFromCache("Disable_" + strLastActiveForm) == null ? false : true;
                        if (blAutoActivate.Trim().Length == 0)
                        {
                            objForm = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                            if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && !isDisabled)
                            {
                                if (MainObject.Instance.B1Application.MessageBox("La modificación manual de las retenciones deshabilitará el cálculo automatico para este documento. Desea Continuar? ", 2, "Sí", "No", "") != 2)
                                {

                                    if (strLastActiveForm.Trim().Length > 0)
                                    {
                                        CacheManager.CacheManager.Instance.addToCache(string.Concat("Disable_", strLastActiveForm), true, CacheManager.CacheManager.objCachePriority.Default);
                                    }
                                    else
                                    {
                                        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                        {
                                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        }
                                        else
                                        {
                                            BubbleEvent = false;
                                        }
                                    }

                                    objForm.Close();
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    objForm.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);


                                }
                            }
                            CacheManager.CacheManager.Instance.removeFromCache("WTLastActiveForm");
                            T1.B1.Base.UIOperations.Operations.stopProgressBar();

                        }
                    }
                }





                #endregion

                #region Autoretenciones

                #region Cancel Wizard
                if (pVal.FormTypeEx == Settings._SelfWithHoldingTax.CancelFormUID
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "btnGet"
                    )
                {
                    SelfWithholdingTax.getPostedSWTaxDocuments(FormUID, pVal);
                }


                if (pVal.FormTypeEx == Settings._SelfWithHoldingTax.CancelFormUID
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "txtSWTCode"
                    )
                {
                    SelfWithholdingTax.setSelectedCode(pVal);
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
                     SelfWithholdingTax.cancelPostedTaxDocuments(FormUID, pVal);
                }
                #endregion

                #region SelfWithholdingTax Config Form

                if(pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    &&pVal.BeforeAction
                    && pVal.FormTypeEx == "BYB_T1SWT100UDO"
                    && pVal.ItemUID == "13_U_E")
                {
                    SelfWithholdingTax.filterAccounts(pVal,"CFL_DB");
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.FormTypeEx == "BYB_T1SWT100UDO"
                    && pVal.ItemUID == "13_U_E")
                {
                    SelfWithholdingTax.clearfilterAccounts(pVal, "CFL_DB");
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.FormTypeEx == "BYB_T1SWT100UDO"
                    && pVal.ItemUID == "14_U_E")
                {
                    SelfWithholdingTax.filterAccounts(pVal, "CFL_CR");
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.FormTypeEx == "BYB_T1SWT100UDO"
                    && pVal.ItemUID == "14_U_E")
                {
                    SelfWithholdingTax.clearfilterAccounts(pVal, "CFL_CR");
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.FormTypeEx == "BYB_T1SWT100UDO"
                    && pVal.ItemUID == "0_U_G")
                {
                    SelfWithholdingTax.setBPNameColumn(pVal);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && pVal.BeforeAction
                    && pVal.FormTypeEx == "BYB_T1SWT100UDO"
                    && pVal.ItemUID == "0_U_G")
                {
                    SelfWithholdingTax.filterBPs(pVal);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    && !pVal.BeforeAction
                    && pVal.FormTypeEx == "BYB_T1SWT100UDO"
                    && pVal.ItemUID == "0_U_G")
                {
                    SelfWithholdingTax.clearfilterBPs(pVal);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.FormTypeEx == "BYB_T1SWT100UDO"
                    && pVal.ItemUID == "btnAddAll")
                {
                    SelfWithholdingTax.addAllPBS(pVal);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.FormTypeEx == "BYB_T1SWT100UDO"
                    && pVal.ItemUID == "btnClear")
                {
                    SelfWithholdingTax.clearAllPBS(pVal);
                }

                #endregion

                #region SelfWithholdingTax Folder in Documents

                

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    && !pVal.BeforeAction
                    && pVal.ActionSuccess
                    
                    )
                {
                    showInFolderList = Settings._SelfWithHoldingTax.showFolderInDocumentsList.Split(',');
                    for(int i=0; i < showInFolderList.Length; i++)
                    {
                        if(showInFolderList[i] == pVal.FormTypeEx)
                        {
                            blInList = true;
                            break;
                        }
                    }
                    if (blInList)
                    {

                        SelfWithholdingTax.BYBSelfWithHoldingFolderAdd(pVal.FormUID);
                        runResizelogic = false;
                    }


                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                    && !pVal.BeforeAction
                    
                    )
                {

                    if (runResizelogic)
                    {
                        showInFolderList = Settings._SelfWithHoldingTax.showFolderInDocumentsList.Split(',');
                        for (int i = 0; i < showInFolderList.Length; i++)
                        {
                            if (showInFolderList[i] == pVal.FormTypeEx)
                            {
                                blInList = true;
                                break;
                            }
                        }
                        if (blInList)
                        {
                            SelfWithholdingTax.BYBSelfWithHoldingFolderAdd(pVal.FormUID);
                            blInList = false;
                        }
                        
                    }
                    runResizelogic = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.ItemUID == Settings._SelfWithHoldingTax.SelfWithHoldingFolderId
                    )
                {
                    showInFolderList = Settings._SelfWithHoldingTax.showFolderInDocumentsList.Split(',');
                    for (int i = 0; i < showInFolderList.Length; i++)
                    {
                        if (showInFolderList[i] == pVal.FormTypeEx)
                        {
                            blInList = true;
                            break;
                        }
                    }
                    if (blInList)
                    {
                        MainObject.Instance.B1Application.Forms.Item(pVal.FormUID).PaneLevel = Settings._SelfWithHoldingTax.SelfWithHoldingFolderPane;
                    }
                }
                #endregion
                

                #region AddWizard
                if (pVal.FormTypeEx == Settings._SelfWithHoldingTax.MissingSWTFormUID
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "btnGet"
                    )
                {
                    SelfWithholdingTax.getMissingSWTaxDocuments(FormUID, pVal);
                }


                

                if (!pVal.BeforeAction
                    && pVal.EventType != SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    && pVal.ItemUID == "grdSWT"
                    && pVal.FormTypeEx == Settings._SelfWithHoldingTax.MissingSWTFormUID)
                {
                    T1.B1.Base.UIOperations.Operations.toggleSelectCheckBox(pVal, "dtSelfWT", "1");
                }

                if (!pVal.BeforeAction
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    && pVal.ItemUID == "btnCalc"
                    && pVal.FormTypeEx == Settings._SelfWithHoldingTax.MissingSWTFormUID)
                {
                    SelfWithholdingTax.addMisingSWTDocuments(FormUID, pVal);
                }

                #endregion
                #endregion
                #endregion


            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                T1.B1.Base.UIOperations.Operations.stopProgressBar();
            }
        }

        public static void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            if (objWithHoldingTax == null)
            {
                objWithHoldingTax = new Operations();
            }

            SAPbouiCOM.Form objForm = null;
            BubbleEvent = true;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
               
                #region UDO Form
                if (objForm.TypeEx == "BYB_T1SWT100UDO"
                    && eventInfo.BeforeAction
                    && eventInfo.ItemUID == "0_U_G"

                    )
                {
                    SelfWithholdingTax.addInsertRowRelationMenuUDO(objForm, eventInfo);
                    SelfWithholdingTax.addDeleteRowRelationMenuUDO(objForm, eventInfo);
                    


                }

                if (objForm.TypeEx == "BYB_T1SWT100UDO"
                    && !eventInfo.BeforeAction
                    && eventInfo.ItemUID == "0_U_G"

                    )
                {
                    SelfWithholdingTax.removeDeleteRowRelationMenuUDO();
                    SelfWithholdingTax.removeInsertRowRelationMenuUDO();


                }
                #endregion
                #region Invoice
                if(objForm.TypeEx == "133"
                    && eventInfo.BeforeAction
                    )
                {
                    string strLastActiveForm = MainObject.Instance.B1Application.Forms.ActiveForm.UniqueID;
                    CacheManager.CacheManager.Instance.addToCache("LastActiveForm", strLastActiveForm, CacheManager.CacheManager.objCachePriority.Default);

                }

                if(objForm.TypeEx == "133"
                    && !eventInfo.BeforeAction)
                {
                    CacheManager.CacheManager.Instance.removeFromCache("LastActiveForm");

                }
                #endregion
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
