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

namespace T1.B1.ReletadParties
{
    public class Operations
    {
        private static Operations objReletadParties;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static bool runResizelogic = true;

        private Operations()
        {

        }

        public static void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            if (objReletadParties == null)
            {
                objReletadParties = new Operations();
            }

            BubbleEvent = true;
            try
            {

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    && !pVal.BeforeAction
                    && pVal.ActionSuccess
                    && pVal.FormTypeEx == Settings._Main.BPFormTypeEx
                    )
                {
                    Instance.BYBRelatedPartiesFolderAdd(pVal.FormUID);
                    runResizelogic = false;


                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.ActionSuccess
                    && pVal.FormTypeEx == "BYB_FTRA1"
                    && pVal.ItemUID == "btnAdd"
                    )
                {
                    Instance.createMissingRelatedParties(pVal);
                    


                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                    && !pVal.BeforeAction
                    && pVal.FormTypeEx == Settings._Main.BPFormTypeEx
                    )
                {

                    if (runResizelogic)
                    {
                        Instance.BYBRelatedPartiesFolderAdd(pVal.FormUID);
                    }
                    runResizelogic = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.FormTypeEx == Settings._Main.BPFormTypeEx
                    && pVal.ItemUID == Settings._Main.RelatedPartiesFolderId
                    )
                {

                    MainObject.Instance.B1Application.Forms.Item(pVal.FormUID).PaneLevel = Settings._Main.RelatedPartiesFolderPane;
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
        public static void formDataAddEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool blBubbleEvent)
        {
            if (objReletadParties == null)
            {
                objReletadParties = new Operations();
            }

            blBubbleEvent = true;

            try
            {

                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._Main.BPFormTypeEx &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    )
                {
                    Instance.BYBRelatedPartiesFolderAdd(BusinessObjectInfo.FormUID);
                    Instance.getRelatedpartyInfo(BusinessObjectInfo);


                }

                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._Main.BPFormTypeEx &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    )
                {
                    
                    Instance.addRelatedPartyInfo(BusinessObjectInfo);
                    Instance.cleanEditTexts(BusinessObjectInfo.FormUID);


                }
                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._Main.BPFormTypeEx &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    )
                {

                    Instance.updateRelatedPartyInfo(BusinessObjectInfo);
                    


                }

                if (
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.FormTypeEx == Settings._Main.BPFormTypeEx &&
                    !BusinessObjectInfo.BeforeAction &&
                    BusinessObjectInfo.ActionSuccess &&
                    BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE
                    )
                {

                    Instance.deleteRelatedPartyInfo(BusinessObjectInfo);
                    Instance.cleanEditTexts(BusinessObjectInfo.FormUID);


                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                _Logger.Error("FormDataEvent Error", comEx);

            }
            catch (Exception er)
            {
                _Logger.Error("FormDataEvent Error", er);

            }
        }
        public static void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            if(objReletadParties == null)
            {
                objReletadParties = new Operations();
            }

            BubbleEvent = true;
            try
            {
                if (pVal.MenuUID == "BYB_MRP05"
                    && !pVal.BeforeAction)
                {
                    Instance.loadRelatedPartiesUDOForm();
                }
                if (pVal.MenuUID == "1282"
                    && !pVal.BeforeAction)
                {
                    
                    Instance.addLineToDS();
                }
                if (pVal.MenuUID == "BYB_MRPAR"
                    && !pVal.BeforeAction)
                {
                    EventInfoClass eventInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                    Instance.relatedPartiedMatrixOperation(eventInfo, "Add");
                }
                if (pVal.MenuUID == "BYB_MRPDR"
                    && !pVal.BeforeAction)
                {

                    EventInfoClass eventInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                    Instance.relatedPartiedMatrixOperation(eventInfo, "Delete");
                }

                if (pVal.MenuUID == "BYB_MRP03"
                    && !pVal.BeforeAction)
                {

                    Instance.loadMissingRelatedPartiesForm();
                }

                if (pVal.MenuUID == "BYB_MRPARU"
                    && !pVal.BeforeAction)
                {
                    EventInfoClass eventInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                    Instance.relatedPartiedMatrixOperationUDO(eventInfo, "Add");
                }
                if (pVal.MenuUID == "BYB_MRPDRU"
                    && !pVal.BeforeAction)
                {

                    EventInfoClass eventInfo = CacheManager.CacheManager.Instance.getFromCache(Settings._Main.lastRightClickEventInfo);
                    Instance.relatedPartiedMatrixOperationUDO(eventInfo, "Delete");
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

        public static void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            if (objReletadParties == null)
            {
                objReletadParties = new Operations();
            }

            SAPbouiCOM.Form objForm = null;
            BubbleEvent = true;
            try
            {
                objForm = MainObject.Instance.B1Application.Forms.Item(eventInfo.FormUID);
                #region BP Form
                if (objForm.TypeEx == Settings._Main.BPFormTypeEx
                    && eventInfo.BeforeAction
                    && eventInfo.ItemUID == Settings._Main.BPFormMatrixId
                    
                    )
                {
                    Instance.addInsertRowRelationMenu(objForm,eventInfo);
                    Instance.addDeleteRowRelationMenu(objForm, eventInfo);


                }

                if (objForm.TypeEx == Settings._Main.BPFormTypeEx
                    && !eventInfo.BeforeAction
                    && eventInfo.ItemUID == Settings._Main.BPFormMatrixId

                    )
                {
                    Instance.removeDeleteRowRelationMenu();
                    Instance.removeInsertRowRelationMenu();


                }
                #endregion
                #region UDO Form
                if (objForm.TypeEx == "BYB_T1RPA100UDO"
                    && eventInfo.BeforeAction
                    && eventInfo.ItemUID == "Item_1"

                    )
                {
                    Instance.addInsertRowRelationMenuUDO(objForm, eventInfo);
                    Instance.addDeleteRowRelationMenuUDO(objForm, eventInfo);


                }

                if (objForm.TypeEx == "BYB_T1RPA100UDO"
                    && !eventInfo.BeforeAction
                    && eventInfo.ItemUID == "Item_1"

                    )
                {
                    Instance.removeDeleteRowRelationMenuUDO();
                    Instance.removeInsertRowRelationMenuUDO();


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
