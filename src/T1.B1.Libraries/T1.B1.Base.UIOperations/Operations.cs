using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using log4net;
using System.Runtime.InteropServices;

namespace T1.B1.Base.UIOperations
{
    public class Operations
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        //private static Operations objOper = null;

        private Operations()
        {

        }

        public static void setStatusBarMessage(string strMessage, bool isError, SAPbouiCOM.BoMessageTime msgTime)
        {
            try
            {
                MainObject.Instance.B1Application.SetStatusBarMessage(strMessage, msgTime, isError);
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }
        public static void startProgressBar(string Message, int Max)
        {
            try
            {
                SAPbouiCOM.ProgressBar objProgressbar = null;
                objProgressbar = T1.CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
                if (objProgressbar == null)
                {
                    objProgressbar = MainObject.Instance.B1Application.StatusBar.CreateProgressBar(Message, Max, false);
                    objProgressbar.Value = 1;
                    objProgressbar.Text = Message;
                    objProgressbar.Value = 1;
                    CacheManager.CacheManager.Instance.addToCache(T1.CacheManager.Settings._Main.progressBarCacheName, objProgressbar, CacheManager.CacheManager.objCachePriority.Default);
                }
            }
            catch (Exception er)
            {
                CacheManager.CacheManager.Instance.removeFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
                _Logger.Error("", er);
            }
        }

        public static void stopProgressBar()
        {
            try
            {
                SAPbouiCOM.ProgressBar objProgressbar = null;
                objProgressbar = CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
                if (objProgressbar != null)
                {
                    objProgressbar.Stop();

                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                CacheManager.CacheManager.Instance.removeFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
            }
        }

        public static void setProgressBarMessage(string strMessage, int Value)
        {
            try
            {
                SAPbouiCOM.ProgressBar objProgressbar = null;
                objProgressbar = CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.progressBarCacheName);
                if (objProgressbar != null)
                {
                    objProgressbar.Text = strMessage;
                    objProgressbar.Value = Value;

                }
                else
                {
                    startProgressBar(strMessage, Value);
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static void toggleSelectCheckBox(SAPbouiCOM.ItemEvent pVal, string dtName, string CellNumber)
        {
            SAPbouiCOM.DataTable variable = null;
            SAPbouiCOM.Form variable1 = null;
            XmlDocument xmlDocument = null;
            XmlNodeList xmlNodeLists = null;
            try
            {
                xmlDocument = new XmlDocument();
                variable1 = MainObject.Instance.B1Application.Forms.Item(pVal.FormUID);
                variable = variable1.DataSources.DataTables.Item(dtName);
                xmlDocument.LoadXml(variable.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));
                xmlNodeLists = xmlDocument.SelectNodes("/DataTable/Rows/Row/Cells/Cell["+ CellNumber +"]/Value");
                if (xmlNodeLists.Count > 0)
                {
                    foreach (XmlNode xmlNodes in xmlNodeLists)
                    {
                        if (xmlNodes.InnerText != "Y")
                        {
                            xmlNodes.InnerText = "Y";
                        }
                        else
                        {
                            xmlNodes.InnerText = "N";
                        }
                    }
                }
                variable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly, xmlDocument.InnerXml);
            }
            catch (COMException cOMException1)
            {
                COMException cOMException = cOMException1;
                Exception exception = new Exception(Convert.ToString(string.Concat(new object[] { "COM Error::", cOMException.ErrorCode, "::", cOMException.Message, "::", cOMException.StackTrace })));
                _Logger.Error("", exception);
            }
            catch (Exception exception2)
            {
                Exception exception1 = exception2;
                _Logger.Error("", exception2);
            }
        }

        public static SAPbouiCOM.Form openFormfromXML(string strXML, string FormType, bool Modal)
        {
            SAPbouiCOM.Form objForm = null;
            //int intLeft = 0;
            string strModFile = "";
            XmlDocument xmlResult = new XmlDocument();
            string strGUID = "";

            try
            {
                 strGUID = Modal ? FormType : Guid.NewGuid().ToString().Substring(1, 10);
                strModFile = strXML.Replace("[--UniqueId--]", strGUID)
                    .Replace("[--FormType--]", FormType);
                MainObject.Instance.B1Application.LoadBatchActions(ref strModFile);
                string strResult = MainObject.Instance.B1Application.GetLastBatchResults();
                        xmlResult = new XmlDocument();
                        xmlResult.LoadXml(strResult);
                        bool errors = xmlResult.SelectSingleNode(Settings._FormLoad.errorPath).HasChildNodes != true ? false : true;
                        if (!errors)
                        {
                    objForm = MainObject.Instance.B1Application.Forms.Item(strGUID);
                            
                        }
                        



            }
            catch (COMException ex)
            {
                _Logger.Error("", ex);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            return objForm;
        }


        
        //public static void cacheDimensionsComboBox()
        //{
            

        //    bool isMultiDim = false;
        //    List<T1.B1.InternalClasses.dimensionInfo> objDimList = null;
        //    List<T1.B1.InternalClasses.costCenterInfo> objProfitCenterList = null;
        //    List<T1.B1.InternalClasses.projectInfo> objProjectList = null;
        //    SAPbouiCOM.ComboBox objCombo = null;
        //    int intNumberOfRows = -1;

        //    try
        //    {
        //        objForm = MainObject.Instance.B1Application.Forms.Item(FormUID);
        //        objForm.Freeze(true);


        //        if (objForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || objForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //        {
        //            isMultiDim = isMultipleDimension();
        //            if (isMultiDim)
        //            {
        //                objDimList = getDimensionsList();

        //            }


        //            objProfitCenterList = getProfitCenterList();

        //            objProjectList = getProjectList();
        //            objMatrix = objForm.Items.Item("0_U_G").Specific;
        //            intNumberOfRows = objMatrix.RowCount;


        //            #region load DropDowns Dimensions and ProfitCenter


        //            if (!isMultiDim)
        //            {
        //                #region ProfitCenter
        //                if (intNumberOfRows > 0)
        //                {
        //                    objColumn = objMatrix.Columns.Item("C_0_10");
        //                    objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
        //                    for (int i = 0; i < objCombo.ValidValues.Count; i++)
        //                    {
        //                        objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                    }

        //                    foreach (costCenterInfo oPC in objProfitCenterList)
        //                    {
        //                        objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
        //                    }
        //                    updateVisualizarionString(objForm.TypeEx, "C_0_10", "0_U_G", "Centro de Costo");
        //                    objColumn = objMatrix.Columns.Item("C_0_11");
        //                    objColumn.Visible = false;
        //                    objColumn = objMatrix.Columns.Item("C_0_12");
        //                    objColumn.Visible = false;
        //                    objColumn = objMatrix.Columns.Item("C_0_13");
        //                    objColumn.Visible = false;
        //                    objColumn = objMatrix.Columns.Item("Col_14");
        //                    objColumn.Visible = false;

        //                }
        //                #endregion
        //            }
        //            else
        //            {
        //                #region Dimensions

        //                foreach (dimensionInfo oDIM in objDimList)
        //                {
        //                    switch (oDIM.DimentionCode)
        //                    {
        //                        case 1:
        //                            if (oDIM.isActive)
        //                            {
        //                                objColumn = objMatrix.Columns.Item("C_0_10");
        //                                objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
        //                                for (int i = 0; i < objCombo.ValidValues.Count; i++)
        //                                {
        //                                    objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                                }

        //                                foreach (costCenterInfo oPC in objProfitCenterList)
        //                                {
        //                                    if (oPC.DimensionCode == 1)
        //                                    {
        //                                        objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
        //                                    }
        //                                }
        //                                updateVisualizarionString(objForm.TypeEx, "C_0_10", "0_U_G", oDIM.DimensionName);

        //                            }

        //                            break;
        //                        case 2:
        //                            if (oDIM.isActive)
        //                            {
        //                                objColumn = objMatrix.Columns.Item("C_0_11");
        //                                objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
        //                                for (int i = 0; i < objCombo.ValidValues.Count; i++)
        //                                {
        //                                    objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                                }

        //                                foreach (costCenterInfo oPC in objProfitCenterList)
        //                                {
        //                                    if (oPC.DimensionCode == 2)
        //                                    {
        //                                        objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
        //                                    }
        //                                }
        //                                updateVisualizarionString(objForm.TypeEx, "C_0_11", "0_U_G", oDIM.DimensionName);

        //                            }
        //                            else
        //                            {
        //                                objColumn = objMatrix.Columns.Item("C_0_11");
        //                                objColumn.Visible = false;
        //                            }
        //                            break;
        //                        case 3:
        //                            if (oDIM.isActive)
        //                            {
        //                                objColumn = objMatrix.Columns.Item("C_0_12");
        //                                objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
        //                                for (int i = 0; i < objCombo.ValidValues.Count; i++)
        //                                {
        //                                    objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                                }

        //                                foreach (costCenterInfo oPC in objProfitCenterList)
        //                                {
        //                                    if (oPC.DimensionCode == 3)
        //                                    {
        //                                        objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
        //                                    }
        //                                }
        //                                updateVisualizarionString(objForm.TypeEx, "C_0_12", "0_U_G", oDIM.DimensionName);

        //                            }
        //                            else
        //                            {
        //                                objColumn = objMatrix.Columns.Item("C_0_12");
        //                                objColumn.Visible = false;
        //                            }
        //                            break;
        //                        case 4:
        //                            if (oDIM.isActive)
        //                            {
        //                                objColumn = objMatrix.Columns.Item("C_0_13");
        //                                objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
        //                                for (int i = 0; i < objCombo.ValidValues.Count; i++)
        //                                {
        //                                    objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                                }

        //                                foreach (costCenterInfo oPC in objProfitCenterList)
        //                                {
        //                                    if (oPC.DimensionCode == 4)
        //                                    {
        //                                        objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
        //                                    }
        //                                }
        //                                updateVisualizarionString(objForm.TypeEx, "C_0_13", "0_U_G", oDIM.DimensionName);

        //                            }
        //                            else
        //                            {
        //                                objColumn = objMatrix.Columns.Item("C_0_13");
        //                                objColumn.Visible = false;
        //                            }
        //                            break;
        //                        case 5:
        //                            if (oDIM.isActive)
        //                            {
        //                                objColumn = objMatrix.Columns.Item("C_0_14");
        //                                objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
        //                                for (int i = 0; i < objCombo.ValidValues.Count; i++)
        //                                {
        //                                    objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                                }

        //                                foreach (costCenterInfo oPC in objProfitCenterList)
        //                                {
        //                                    if (oPC.DimensionCode == 5)
        //                                    {
        //                                        objCombo.ValidValues.Add(oPC.CostCenterCode, oPC.CostCenterName);
        //                                    }
        //                                }
        //                                updateVisualizarionString(objForm.TypeEx, "C_0_14", "0_U_G", oDIM.DimensionName);

        //                            }
        //                            else
        //                            {
        //                                objColumn = objMatrix.Columns.Item("C_0_14");
        //                                objColumn.Visible = false;
        //                            }
        //                            break;
        //                    }
        //                }
        //                #endregion


        //            }
        //            #endregion


        //            #region Projects

        //            if (intNumberOfRows > 0)
        //            {
        //                objColumn = objMatrix.Columns.Item("C_0_9");
        //                objCombo = objColumn.Cells.Item(intNumberOfRows).Specific;
        //                for (int i = 0; i < objCombo.ValidValues.Count; i++)
        //                {
        //                    objCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                }

        //                foreach (projectInfo oPC in objProjectList)
        //                {
        //                    objCombo.ValidValues.Add(oPC.ProjectCode, oPC.ProjectName);

        //                }
        //            }

        //            #endregion

        //        }

        //    }
        //    catch (Exception er)
        //    {
        //        _Logger.Error("", er);
        //    }
        //    finally
        //    {
        //        objForm.Refresh();
        //        objForm.Freeze(false);
        //    }
        //}
    }
}
