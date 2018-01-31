using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;

namespace T1.B1.EventManager
{
    public class Operations
    {
        private SAPbouiCOM.Application objApplication = null;
        public bool objStatus = false;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);


        public bool Status
        {
            get
            {
                return objStatus;
            }
        }

        public Operations()
        {
            try
            {
                objApplication = T1.B1.MainObject.Instance.B1Application;
                objApplication.AppEvent += objApplication_AppEvent;

                objApplication.EventLevel = SAPbouiCOM.BoEventLevelType.elf_Both;

                objApplication.FormDataEvent += T1.B1.WithholdingTax.Operations.formDataEvent;
                objApplication.MenuEvent += T1.B1.WithholdingTax.Operations.MenuEvent;
                objApplication.ItemEvent += T1.B1.WithholdingTax.Operations.ItemEvent;
                objApplication.RightClickEvent += T1.B1.WithholdingTax.Operations.RightClickEvent;

                objApplication.MenuEvent += T1.B1.ReletadParties.Operations.MenuEvent;
                objApplication.FormDataEvent += T1.B1.ReletadParties.Operations.formDataAddEvent;
                objApplication.ItemEvent += T1.B1.ReletadParties.Operations.ItemEvent;
                objApplication.RightClickEvent += T1.B1.ReletadParties.Operations.RightClickEvent;



                objApplication.MenuEvent += T1.B1.Expenses.Operations.MenuEvent;
                objApplication.RightClickEvent += T1.B1.Expenses.Operations.RightClickEvent;
                
                objApplication.ItemEvent += T1.B1.Expenses.Operations.ItemEvent;
                objApplication.FormDataEvent += T1.B1.Expenses.Operations.formDataAddEvent;

                objApplication.MenuEvent += T1.B1.InformesTerceros.Operations.MenuEvent;


            


                //objApplication.FormDataEvent += ObjApplication_FormDataEvent;


                ///RightClick events per Module
                //objApplication.RightClickEvent += T1.B1.WithholdingTax.WithholdingTax.RightClickEvent;
                //objApplication.RightClickEvent += T1.B1.Expenses.Expenses.RightClickEvent;

                ///Menu Events per Module
                //objApplication.MenuEvent += T1.B1.Expenses.Expenses.MenuEvent;
                //objApplication.MenuEvent += T1.B1.RelatedParties.RelatedParties.MenuEvent;
                //objApplication.MenuEvent += T1.B1.WithholdingTax.WithholdingTax.MenuEvent;


                //ItemEvents per Module
                //objApplication.ItemEvent += T1.B1.WithholdingTax.WithholdingTax.ItemEvent;
                //objApplication.ItemEvent += T1.B1.Expenses.Expenses.ItemEvent;



                ///UDO Event per Module Can I Inercept the event to change the XML?
                //objApplication.UDOEvent += T1.B1.WithholdingTax.WithholdingTax.UDOEvent;


                ///Move this to each class and then to dll for easy update without versioning

                //objApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(objApplication_ItemEvent);

                //objApplication.FormDataEvent += ObjApplication_FormDataEvent;
                //objApplication.FormDataEvent += B1.WithholdingTax.InternalRegistry.InternalRegistry.ObjApplication_FormDataEvent;
                //objApplication.FormDataEvent += T1.B1.Expenses.Expenses.LoadDataEvent;





                objStatus = true;


                T1.B1.MainObject.Instance.B1Application = objApplication;


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

        private void ObjApplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            throw new NotImplementedException();
        }

        private void ObjApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //SAPbobsCOM.Documents objInvoice = null;
            XmlDocument objMatrix = new XmlDocument();
            SAPbouiCOM.Form objForm = null;
            Dictionary<string,ItemInfo> objInfo = new Dictionary<string,ItemInfo>();
            SAPbouiCOM.Matrix objMatrixItem = null;
            SAPbouiCOM.Item objItem = null;
            SAPbouiCOM.Item objItemRemarks = null;

            try
            {
                if (
                    BusinessObjectInfo.FormTypeEx == "133"
                    && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    && BusinessObjectInfo.BeforeAction
                    )
                {
                    objForm = MainObject.Instance.B1Application.Forms.Item(BusinessObjectInfo.FormUID);
                    objItem = objForm.Items.Item("38");
                    objItemRemarks = objForm.Items.Item("16");
                    objMatrixItem = objItem.Specific;
                    objMatrix.LoadXml(objMatrixItem.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All));
                    XmlNodeList objNode = objMatrix.SelectNodes("/Matrix/Rows/Row[./Visible=1]/Columns/Column[./ID=1]/Value");
                    if (objNode.Count > 0)
                    {
                        
                            #region Query Dimension per line
                            foreach (XmlNode xn in objNode)
                            {
                                string strItemCode = xn.InnerText;
                                #region Consulta

                                string strSQL = "select '18' as 'WhsCode','AdmGF' as 'Dim1','CompGVV' as 'Dim2','MercGVN' as 'Dim3', 'CompFIN' as 'Dim4' ";
                                SAPbobsCOM.Recordset objRec = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                objRec.DoQuery(strSQL);


                                #endregion Consulta


                                //Hacer la Busqueda para encontrar los centros de costo
                                ItemInfo objTemp = new ItemInfo();
                                objTemp.WareHouse = objRec.Fields.Item("WhsCode").Value;
                                objTemp.Dimension1 = objRec.Fields.Item("Dim1").Value;
                                objTemp.Dimension2 = objRec.Fields.Item("Dim2").Value;
                                objTemp.Dimension3 = objRec.Fields.Item("Dim3").Value;
                                objTemp.Dimension4 = objRec.Fields.Item("Dim4").Value;

                                objInfo.Add(strItemCode, objTemp);

                                
                            }
                        #endregion

                        objForm.Freeze(true);

                        int intVisualRowCount = objMatrixItem.RowCount;

                        for (int i = 1; i < intVisualRowCount; i++)
                        {
                            SAPbouiCOM.EditText oItem = objMatrixItem.Columns.Item("1").Cells.Item(i).Specific;
                            string strItemCode = oItem.Value;
                            ItemInfo oItm = objInfo[strItemCode];

                            SAPbouiCOM.EditText oTemp = objMatrixItem.Columns.Item("110000310").Cells.Item(i).Specific;
                            oTemp.Value = oItm.Dimension1;
                            
                            oTemp = objMatrixItem.Columns.Item("10002039").Cells.Item(i).Specific;
                            oTemp.Value = oItm.Dimension2;
                            oTemp = objMatrixItem.Columns.Item("10002041").Cells.Item(i).Specific;
                            oTemp.Value = oItm.Dimension3;
                            oTemp = objMatrixItem.Columns.Item("10002043").Cells.Item(i).Specific;
                            oTemp.Value = oItm.Dimension4;
                        }

                        objForm.Freeze(false);



                        //for (int i =0; i < intVisualRowCount; i++)
                        //{   
                        //    objMatrixItem.SetCellFocus(i,)


                        //        objInvoice.Lines.SetCurrentLine(i);
                        //    ItemInfo objTemp2 = objInfo[objInvoice.Lines.ItemCode];
                        //    objInvoice.Lines.CostingCode = objTemp2.Dimension1;
                        //    objInvoice.Lines.CostingCode2 = objTemp2.Dimension2;
                        //    objInvoice.Lines.CostingCode3 = objTemp2.Dimension3;
                        //    objInvoice.Lines.CostingCode4 = objTemp2.Dimension4;
                        //}

                        //int intRes = objInvoice.Update();

                        //string strMsg = MainObject.Instance.B1Company.GetLastErrorDescription();

                        //BubbleEvent = false;

                        
                    }
                    

                    
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                if(objForm != null)
                objForm.Freeze(false);
            }
        }


        internal class ItemInfo
        {
            public string WareHouse { get; set; }
            public string Dimension1 { get; set; }
            public string Dimension2 { get; set; }
            public string Dimension3 { get; set; }
            public string Dimension4 { get; set; }
            public string Dimension5 { get; set; }
        }



        void objApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            //T1.B1.MenuManager.Operations objMenuManager = new MenuManager.Operations();
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
                _Logger.Error("", comEx);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

    }
}
