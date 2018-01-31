using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Runtime.InteropServices;
using System.Xml;
using SAPbobsCOM;
using System.Globalization;
using System.IO;
using Newtonsoft.Json;

namespace T1.B1.InformesTerceros
{
    class BalanceTerceros
    {
        private static BalanceTerceros objWithHoldingTax;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private BalanceTerceros()
        {
            
        }

        public static Tuple<string,string,string,string,string> getAccountSegments(string AccountCode)
        {
            //This function must be upgraded to segmentation in the future

            Tuple<string, string, string, string, string> oAccounts = null;
            try
            {
                string strLevel1 = AccountCode.Substring(0, 1);
                string strLevel2 = AccountCode.Substring(0, 2);
                string strLevel3 = AccountCode.Substring(0, 4);
                string strLevel4 = AccountCode.Substring(0, 6);
                string strLevel5 = AccountCode;

                oAccounts = new Tuple<string, string, string, string, string>(strLevel1, strLevel2, strLevel3, strLevel4, strLevel5);

            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                oAccounts = new Tuple<string, string, string, string, string>("","","","","");
            }
            return oAccounts;
        }
        public static List<Tuple<DateType,int, int, int>> getDates(DateTime DueDate, DateTime DocDate, DateTime RefDate)
        {
            List<Tuple<DateType, int, int, int>> objTupleList = new List<Tuple<DateType, int, int, int>>(); 

            Tuple<DateType,int, int, int> oDates = null;
            try
            {
                oDates = new Tuple<DateType, int, int, int>(DateType.DueDate, DueDate.Year, DueDate.Month, DueDate.Day);
                objTupleList.Add(oDates);
                oDates = new Tuple<DateType, int, int, int>(DateType.DocumentDate, DocDate.Year, DocDate.Month, DocDate.Day);
                objTupleList.Add(oDates);
                oDates = new Tuple<DateType, int, int, int>(DateType.TransactionDate, RefDate.Year, RefDate.Month, RefDate.Day);
                objTupleList.Add(oDates);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                new List<Tuple<DateType, int, int, int>>();
            }
            return objTupleList;
        }

        public static List<string> buildCommand(TransactionDetail objDetail, Tuple<string, string, string, string, string> accountLevel, List<Tuple<DateType, int, int, int>> oTupleDateList)
        {
            List<string> strCommand = new List<string>(); ;
            
            List<JsonQueryConfig> objQueryList = new List<JsonQueryConfig>();
            try
            {
                
                    //JsonQueryConfig objFound = objQueryList.Find(o => o.Name == Settings._BalanceTerceros.upsertCurrentBalance);
                   
                
                        string strCmd = Settings._BalanceTerceros.UpsertCurrentBalance;

                        strCmd = strCmd
                            .Replace("[--TransCode--]", objDetail.TransactionCode)
                            .Replace("[--ThirdParty--]", objDetail.Thirdparty)
                            .Replace("[--Debit--]", objDetail.Debit.ToString(CultureInfo.InvariantCulture))
                            .Replace("[--Credit--]", objDetail.Credit.ToString(CultureInfo.InvariantCulture))
                            .Replace("[--CurrBalance--]", Math.Abs(objDetail.Debit - objDetail.Credit).ToString(CultureInfo.InvariantCulture));

                        foreach (Tuple<DateType, int, int, int> BoDateTemplate in oTupleDateList)
                        {
                            #region FirstLevel
                            string strInternalCmd = strCmd
                        .Replace("[--Account--]", accountLevel.Item1)
                        .Replace("[--Level--]", 1.ToString())
                            .Replace("[--Year--]", BoDateTemplate.Item2.ToString())
                            .Replace("[--Month--]", BoDateTemplate.Item3.ToString())
                            .Replace("[--Day--]", BoDateTemplate.Item4.ToString())
                            .Replace("[--DateType--]", BoDateTemplate.Item1.ToString());


                            string strDIM1 = "";
                            string strDIM2 = "";
                            string strDIM3 = "";
                            string strDIM4 = "";
                            string strDIM5 = "";




                            if (!String.IsNullOrEmpty(objDetail.DIM1))
                            {
                                strDIM1 = strInternalCmd.Replace("[--DIM1--]", objDetail.DIM1)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM1);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM2))
                            {
                                strDIM2 = strInternalCmd.Replace("[--DIM2--]", objDetail.DIM2)
                                .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM2);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM3))
                            {
                                strDIM3 = strInternalCmd.Replace("[--DIM3--]", objDetail.DIM3)
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM3);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM4))
                            {
                                strDIM4 = strInternalCmd.Replace("[--DIM4--]", objDetail.DIM4)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM4);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM5))
                            {
                                strDIM5 = strInternalCmd.Replace("[--DIM5--]", objDetail.DIM5)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM1--]", "");
                                strCommand.Add(strDIM5);
                            }




                            strCommand.Add(strInternalCmd
                                .Replace("[--DIM1--]", "")
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", ""));

                            #endregion

                            #region Second Level
                            strInternalCmd = strCmd
                        .Replace("[--Account--]", accountLevel.Item2)
                        .Replace("[--Level--]", 1.ToString())
                            .Replace("[--Year--]", BoDateTemplate.Item2.ToString())
                            .Replace("[--Month--]", BoDateTemplate.Item3.ToString())
                            .Replace("[--Day--]", BoDateTemplate.Item4.ToString())
                            .Replace("[--DateType--]", BoDateTemplate.Item1.ToString());


                            strDIM1 = "";
                            strDIM2 = "";
                            strDIM3 = "";
                            strDIM4 = "";
                            strDIM5 = "";




                            if (!String.IsNullOrEmpty(objDetail.DIM1))
                            {
                                strDIM1 = strInternalCmd.Replace("[--DIM1--]", objDetail.DIM1)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM1);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM2))
                            {
                                strDIM2 = strInternalCmd.Replace("[--DIM2--]", objDetail.DIM2)
                                .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM2);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM3))
                            {
                                strDIM3 = strInternalCmd.Replace("[--DIM3--]", objDetail.DIM3)
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM3);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM4))
                            {
                                strDIM4 = strInternalCmd.Replace("[--DIM4--]", objDetail.DIM4)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM4);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM5))
                            {
                                strDIM5 = strInternalCmd.Replace("[--DIM5--]", objDetail.DIM5)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM1--]", "");
                                strCommand.Add(strDIM5);
                            }




                            strCommand.Add(strInternalCmd
                                .Replace("[--DIM1--]", "")
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", ""));

                            #endregion

                            #region Third Level
                            strInternalCmd = strCmd
                        .Replace("[--Account--]", accountLevel.Item3)
                        .Replace("[--Level--]", 1.ToString())
                            .Replace("[--Year--]", BoDateTemplate.Item2.ToString())
                            .Replace("[--Month--]", BoDateTemplate.Item3.ToString())
                            .Replace("[--Day--]", BoDateTemplate.Item4.ToString())
                            .Replace("[--DateType--]", BoDateTemplate.Item1.ToString());


                            strDIM1 = "";
                            strDIM2 = "";
                            strDIM3 = "";
                            strDIM4 = "";
                            strDIM5 = "";




                            if (!String.IsNullOrEmpty(objDetail.DIM1))
                            {
                                strDIM1 = strInternalCmd.Replace("[--DIM1--]", objDetail.DIM1)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM1);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM2))
                            {
                                strDIM2 = strInternalCmd.Replace("[--DIM2--]", objDetail.DIM2)
                                .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM2);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM3))
                            {
                                strDIM3 = strInternalCmd.Replace("[--DIM3--]", objDetail.DIM3)
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM3);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM4))
                            {
                                strDIM4 = strInternalCmd.Replace("[--DIM4--]", objDetail.DIM4)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM4);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM5))
                            {
                                strDIM5 = strInternalCmd.Replace("[--DIM5--]", objDetail.DIM5)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM1--]", "");
                                strCommand.Add(strDIM5);
                            }




                            strCommand.Add(strInternalCmd
                                .Replace("[--DIM1--]", "")
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", ""));

                            #endregion

                            #region Fourth Level
                            strInternalCmd = strCmd
                        .Replace("[--Account--]", accountLevel.Item4)
                        .Replace("[--Level--]", 1.ToString())
                            .Replace("[--Year--]", BoDateTemplate.Item2.ToString())
                            .Replace("[--Month--]", BoDateTemplate.Item3.ToString())
                            .Replace("[--Day--]", BoDateTemplate.Item4.ToString())
                            .Replace("[--DateType--]", BoDateTemplate.Item1.ToString());


                            strDIM1 = "";
                            strDIM2 = "";
                            strDIM3 = "";
                            strDIM4 = "";
                            strDIM5 = "";




                            if (!String.IsNullOrEmpty(objDetail.DIM1))
                            {
                                strDIM1 = strInternalCmd.Replace("[--DIM1--]", objDetail.DIM1)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM1);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM2))
                            {
                                strDIM2 = strInternalCmd.Replace("[--DIM2--]", objDetail.DIM2)
                                .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM2);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM3))
                            {
                                strDIM3 = strInternalCmd.Replace("[--DIM3--]", objDetail.DIM3)
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM3);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM4))
                            {
                                strDIM4 = strInternalCmd.Replace("[--DIM4--]", objDetail.DIM4)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM4);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM5))
                            {
                                strDIM5 = strInternalCmd.Replace("[--DIM5--]", objDetail.DIM5)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM1--]", "");
                                strCommand.Add(strDIM5);
                            }




                            strCommand.Add(strInternalCmd
                                .Replace("[--DIM1--]", "")
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", ""));

                            #endregion

                            #region Fifth Level
                            strInternalCmd = strCmd
                        .Replace("[--Account--]", accountLevel.Item5)
                        .Replace("[--Level--]", 1.ToString())
                            .Replace("[--Year--]", BoDateTemplate.Item2.ToString())
                            .Replace("[--Month--]", BoDateTemplate.Item3.ToString())
                            .Replace("[--Day--]", BoDateTemplate.Item4.ToString())
                            .Replace("[--DateType--]", BoDateTemplate.Item1.ToString());


                            strDIM1 = "";
                            strDIM2 = "";
                            strDIM3 = "";
                            strDIM4 = "";
                            strDIM5 = "";




                            if (!String.IsNullOrEmpty(objDetail.DIM1))
                            {
                                strDIM1 = strInternalCmd.Replace("[--DIM1--]", objDetail.DIM1)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM1);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM2))
                            {
                                strDIM2 = strInternalCmd.Replace("[--DIM2--]", objDetail.DIM2)
                                .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM2);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM3))
                            {
                                strDIM3 = strInternalCmd.Replace("[--DIM3--]", objDetail.DIM3)
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM3);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM4))
                            {
                                strDIM4 = strInternalCmd.Replace("[--DIM4--]", objDetail.DIM4)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM1--]", "")
                                    .Replace("[--DIM5--]", "");
                                strCommand.Add(strDIM4);
                            }
                            if (!String.IsNullOrEmpty(objDetail.DIM5))
                            {
                                strDIM5 = strInternalCmd.Replace("[--DIM5--]", objDetail.DIM5)
                                    .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM1--]", "");
                                strCommand.Add(strDIM5);
                            }




                            strCommand.Add(strInternalCmd
                                .Replace("[--DIM1--]", "")
                                .Replace("[--DIM2--]", "")
                                    .Replace("[--DIM3--]", "")
                                    .Replace("[--DIM4--]", "")
                                    .Replace("[--DIM5--]", ""));

                            #endregion


                        }

                    
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                strCommand = new List<string>(); ;
            }
            return strCommand;
        }


        public static void exploteJEInformation(int JournalEntryNumber)
        {
            List<string> cmdList = new List<string>();
            SAPbobsCOM.JournalEntries objJE = null;
            TransactionDetail objDetail = null;
            Dictionary<string, string> objBPDict = null;
            string strThirdParty = "";
            try
            {
                objJE = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                if(objJE.GetByKey(JournalEntryNumber))
                {
                    objBPDict = T1.B1.ReletadParties.Instance.getBPThirdPartyRelation();
                    for (int i = 0; i < objJE.Lines.Count; i++)
                    {
                        objJE.Lines.SetCurrentLine(i);
                        Tuple<string, string, string, string, string> accountLevel = getAccountSegments(objJE.Lines.AccountCode);
                        
                        objDetail = new TransactionDetail();
                        
                        objDetail.Debit = objJE.Lines.Debit;
                        objDetail.Credit = objJE.Lines.Credit;
                        objDetail.DIM1 = objJE.Lines.CostingCode;
                        objDetail.DIM2 = objJE.Lines.CostingCode2;
                        objDetail.DIM3 = objJE.Lines.CostingCode3;
                        objDetail.DIM4 = objJE.Lines.CostingCode4;
                        objDetail.DIM5 = objJE.Lines.CostingCode5;
                        objDetail.TransactionCode = objJE.TransactionCode;
                        List<Tuple<DateType, int, int, int>> oTupleDateList = getDates(objJE.Lines.DueDate, objJE.Lines.TaxDate, objJE.Lines.ReferenceDate1);

                        #region ThirdParty Assignment
                        if (!String.IsNullOrEmpty(objJE.Lines.ShortName))
                        {
                            if (objBPDict != null && objBPDict.ContainsKey(objJE.Lines.ShortName))
                            {
                                strThirdParty = objBPDict[objJE.Lines.ShortName];
                            }
                        }
                        
                        try
                        {
                            if (!String.IsNullOrEmpty(objJE.Lines.UserFields.Fields.Item("U_BYB_RELPAR").Value))
                            {
                                strThirdParty = objJE.Lines.UserFields.Fields.Item("U_BYB_RELPAR").Value;
                            }
                        }
                        catch (Exception er)
                        {
                            _Logger.Error("The UDF U_BYB_RELPAR was not found in DB");
                        }
                        #endregion

                        List<string> strResult = buildCommand(objDetail, accountLevel, oTupleDateList);
                        cmdList.AddRange(strResult);
                    }
                    
                    SAPbobsCOM.Recordset objRs = null;
                    foreach (string strCommand in cmdList)
                    {
                        try
                        {
                            objRs = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                            objRs.DoQuery(strCommand);
                        }
                        catch(Exception er)
                        {
                            _Logger.Error("Error while executing command " + strCommand + " for JE " + JournalEntryNumber.ToString(), er);
                        }
                    }
                    string strComm = Settings._BalanceTerceros.upsertObjectControl
                        .Replace("[--LastTrans--]", JournalEntryNumber.ToString());

                    objRs = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    objRs.DoQuery(strComm);
                    objRs = null;
                    cmdList = null;

                }
                else
                {
                    _Logger.Error("Could not retrieve JE number " + JournalEntryNumber);
                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                cmdList = new List<string>();

            }
            finally
            {
                if(objJE != null)
                {
                    objJE = null;
                }
            }
            //return cmdList;
        }

        public static List<int> getTransactionList()
        {
            SAPbobsCOM.Recordset objRSControl = null;
            List<int> lTransaction = null;
            string strSQL = "";
            try
            {
                lTransaction = new List<int>();
                strSQL = Settings._BalanceTerceros.getTransactionList;
                objRSControl = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                if(objRSControl != null)
                {
                    objRSControl.DoQuery(strSQL);
                    if(objRSControl.RecordCount > 0)
                    {
                        while(!objRSControl.EoF)
                        {
                            int intTransId = objRSControl.Fields.Item(0).Value;
                            if (!lTransaction.Contains(intTransId))
                            {
                                lTransaction.Add(intTransId);
                            }
                            objRSControl.MoveNext();
                        }
                        objRSControl = null;
                        //List<string> lCommands = new List<string>();
                        foreach(int JDTId in lTransaction)
                        {
                            //lCommands.AddRange(exploteJEInformation(JDTId));
                            exploteJEInformation(JDTId);

                        }
                        int j = 0;
                    }
                }
                else
                {
                    _Logger.Error("Could not initialize RecordSet");
                }


            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                lTransaction = new List<int>();
            }
            return lTransaction;
        }
    }
}
