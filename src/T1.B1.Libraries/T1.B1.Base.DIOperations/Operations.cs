using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using log4net;
using System.Runtime.InteropServices;
using SAPbobsCOM;

namespace T1.B1.Base.DIOperations
{
    public class Operations
    {


        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        
        private Operations()
        {

        }

        public static bool addUDOInfo(object UDOInfo, string UDOName)
        {
            bool blResult = false;
            CompanyService objCompanyService = null;
            GeneralService UDOService = null;
            GeneralData headerInfo = null;
            GeneralDataParams addResult = null;

            try
            {
                
                    objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                    UDOService = objCompanyService.GetGeneralService(UDOName);
                    headerInfo = UDOService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                    foreach (var prop in UDOInfo.GetType().GetProperties())
                    {
                        headerInfo.SetProperty(prop.Name, prop.GetValue(UDOInfo,null));
                    }
                    addResult = UDOService.Add(headerInfo);

                    



            }
            catch (Exception er)
            {
                
                _Logger.Error("", er);
            }


            return blResult;
        }

        public static GeneralData getUDOInfo(object objKey, string UDOName)
        {
            GeneralData objResult = null;
            CompanyService objCompanyService = null;
            GeneralService UDOService = null;
            GeneralDataParams getInfo = null;

            try
            {

                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                UDOService = objCompanyService.GetGeneralService(UDOName);
                foreach (var prop in objKey.GetType().GetProperties())
                {
                    getInfo.SetProperty(prop.Name, prop.GetValue(objKey, null));
                }
                objResult = UDOService.GetByParams(getInfo);
            }
            catch(COMException comEx)
            {
                _Logger.Error("", comEx);
            }
            catch (Exception er)
            {
                
                _Logger.Error("", er);
            }
            return objResult;
        }

        public static bool addupdTransactionCode(string TransCode, string TransCodeDescription)
        {
            bool blResult = false;
            SAPbobsCOM.TransactionCodesService objTransCodeService = null;
            SAPbobsCOM.CompanyService objCompanyService = null;
            SAPbobsCOM.TransactionCodeParams objParamas = null;
            SAPbobsCOM.TransactionCode objTransCode = null;

            

            try
            {
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();
                objTransCodeService = objCompanyService.GetBusinessService(ServiceTypes.TransactionCodesService);
                objParamas = objTransCodeService.GetDataInterface(TransactionCodesServiceDataInterfaces.tcsTransactionCodeParams);
                objParamas.Code = TransCode;
                objTransCode = objTransCodeService.Get(objParamas);
                objTransCode.Description = TransCodeDescription;
                objTransCodeService.Update(objTransCode);
                
            }
            catch(COMException comEx)
            {
                if(comEx.ErrorCode == -2028)
                {
                    objTransCode = objTransCodeService.GetDataInterface(TransactionCodesServiceDataInterfaces.tcsTransactionCode);
                    objTransCode.Code = TransCode;
                    objTransCode.Description = TransCodeDescription;
                    try
                    {
                        objParamas = objTransCodeService.Add(objTransCode);
                    }
                    catch (COMException comE)
                    {
                        _Logger.Error("", comE);
                    }
                    catch(Exception er)
                    {
                        _Logger.Error("", er);
                    }


                    }
                else
                {
                    _Logger.Error("", comEx);
                }
                
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
            

            return blResult;
        }

        public static double roundValue(double LocalAmount, RoundingContextEnum B1Context)
        {
            double dbSCAmount = 0;
            
            CompanyService objCompanyService = null;
            
            try
            {
                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();

                DecimalData objDecimalData = objCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiDecimalData);
                objDecimalData.Context = B1Context;
                objDecimalData.Currency = MainObject.Instance.B1AdminInfo.SystemCurrency;
                objDecimalData.Value = LocalAmount;
                RoundedData objResult = objCompanyService.RoundDecimal(objDecimalData);

                dbSCAmount = objResult.Value;

            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
            finally
            {
                if (objCompanyService != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objCompanyService);
                    objCompanyService = null;
                }
               
            }
            return dbSCAmount;
        }

        public static double getSCValue(double LocalAmount, DateTime date, out string Message, RoundingContextEnum B1Context)
        {
            double dbSCAmount = 0;
            SAPbobsCOM.SBObob objBob = null;
            SAPbobsCOM.Recordset objRS = null;
            CompanyService objCompanyService = null;
            Message = "";
            try
            {
                objBob = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoBridge);
                objRS = objBob.GetCurrencyRate(MainObject.Instance.B1AdminInfo.SystemCurrency, date);
                double dbRate = objRS.Fields.Item(0).Value;

                objCompanyService = MainObject.Instance.B1Company.GetCompanyService();

                DecimalData objDecimalData = objCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiDecimalData);
                objDecimalData.Context = B1Context;
                objDecimalData.Currency = MainObject.Instance.B1AdminInfo.SystemCurrency;
                objDecimalData.Value = LocalAmount / dbRate;
                RoundedData objResult = objCompanyService.RoundDecimal(objDecimalData);

                dbSCAmount = objResult.Value;

            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                Message = er.Message;
            }
            finally
            {
                if (objCompanyService != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objCompanyService);
                    objCompanyService = null;
                }
                if (objRS != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objRS);
                    objRS = null;
                }
                if (objBob != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objBob);
                    objBob = null;
                }
            }
            return dbSCAmount;
        }

        public static double getWTAmountLC(string WTCode, double dbAmount)
        {
            double dbResult = 0;
            SAPbobsCOM.WithholdingTaxCodes objWHT = null;
            try
            {

                objWHT = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                if (objWHT.GetByKey(WTCode))
                {
                    double dbPercent = objWHT.BaseAmount;
                    dbResult = dbAmount * (dbPercent / 100);
                }
            }
            catch (Exception er)
            {
                _Logger.Error(er.Message);
                dbResult = 0;
            }
            finally
            {
                if(objWHT != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objWHT);
                    objWHT = null;
                }
            }
            
            return dbResult;
        }

        public static double getTaxAmountLC(string TaxCode, double dbAmount)
        {
            double dbResult = 0;
            SAPbobsCOM.SalesTaxCodes objVAT = null;
            try
            {
                objVAT = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesTaxCodes);
                if (objVAT.GetByKey(TaxCode))
                {
                    double dbPercent = objVAT.Rate;
                    dbResult = dbAmount * (dbPercent / 100);

                }

            }
            catch (Exception er)
            {
                _Logger.Error(er.Message);
                dbResult = 0;
            }
            finally
            {
                if (objVAT != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objVAT);
                    objVAT = null;
                }
            }
            return dbResult;
        }

        public static bool isAccountAsociated(string AccountCode)
        {
            bool blResult = false;
            SAPbobsCOM.ChartOfAccounts objAcct = null;
            try
            {
                objAcct = MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oChartOfAccounts);
                if(objAcct.GetByKey(AccountCode))
                {
                    if(objAcct.LockManualTransaction == BoYesNoEnum.tYES)
                    {
                        blResult = true;
                    }
                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                blResult = false;
            }
            return blResult;
        }

        public static void getProjectList()
        {
            List<projectInfo> objResult = null;
            SAPbobsCOM.CompanyService objCmpService = null;
            SAPbobsCOM.ProjectsService objProjectService = null;
            SAPbobsCOM.ProjectsParams objProjectList = null;
            SAPbobsCOM.ProjectParams objProjectInfo = null;

            try
            {
                objResult = new List<projectInfo>();
                objCmpService = MainObject.Instance.B1Company.GetCompanyService();
                objProjectService = objCmpService.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService);
                objProjectList = objProjectService.GetProjectList();
                for (int i = 0; i < objProjectList.Count; i++)
                {
                    objProjectInfo = objProjectList.Item(i);
                    projectInfo oProj = new projectInfo();
                    oProj.ProjectCode = objProjectInfo.Code;
                    oProj.ProjectName = objProjectInfo.Name;
                    objResult.Add(oProj);

                }


            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                objResult = new List<projectInfo>();
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objResult = new List<projectInfo>();
            }
            finally
            {
                CacheManager.CacheManager.Instance.addToCache("ProjectList", objResult, CacheManager.CacheManager.objCachePriority.NotRemovable);
            }
            
        }

        public void getProfitCenterList()
        {
            List<costCenterInfo> objResult = null;
            SAPbobsCOM.CompanyService objCmpService = null;
            SAPbobsCOM.ProfitCentersService objProfitCenterService = null;
            SAPbobsCOM.ProfitCentersParams objProfitCenterList = null;
            SAPbobsCOM.ProfitCenterParams objProfitCenterInfo = null;


            try
            {
                objResult = new List<costCenterInfo>();
                objCmpService = MainObject.Instance.B1Company.GetCompanyService();
                objProfitCenterService = objCmpService.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService);
                objProfitCenterList = objProfitCenterService.GetProfitCenterList();
                for (int i = 0; i < objProfitCenterList.Count; i++)
                {
                    objProfitCenterInfo = objProfitCenterList.Item(i);
                    SAPbobsCOM.ProfitCenterParams objTemp = objProfitCenterService.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenterParams);
                    objTemp.CenterCode = objProfitCenterInfo.CenterCode;
                    SAPbobsCOM.ProfitCenter objProfitCenter = objProfitCenterService.GetProfitCenter(objTemp);
                    costCenterInfo oProfit = new costCenterInfo();
                    oProfit.CostCenterCode = objProfitCenter.CenterCode;
                    oProfit.CostCenterName = objProfitCenter.CenterName;
                    oProfit.DimensionCode = objProfitCenter.InWhichDimension;


                    objResult.Add(oProfit);

                }


            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                objResult = new List<costCenterInfo>();
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objResult = new List<costCenterInfo>();
            }
            finally
            {
                CacheManager.CacheManager.Instance.addToCache("ProfitCenterList", objResult, CacheManager.CacheManager.objCachePriority.NotRemovable);
            }
        }

        public void getDimensionsList()
        {
            List<dimensionInfo> objResult = null;
            SAPbobsCOM.CompanyService objCmpService = null;
            SAPbobsCOM.DimensionsService objPDimensionsService = null;
            SAPbobsCOM.DimensionsParams objDimensionsList = null;
            SAPbobsCOM.DimensionParams objDimensionInfo = null;


            try
            {
                objResult = new List<dimensionInfo>();
                objCmpService = MainObject.Instance.B1Company.GetCompanyService();
                objPDimensionsService = objCmpService.GetBusinessService(SAPbobsCOM.ServiceTypes.DimensionsService);
                objDimensionsList = objPDimensionsService.GetDimensionList();
                for (int i = 0; i < objDimensionsList.Count; i++)
                {
                    objDimensionInfo = objDimensionsList.Item(i);
                    SAPbobsCOM.DimensionParams objTemp = objPDimensionsService.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams);
                    objTemp.DimensionCode = objDimensionInfo.DimensionCode;
                    SAPbobsCOM.Dimension objDimensions = objPDimensionsService.GetDimension(objTemp);
                    dimensionInfo oDImension = new dimensionInfo();
                    oDImension.DimentionCode = objDimensions.DimensionCode;
                    oDImension.DimensionName = objDimensions.DimensionDescription;
                    oDImension.isActive = objDimensions.IsActive == SAPbobsCOM.BoYesNoEnum.tYES ? true : false;

                    //oDImension.DimensionLevel = objDimensions..InWhichDimension;


                    objResult.Add(oDImension);

                }


            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                objResult = new List<dimensionInfo>();
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                objResult = new List<dimensionInfo>();
            }
            finally
            {
                CacheManager.CacheManager.Instance.addToCache("DimensionList", objResult, CacheManager.CacheManager.objCachePriority.NotRemovable);
            }
        }

        public static void isMultipleDimension()
        {
            bool blResult = false;
            bool blIsHANA = false;
            SAPbobsCOM.Recordset objRcordSet = null;
            string strSQL = "";
            string strResult = "N";
            try
            {
                blIsHANA = CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName) != null ? CacheManager.CacheManager.Instance.getFromCache(T1.CacheManager.Settings._Main.isHANACacheName) : false;
                strSQL = blIsHANA ? Settings._HANA.getMultipleDimQuery : Settings._SQL.getMultipleDimQuery;
                objRcordSet = MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRcordSet.DoQuery(strSQL);
                if (objRcordSet.RecordCount > 0)
                {
                    strResult = objRcordSet.Fields.Item(0).Value;
                    blResult = strResult.Trim() == "N" ? false : true;
                }
            }
            catch (COMException comEx)
            {
                _Logger.Error("", comEx);
                blResult = false;
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                blResult = false;
            }
            finally
            {
                CacheManager.CacheManager.Instance.addToCache("isMultiDimension", blResult, CacheManager.CacheManager.objCachePriority.NotRemovable);
            }
        }



    }
}
