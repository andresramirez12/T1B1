using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace T1.B1.InternalClasses
{
    public class AdminInfo
    {
        public string DecimalSeparator { get; set; }
        public string ThousandsSeparator { get; set; }
        public string DateSeparator { get; set; }
        public string SystemCurrency { get; set; }
        public int RateAccuracy { get; set; }
        public int QueryAccuracy { get; set; }
        public int AccuracyofQuantities { get; set; }
        public int PercentageAccuracy { get; set; }
        public int PriceAccuracy { get; set; }
        public int TotalsAccuracy { get; set; }
        public string LocalCurrency { get; set; }

        public SAPbobsCOM.BoDateTemplate DateTemplate { get; set; }
        public bool DisplayCurrencyontheRight { get; set; }
        public string FederalTaxID { get; set; }
        public int MeasuringAccuracy { get; set; }
        public string State { get; set; }
    }

    public class projectInfo
    {
        public string ProjectCode { get; set; }
        public string ProjectName { get; set; }
    }

    public class costCenterInfo
    {
        public string CostCenterCode { get; set; }
        public string CostCenterName { get; set; }
        public int DimensionCode { get; set; }
    }

    public class dimensionInfo
    {
        public int DimentionCode { get; set; }
        public string DimensionName { get; set; }
        public bool isActive { get; set; }
    }

}
