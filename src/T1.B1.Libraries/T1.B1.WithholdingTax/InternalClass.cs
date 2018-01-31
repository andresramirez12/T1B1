using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.WithholdingTax
{
    public class EventInfoClass
    {
        public string ColUID { get; set; }
        public string FormUID { get; set; }
        public string ItemUID { get; set; }
        public int Row { get; set; }
    }

    public class WithholdingTaxConfigDetail
    {
        public string WTCode { get; set; }
        public string U_BYBCOMM { get; set; }
        public double U_BYB_MIN { get; set; }
        public string U_BYB_TIPO { get; set; }
        public string U_BYB_MUNI { get; set; }
        public string U_BYB_AFEC { get; set; }
        public List<WithholdingTaxConfigMun> MUNI { get; set; }
    }

    public class WithholdingTaxConfigMun
    {
        public string U_MUNCODE { get; set; }

    }

    public class B1WithHoldingInfoMatrixLine
    {
        public double BaseAmount { get; set; }
        public double WTAmount { get; set; }
    }

    
    public class SelfWithholdingTaxInfo
    {
        public string Code { get; set; }
        public string Debit { get; set; }
        public string Credit { get; set; }
        public double Percentage { get; set; }

        public int DocEntry { get; set; }
        public int DocNum { get; set; }
        public string CardCode { get; set; }
        public double dbWtAmount { get; set; }
        public double dbBaseAmount { get; set; }
        public string DocType { get; set; }
        public double MinMount { get; set; }
    }

    public class SelfWithholdingTaxTransaction
    {
        public int JournalEntry { get; set; }
        public double BaseAmount { get; set; }
        public bool Cancelled { get; set; }
        public string DocType { get; set; }
        public int DocEntry { get; set; }
        public string CardCode { get; set; }
        public string ThirdParty { get; set; }
        public string Code { get; set; }
        public double Total { get; set; }
        public int DocNum { get; set; }
        public string DocSeries { get; set; }
    }

    public class SelfWothholdingTaxResult
    {
        public string Message { get; set; }
        public string MessageCode { get; set; }
    }


    public class AddDocumentInfoArgs
    {
        public string ObjectType { get; set; }
        public string ObjectKey { get; set; }
        public string FormtTypeEx { get; set; }
        public string FormUID { get; set; }

    }

    public class InternalRegistryWTData
    {
        public double WTAmount { get; set; }
        public double PercentFromCode { get; set; }
        public double WT { get; set; }
    }
}
