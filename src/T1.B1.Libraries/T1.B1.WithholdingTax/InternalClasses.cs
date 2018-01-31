using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.WithholdingTax
{
    public class WithholdingTaxConfigDetail
    {
        public string Code { get; set; }
        
        public string U_TYPEMM { get; set; }
        public double U_MONTOMIN { get; set; }
        public string U_TYPEINT { get; set; }
        public List<WithholdingTaxConfigMun> MUNI { get; set; }
    }

    public class WithholdingTaxConfigMun
    {
        public string Code { get; set; }
        public string U_MUNI { get; set; }

    }



    public class SelfWithholdingTaxInfo
    {
        public string Code { get; set; }
        public string Debit { get; set; }
        public string Credit { get; set; }
        public double Percentage { get; set; }

        public int DocEntry { get; set; }
        public int DocNum {get;set;}
        public string CardCode { get; set; }
        public double dbWtAmount { get; set; }
        public double dbBaseAmount { get; set; }
        public string DocType { get; set; }
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
}
