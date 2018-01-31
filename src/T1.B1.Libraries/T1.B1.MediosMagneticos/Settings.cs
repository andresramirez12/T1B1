using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.MediosMagneticos
{
    public class Settings
    {
        public static Main _Main { get; }
        public static SelfWithHoldingTax _SelfWithHoldingTax { get; }

        static Settings()
        {

            _Main = new Main();
            _Main.Initialize();

            _SelfWithHoldingTax = new SelfWithHoldingTax();
            _SelfWithHoldingTax.Initialize();


        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public Main()
            {
                logLevel = "Debug";
                




            }

            public string logLevel { get; }
            
        }

        public class SelfWithHoldingTax : Westwind.Utilities.Configuration.AppConfiguration
        {
            public SelfWithHoldingTax()
            {
                
                getSelfWithHoldingTaxQuery = "SELECT [@BYB_T1SWT100].\"Code\" ,U_CreditAccount,U_DebitAcct,U_Percent FROM [@BYB_T1SWT100] inner join [@BYB_T1SWT101] on [@BYB_T1SWT100].\"Code\" = [@BYB_T1SWT101].\"Code\" where [@BYB_T1SWT100].U_Enabled = 'Y' and [@BYB_T1SWT101].U_CardCode = '[--CardCode--]'";
                WTaxTransCode = "T1SW";
                SWtaxUDO = "";
                SWtaxUDOTransaction = "BYB_T1SWTU002";
                CancelFormUID = "BYB_SWTF001";
                getPostedSWtaxQueryV1 = "select distinct 'N' as \"Sel.\",TransId as \"Asiento\", TaxDate as \"Fecha\", case when credit > 0 then Credit else Debit end as Total, LineMemo as \"Comentario\", space(500) as Resultado from JDT1 where LineMemo like '%Auto%[--SWTCode--]%' and(TaxDate >= Convert(datetime, '[--StartDate--]', 112) and TaxDate <= Convert(datetime, '[--EndDate--]', 112))";
                getPostedSWtaxQueryV2 = "select distinct 'N' as \"Sel.\",TransId as \"Asiento\", TaxDate as \"Fecha\", case when credit > 0 then Credit else Debit end as Total, LineMemo as \"Comentario\", space(500) as Resultado from JDT1 where TransCode = '[--TransCode--]' and(TaxDate >= Convert(datetime, '[--StartDate--]', 112) and TaxDate <= Convert(datetime, '[--EndDate--]', 112))";
                TransactionCodeBase = false;
                getRegistrationFromJEQuery = "SELECT \"DocEntry\" from [@BYB_T1SWT200] where \"U_JEEntry\" = [--JE--]";
                getWTaxDocuments = "T1.B1.WithholdingTax.QRY001";
                getSelfWithHoldingTransactions = "select U_DocEntry, U_DocNum from [@BYB_T1SWT200] where Canceled = 'N' and U_DocDate >= convert(datetime, '[--StartDate--]', 112) and U_DocDate <= convert(datetime, '[--EndDate--]', 112)";
            }

            
            public string getSelfWithHoldingTaxQuery { get; set; }
            public string WTaxTransCode { get; set; }
            public string SWtaxUDO { get; set; }
            public string CancelFormUID { get; }
            public string getPostedSWtaxQueryV1 { get; }
            public string getPostedSWtaxQueryV2 { get; }
            public bool TransactionCodeBase { get; }
            public string getRegistrationFromJEQuery { get; }
            public string SWtaxUDOTransaction { get; set; }
            public string getWTaxDocuments { get; set; }
            public string getSelfWithHoldingTransactions { get; set; }
            


        }


    }
}
