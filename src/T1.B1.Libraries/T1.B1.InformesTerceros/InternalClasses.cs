using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.InformesTerceros
{
    public enum DateType
    {
        DocumentDate, TransactionDate,  DueDate
    }

    public class TransactionDetail
    {
        public int Level { get; set; }
        public string Account { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }
        public int Day { get; set; }
        public DateType DateType {get;set;}
        public double CurrentBalance { get; set; }
        public double PreviousBalance { get; set; }
        public string Thirdparty { get; set; }
        public string TransactionCode { get; set; }
        public double Debit { get; set; }
        public double Credit { get; set; }
        public string DIM1 { get; set; }
        public string DIM2 { get; set; }
        public string DIM3 { get; set; }
        public string DIM4 { get; set; }
        public string DIM5 { get; set; }
        

    }

    public class TransactionControl
    {
        public int LastJE { get; set; }
        public DateTime LastUpdate { get; set; }
        public string TranscationType { get; set; }
    }

    public class JsonQueryConfig
    {
        public string Name { get; set; }
        public string Query { get; set; }
    }
}
