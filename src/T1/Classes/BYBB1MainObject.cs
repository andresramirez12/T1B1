using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace T1.Classes
{
    public sealed class BYBB1MainObject
    {
        private static readonly Lazy<BYBB1MainObject> lazy =
            new Lazy<BYBB1MainObject>(() => new BYBB1MainObject());

        private static SAPbouiCOM.Application objB1Application = null;
        private static SAPbobsCOM.Company objB1Company = null;
        

        public static BYBB1MainObject Instance
        {
            get
            {
                return lazy.Value;
            }
        }

        public SAPbobsCOM.Company B1Company
        {
            get
            {
                return objB1Company;
            }
            set
            {
                objB1Company = value;
            }
        }

        public SAPbouiCOM.Application B1Application
        {
            get
            {
                return objB1Application;
            }
            set
            {
                objB1Application = value;
            }
        }

        private BYBB1MainObject()
        {
        }






    }
}
