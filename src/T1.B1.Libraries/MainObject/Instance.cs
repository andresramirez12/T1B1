using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace T1.B1
{
    public sealed class MainObject
    {
        private static readonly Lazy<MainObject> lazy =
            new Lazy<MainObject>(() => new MainObject());

        private static SAPbouiCOM.Application objB1Application = null;
        private static SAPbobsCOM.Company objB1Company = null;
        private T1.B1.InternalClasses.AdminInfo objB1AdmInfo = null;
        


        public static MainObject Instance
        {
            get
            {
                return lazy.Value;
            }
        }

        public T1.B1.InternalClasses.AdminInfo B1AdminInfo
        {
            get
            {
                return objB1AdmInfo;
            }
            set
            {
                objB1AdmInfo = value;
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

        private MainObject()
        {
        }

        






    }
}
