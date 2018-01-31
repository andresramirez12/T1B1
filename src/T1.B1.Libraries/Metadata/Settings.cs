using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.MetaData
{
    public class Settings
    {
        public static Main _Main { get; }

        static Settings()
        {

            _Main = new Main();
            _Main.Initialize();
            
        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public Main()
            {
                logLevel = "Debug";
                resourceFodler = "MetaData";
                updateUDOForm = false;
                singleFile = true;
                singleFileName = "MDInfo.xml";
                loadMuniUDO = "BYB_T1RPA200";
                loadDeptUDO = "BYB_T1RPA201";




            }

            public string logLevel { get; }
            public string resourceFodler { get; }
            public bool updateUDOForm { get;}
            public bool singleFile { get; }
            public string singleFileName { get; }
            public string loadMuniUDO { get; set; }
            public string loadDeptUDO { get; set; }


        }
    }
}
