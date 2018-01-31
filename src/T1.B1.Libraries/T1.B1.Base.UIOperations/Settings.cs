using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.Base.UIOperations
{
    public class Settings
    {
        public static Main _Main { get; }
        public static FormLoad _FormLoad { get; }

        static Settings()
        {

            _Main = new Main();
            _Main.Initialize();

            _FormLoad = new FormLoad();
            _FormLoad.Initialize();




        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public Main()
            {
                logLevel = "Debug";
                




            }

            public string logLevel { get; }
            
        }

        public class FormLoad : Westwind.Utilities.Configuration.AppConfiguration
        {
            public FormLoad()
            {
                errorPath = "/result/errors";
            }
            public string errorPath { get; }
        }

        


    }
}
