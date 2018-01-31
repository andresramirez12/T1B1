using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.Connection
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
                useCompatibilityConnection = false;
                useCompanyApplication = true;
                isHANACache = "isHANA";








            }

            public string logLevel { get; }
            public bool useCompatibilityConnection { get; }
            public bool useCompanyApplication { get; }
            public string isHANACache { get; }






        }
    }
}
