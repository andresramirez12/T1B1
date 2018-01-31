using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.EventFilter
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
                




            }

            public string logLevel { get; }
            
            
        }
    }
}
