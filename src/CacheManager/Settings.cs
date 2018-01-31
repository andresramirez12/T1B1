using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.CacheManager
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
                useAppDomain = false;
                connStringCacheName = "connString";
                isHANACacheName = "isHANA";
                progressBarCacheName = "progressbar";







            }

            public string logLevel { get; }
            public bool useAppDomain { get; }
            public string connStringCacheName { get; }
            public string isHANACacheName { get; }
            public string progressBarCacheName { get; }



        }
    }
}
