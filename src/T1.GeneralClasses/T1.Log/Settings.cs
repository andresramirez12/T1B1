using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace T1.Log
{
    public class Settings
    {
        public static string AppDataPath { get; set; }
        public static Main _Main { get; set; }

        static Settings()
        {
            AppDataPath = AppDomain.CurrentDomain.BaseDirectory;
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }


            _Main = new Main();
            _Main.Initialize();
            
        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public Main()
            {
                logFolder = "\\logFiles\\";
                masterLogName = "T1.log";
                pattern = "%date [%thread] %level %logger - %ndc - %message%newline";
                masterSize = "10MB";
                appenderSufix = "App";
            }

            public string logFolder { get; }
            public string masterLogName { get; }
            public string pattern { get;}
            public string masterSize { get;}
            public string appenderSufix { get; }
        }
    }
}
