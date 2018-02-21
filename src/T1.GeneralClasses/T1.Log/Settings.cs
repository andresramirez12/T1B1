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
        public static string LogPath { get; set; }
        public static Main _Main { get; set; }

        static Settings()
        {

            _Main = new Main();

            if(_Main.useAppData)
            {
                AppDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\T1\\json\\";
                LogPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\T1\\logs\\";
                
            }
            else
            {
                AppDataPath = AppDomain.CurrentDomain.BaseDirectory + "\\T1\\json\\";
                LogPath = AppDomain.CurrentDomain.BaseDirectory + "\\T1\\logs\\";
                
            }

            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }
            if (!Directory.Exists(LogPath))
            {
                Directory.CreateDirectory(Settings.LogPath);
            }
            _Main.Initialize();
            _Main.logFolder = LogPath;
            _Main.jsonFolder = AppDataPath;

        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public string logFolder { get; set; }
            public string jsonFolder { get; set; }
            public string masterLogName { get; set; }
            public string pattern { get; set; }
            public string masterSize { get; set; }
            public string appenderSufix { get; set; }
            public int numberOfLogs { get; set; }
            public bool useAppData { get; set; }
            
            public Main()
            {
                logFolder = "";
                jsonFolder = "";
                masterLogName = "T1.log";
                pattern = "%date [%thread] %level %logger - %ndc - %message [%line] %newline";
                masterSize = "10MB";
                appenderSufix = "App";
                numberOfLogs = 10;
                useAppData = true;
                
            }

            protected override void OnInitialize(Westwind.Utilities.Configuration.IConfigurationProvider provider, string sectionName, object configData)
            {
                var providerJSON = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = providerJSON;
                base.OnInitialize(providerJSON, sectionName, configData);
            }

            
        }
    }
}
