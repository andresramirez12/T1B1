using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace T1
{
    public class Settings
    {
        public static Main _Main { get; set; }
        public static string AppDataPath { get; set; }

        static Settings()
        {

            AppDataPath = T1.Log.Settings._Main.jsonFolder;
            _Main = new Main();
            _Main.Initialize();
            
        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public string logLevel { get; }
            public string connStringCacheName { get; }
            public bool createMD { get; set; }
            public bool loadInitialData { get; set; }

            public Main()
            {
                logLevel = "Error";
                connStringCacheName = "connectionString";
                createMD = false;
                loadInitialData = false;
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
