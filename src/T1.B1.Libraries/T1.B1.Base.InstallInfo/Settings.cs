using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace T1.B1.Base.InstallInfo
{
    public class Settings
    {
        public static Main _Main { get; }
        public static string AppDataPath { get; set; }

        static Settings()
        {
            _Main = new Main();
            _Main.Initialize();

            AppDataPath = AppDomain.CurrentDomain.BaseDirectory + _Main.configurationBaseFolder;
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }

            
        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            
            public Main()
            {
                logLevel = "Debug";
                nancyLocalAddress = "http://localhsot:9001";
                configurationBaseFolder = "\\BYB\\T1\\";

            }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }
            public string logLevel { get; }
            public string nancyLocalAddress { get; }
            public string configurationBaseFolder { get; }



        }

        


    }
}
