using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace T1.B1.Connection
{
    public class Settings
    {
        public static Main _Main { get; set; }
        public static string AppDataPath { get; set; }

        static Settings()
        {

            AppDataPath = AppDomain.CurrentDomain.BaseDirectory + "\\BYB\\AdminService\\T1\\";
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
                logLevel = "Debug";
                jobId = "T1.B1.Connection.J01";
                groupId = "T1.B1.Connection.G01";
                triggerId = "T1.B1.Connection.T01";
                cron = "0 30 * * * ?";
                createMD = false;
                connectionDirectory = "\\config\\connections\\";
                conectionNameToIdDictionary = "conectionNameToIdDictionary";
                conectionInformationDictionary = "conectionInformationDictionary";
                isEncrypted = false;
                isHANA = false;
                HANADriver = "";


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
            public string jobId { get; }
            public string groupId { get; }
            public string triggerId { get; }
            public string cron { get; }
            public bool createMD { get; set; }
            public string connectionDirectory { get; }
            public string conectionNameToIdDictionary { get; }
            public string conectionInformationDictionary { get; }
            public bool isEncrypted { get; }
            public bool isHANA { get; }
            public string HANADriver { get; }



        }
        
    }
}
