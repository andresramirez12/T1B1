using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace T1.B1.Base.DIOperations
{
    public class Settings
    {
        public static Main _Main { get; set; }

        public static SQL _SQL { get; set; }
        public static HANA _HANA { get; set; }
        public static string AppDataPath { get; set; }

        static Settings()
        {
            AppDataPath = AppDomain.CurrentDomain.BaseDirectory + T1.B1.Base.InstallInfo.InstallInfo.Config.configurationBaseFolder;
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }


            _Main = new Main();
            _Main.Initialize();

            _SQL = new SQL();
            _SQL.Initialize();

            _HANA = new HANA();
            _HANA.Initialize();

        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            
            public Main()
            {
                logLevel = "Debug";
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
            public string logLevel { get; set; }



        }

        public class SQL : Westwind.Utilities.Configuration.AppConfiguration
        {

            public SQL()
            {
                getMultipleDimQuery = "Select \"UseMltDims\" from OADM";
            }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<SQL>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }
            public string getMultipleDimQuery { get; }



        }

        public class HANA : Westwind.Utilities.Configuration.AppConfiguration
        {

            public HANA()
            {
                getMultipleDimQuery = "Select \"UseMltDims\" from OADM";
            }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<HANA>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }
            public string getMultipleDimQuery { get; set; }



        }




    }
}
