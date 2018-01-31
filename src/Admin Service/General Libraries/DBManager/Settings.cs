using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace T1.DBManager
{
    public class Settings
    {
        public static Main _Main { get; set; }
        public static SQL _SQL { get; set; }
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

            _SQL = new SQL();
            _SQL.Initialize();

        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public Main()
            {
                logLevel = "Debug";
                connectionName = "T1";
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
            public string connectionName { get; }

        }

        public class SQL : Westwind.Utilities.Configuration.AppConfiguration
        {
            public SQL()
            {
                getObjectCronListQuery = "SELECT ID,OBJECTDEFID,CRON,JOBID,TRIGGERID,GROUPID, INSTANCE,LIBRARY FROM OBJECTCRON";
                getObjectControlQuery = "SELECT ID,OBJECTCRONID,LASTEXEC,LASTCORRECT,LASTEXECDATE,LASTCORRECTDATE FROM OBJECTCONTROL where OBJECTCRONID = [--ObjectCronId--]";
                
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

            public string getObjectCronListQuery { get; }
            public string getObjectControlQuery { get; }



        }
    }
}
