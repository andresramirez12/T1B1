using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace T1.B1.ReletadParties
{
    public class Settings
    {
        public static Main _Main { get; }
        public static SQL _SQL { get; }
        public static HANA _HANA { get; }
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
                lastFolderId = "1320002081";
                RelatedPartiesFolderId = "BYB_FLRP";
                BPFormTypeEx = "134";
                RelatedPartiesFolderPane = 20;
                EmptyDSMainThirdparties = "EmptyDSMainThirdparties";
                EmptyDSRelationThirdPartied = "EmptyDSRelationThirdPartied";
                BPFormMatrixId = "BYB_I47";
                BPFormBYBEditTextItems = "BYB_I51,BYB_I5,BYB_I7,BYB_I35,BYB_I15,BYB_I17,BYB_I19,BYB_I21,BYB_I31,BYB_I33,BYB_I29,BYB_I41,BYB_I11,BYB_I13,BYB_I39,BYB_I43,BYB_I37,BYB_I23,BYB_I25,BYB_I27";
                lastRightClickEventInfo = "lastRightClickEventInfoBP";
                relatedPartiesUDO = "BYB_T1RPA100";



            }

            public string logLevel { get; }
            public string lastFolderId { get; }
            public string RelatedPartiesFolderId { get; }
            public string BPFormTypeEx { get; }
            public int RelatedPartiesFolderPane { get; }

            public string EmptyDSMainThirdparties { get; }
            public string EmptyDSRelationThirdPartied { get; }
            public string BPFormBYBEditTextItems { get; set; }
            public string BPFormMatrixId { get; }
            public string lastRightClickEventInfo { get; }
            public string relatedPartiesUDO { get; }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }

        }

        public class SQL : Westwind.Utilities.Configuration.AppConfiguration
        {
            public SQL()
            {
                getCodeFromCardCode = "Select \"Code\" from [@BYB_T1RPA100] where \"U_CARDCODE\" = '[--CardCode--]'";
                getMissingRP = "select \"CardCode\" as \"Código\", \"CardName\" as \"Nombre\", \"LicTradNum\" as \"Identificador\" from OCRD where \"CardCode\" not in(select distinct U_CARDCODE from [@BYB_T1RPA100])";
                getBPCodeRelation = "Select \"Code\", \"U_CARDCODE\"  from [@BYB_T1RPA100]";
            }

            public string getCodeFromCardCode { get; }
            public string getMissingRP { get; }
            public string getBPCodeRelation { get; set; }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }



        }

        public class HANA : Westwind.Utilities.Configuration.AppConfiguration
        {
            public HANA()
            {
                getCodeFromCardCode = "Select \"Code\"   from \"@BYB_T1RPA100\" where \"U_CARDCODE\" = '[--CardCode--]'";
                getMissingRP = "select \"CardCode\"as \"Código\", \"CardName\" as \"Nombre\", \"LicTradNum\" as \"Identificador\" from OCRD where \"CardCode\" not in(select distinct U_CARDCODE from \"@BYB_T1RPA100\")";
                getBPCodeRelation = "Select \"Code\", \"U_CARDCODE\"  from [@BYB_T1RPA100]";
            }

            public string getCodeFromCardCode { get; }
            public string getMissingRP { get; }
            public string getBPCodeRelation { get; set; }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }



        }






    }
}
