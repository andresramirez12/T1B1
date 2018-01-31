using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace T1.B1.InformesTerceros
{
    public class Settings
    {
        public static Main _Main { get; }
        public static BalanceTerceros _BalanceTerceros { get; set; }
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

            _BalanceTerceros = new BalanceTerceros();
            _BalanceTerceros.Initialize();





        }

        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public Main()
            {
                logLevel = "Debug";
                





            }

            public string logLevel { get; }
           

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

        public class BalanceTerceros : Westwind.Utilities.Configuration.AppConfiguration
        {
            public BalanceTerceros()
            {
                reportType = "BALTER";
                getObjectControl = "select Code, Name, DocEntry, U_TransType, U_LastDate, U_LastTrans from[@BYB_T1ITR100] where U_LastDate < convert(datetime,'[--LastDate--]',112)";
                getTransactionList = "select distinct TransId from OJDT where TransId > coalesce((select distinct U_LastTrans from [@BYB_T1ITR100]),-1) order by TransId ASC";
                upsertObjectControl = "MERGE [@BYB_T1ITR100] as T0 "+
                    " using (select '" + reportType + "' as  [U_ReportType]) as S1 "+
                    " ( [U_ReportType]) " +
                    " on (T0.[U_ReportType] = S1.[U_ReportType]) " +
                    " when matched then update set T0.[U_LastTrans] = [--LastTrans--], T0.[U_LastDate] = getdate() " +
                    " when not matched then insert([Code], [Name], [DocEntry], [U_LastDate], [U_LastTrans], [U_ReportType] ) values(newid(), newid(), (coalesce((select max(DocEntry) from [@BYB_T1ITR100]),0)+1),getdate(),[--LastTrans--],'" + reportType + "');";
                UpsertCurrentBalance = "MERGE " +
                    " [@BYB_T1ITR101] as T0 " +
                    " using (select '[--Account--]' as [U_Account], [--Level--] as [U_Level]" +
                    " ,[--Year--] as [U_Year],[--Month--] as [U_Month],[--Day--] as [U_Day]" +
                    " ,'[--DateType--]' as [U_DateType],'[--ThirdParty--]' as [U_ThirdParty],'[--DIM1--]' as [U_DIM1]" +
                    " ,'[--DIM2--]' as [U_DIM2],'[--DIM3--]' as [U_DIM3],'[--DIM4--]' as [U_DIM4],'[--DIM5--]' as [U_DIM5]) as S1" +
                    " ( [U_Account], [U_Level], [U_Year], [U_Month], [U_Day], [U_DateType], [U_ThirdParty]," +
                    " [U_DIM1], [U_DIM2], [U_DIM3], [U_DIM4], [U_DIM5]) ON" +
                    " (T0.[U_Account]=S1.[U_Account] and T0.[U_Level] = S1.[U_Level] and" +
                    " T0.[U_Year]=S1.[U_Year] and T0.[U_Month]=S1.[U_Month] and T0.[U_Day]=S1.[U_Day]" +
                    " and T0.[U_DateType] =S1.[U_DateType] and T0.[U_ThirdParty] =S1.[U_ThirdParty]" +
                    " and T0.[U_DIM1] =S1.[U_DIM1] and T0.[U_DIM2] =S1.[U_DIM2]" +
                    " and T0.[U_DIM3]=S1.[U_DIM3] and T0.[U_DIM4] =S1.[U_DIM4]" +
                    " and T0.[U_DIM5] =S1.[U_DIM5]) " +
                    " when matched then update set T0.[U_Debit] += [--Debit--]," +
                    " T0.[U_Credit] += [--Credit--], T0.[U_CurrBalance] += [--CurrBalance--] " +
                    " when not matched then insert( [Code], [Name], [DocEntry], [U_Account], [U_Level], [U_Year], [U_Month]," +
                    " [U_Day], [U_DateType], [U_CurrBalance], [U_ThirdParty], [U_TransCode], [U_Debit]," +
                    " [U_Credit], [U_DIM1], [U_DIM2], [U_DIM3], [U_DIM4], [U_DIM5] )" +
                    " values (newid(),newid(),(coalesce((select max(DocEntry) from [@BYB_T1ITR101]),0)+1),S1.[U_Account]" +
                    " ,S1.[U_Level],S1.[U_Year]" +
                    " ,S1.[U_Month],S1.[U_Day],S1.[U_DateType],[--CurrBalance--],'[--ThirdParty--]','[--TransCode--]',[--Debit--],[--Credit--]" +
                    " ,S1.[U_DIM1],S1.[U_DIM2],S1.[U_DIM3],S1.[U_DIM4],S1.[U_DIM5]);";
            }
            public string upsertObjectControl { get; set; }
            public string getObjectControl { get; set; }
            public string getTransactionList { get; set; }
            public string reportType { get; set; }
            public string UpsertCurrentBalance { get; set; }


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
