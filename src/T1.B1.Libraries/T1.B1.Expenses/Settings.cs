using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Westwind.Utilities.Configuration;

namespace T1.B1.Expenses
{
    public class Settings
    {
        public static Main _Main { get; }
        public static MainPettyCash _MainPettyCash { get; set; }

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

            _MainPettyCash = new MainPettyCash();
            _MainPettyCash.Initialize();






        }

        public class MainPettyCash : Westwind.Utilities.Configuration.AppConfiguration
        {
            public MainPettyCash()
            {
                logLevel = "Debug";
                pettyCashExpenseType = "CM";
                pettyCashPaymentFormType = "BYB_T1PTC001";
                pettyCashUDO = "BYB_T1PTC100";
                pettyCashLegalizationUDO = "BYB_T1PTC300";
                pettyCashConceptUDO = "BYB_T1PTC200";
                pettyCashConceptUDOFormType = "UDO_FT_BYB_T1PTC200";
                pettyCashLegalizationFormType = "UDO_FT_BYB_T1PTC300";
                PCLegalizationFormLastId = "PCLegalizationFormLastId";
                PCLegalizationTransactionCode = "T1A2";
                SysCurrDeviationAccount = "429581005";
            }

            public string logLevel { get; set; }
            public string pettyCashExpenseType { get; set; }
            public string pettyCashPaymentFormType { get; set; }
            public string pettyCashUDO { get; set; }
            public string pettyCashLegalizationUDO { get; set; }
            public string pettyCashConceptUDO { get; set; }
            public string pettyCashConceptUDOFormType { get; set; }
            public string pettyCashLegalizationFormType { get; set; }
            public string PCLegalizationFormLastId { get; set; }
            public string PCLegalizationTransactionCode { get; set; }
            public string SysCurrDeviationAccount { get; set; }


            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<MainPettyCash>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }

        }


        public class Main : Westwind.Utilities.Configuration.AppConfiguration
        {
            public Main()
            {
                logLevel = "Debug";
                recreateMenu = false;
                ConceptUDOFormType = "UDO_FT_BYB_T1EXP100";
                ExpenseTypeUDOFormType = "UDO_FT_BYB_T1EXP200";
                ExpenseRequestUDoFormType = "UDO_FT_BYB_T1EXP600";
                ExpenseRequestFormLastId = "ExpenseRequestFormLastId";
                DefaultCreateStatus = "BORRADOR";
                ExpenseTypeClasificatioonUDOFormType = "BYB_T1EXP500";
                ApprovedStatusValue = "APROBADA";
                PaymentValidValues = "'APROBADA'";
                PaymentFormType = "BYB_EXPPAY001";
                ExpenseTypeUDO = "BYB_T1EXP200";
                ExpenseUDO = "BYB_T1EXP600";
                ExpenseRequestRelatedPartiesChild = "BYB_T1EXP601";
                RelatedPartyUDO = "BYB_T1RPA100";
                PaymentStatusValue = "DESEMBOLSADA";
                LegalizationFormLastId = "LegalizationFormLastId";
                LegalizationRequestUDoFormType = "UDO_FT_BYB_T1EXP400";
                ConceptUDO = "BYB_T1EXP100";
                LegalizationUDO = "BYB_T1EXP400";
                LegalizationTransactionCode = "T1A1";
                LastExpenseActiveForm = "LastExpenseActiveForm";
            }

            public string logLevel { get; }
            public bool recreateMenu { get; }
            public string ConceptUDOFormType { get; }
            public string ExpenseTypeUDOFormType { get; }
            public string ExpenseRequestUDoFormType { get; }
            public string ExpenseRequestFormLastId { get; }
            public string DefaultCreateStatus { get; }
            public string ExpenseTypeClasificatioonUDOFormType { get; }
            public string ApprovedStatusValue { get; set; }
            public string PaymentValidValues { get; set; }
            public string PaymentFormType { get; set; }
            public string ExpenseTypeUDO { get; set; }
            public string ExpenseUDO { get; set; }
            public string ExpenseRequestRelatedPartiesChild { get; set; }
            public string RelatedPartyUDO { get; set; }
            public string PaymentStatusValue { get; set; }
            public string LegalizationFormLastId { get; set; }
            public string LegalizationRequestUDoFormType { get; set; }
            public string ConceptUDO { get; set; }
            public string LegalizationUDO { get; set; }
            public string LegalizationTransactionCode { get; set; }
            public string LastExpenseActiveForm { get; set; }




            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }

            protected override void OnInitialize(IConfigurationProvider provider, string sectionName, object configData)
            {
                var JSONProvider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = JSONProvider;
                base.OnInitialize(JSONProvider, sectionName, configData);
            }

        }


        public class SQL : Westwind.Utilities.Configuration.AppConfiguration
        {
            public SQL()
            {
                getMultipleDimQuery = "Select \"UseMltDims\" from OADM";
                getApprovedRequests = "select 'N' as \"Aprobar\" ,\"DocEntry\" as \"Documento\",\"U_STARTDATE\" as \"Inicio\",U_ENDDATE as \"Finalización\",\"U_VALUE\" as \"Valor\" from [@BYB_T1EXP600] where \"U_STATUS\" = 'DEFINITIVA'";
                getValidPaymentDocuments = "select 'N' as \"Aprobar\" ,\"DocEntry\" as \"Documento\",\"U_STARTDATE\" as \"Inicio\",U_ENDDATE as \"Finalización\",\"U_VALUE\" as \"Valor\", \"U_PAYDATE\" as \"Fecha Pago\", \"U_PAYMENT\" as \"Comprobante de Pago\", \"U_STATUS\" as \"Estado\" from [@BYB_T1EXP600] where \"U_STATUS\" in ([--ValidStatus--])";
                getJEFromPayment = "Select \"TransId\" from OVPM where \"DocEntry\" = [--DocEntry--]";
                getConceptsCode = "select \"Code\" from [@BYB_T1EXP104] where U_EXPTYPE = '[--ExpType--]'";
            }

            public string getMultipleDimQuery { get; set; }
            public string getApprovedRequests { get; set; }
            public string getValidPaymentDocuments { get; set; }
            public string getJEFromPayment { get; set; }
            public string getConceptsCode { get; set; }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }
            protected override void OnInitialize(IConfigurationProvider provider, string sectionName, object configData)
            {
                var JSONProvider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = JSONProvider;
                base.OnInitialize(JSONProvider, sectionName, configData);
            }

        }

        public class HANA : Westwind.Utilities.Configuration.AppConfiguration
        {
            public HANA()
            {
                getMultipleDimQuery = "Select \"UseMltDims\" from OADM";
                getApprovedRequests = "select 'N' as \"Aprobar\" ,\"DocEntry\" as \"Documento\",\"U_STARTDATE\" as \"Inicio\",U_ENDDATE as \"Finalización\",\"U_VALUE\" as \"Valor\" from \"@BYB_T1EXP600\" where \"U_STATUS\" = 'DEFINITIVA'";
                getValidPaymentDocuments = "select 'N' as \"Desembolsar\" ,\"DocEntry\" as \"Documento\",\"U_STARTDATE\" as \"Inicio\",U_ENDDATE as \"Finalización\",\"U_VALUE\" as \"Valor\", \"U_STATUS\" as \"Estado\" from [@BYB_T1EXP600] where \"U_STATUS\" in ([--ValidStatus--])";
                getJEFromPayment = "Select \"TransId\" from OVPM where \"DocEntry\" = [--DocEntry--]";
                getConceptsCode = "select \"Code\" from \"@BYB_T1EXP104\" where U_EXPTYPE = '[--ExpType--]'";
            }

            public string getMultipleDimQuery { get; }
            public string getApprovedRequests { get; set; }
            public string getValidPaymentDocuments { get; set; }
            public string getJEFromPayment { get; set; }
            public string getConceptsCode { get; set; }

            protected override Westwind.Utilities.Configuration.IConfigurationProvider OnCreateDefaultProvider(string sectionName, object configData)
            {
                var provider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = provider;

                return provider;
            }

            protected override void OnInitialize(IConfigurationProvider provider, string sectionName, object configData)
            {
                var JSONProvider = new Westwind.Utilities.Configuration.JsonFileConfigurationProvider<Main>()
                {
                    JsonConfigurationFile = Settings.AppDataPath + this.GetType().FullName.Replace("+", ".") + ".json"
                };
                Provider = JSONProvider;
                base.OnInitialize(JSONProvider, sectionName, configData);
            }
        }




    }
}
