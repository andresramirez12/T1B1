using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.Config
{
    public class Settings
    {

        public static T1B1Metadata _T1B1Metadata { get; set; }
        public static T1B1Connection _T1B1Connection { get; set; }
        public static T1CacheManager _T1CacheManager { get; set; }

        static Settings()
        {

            _T1B1Metadata = new T1B1Metadata();
            _T1B1Metadata.Initialize();

            _T1B1Connection = new T1B1Connection();
            _T1B1Connection.Initialize();

            _T1CacheManager = new T1CacheManager();
            _T1CacheManager.Initialize();


        }
    }

    public class T1B1Metadata : Westwind.Utilities.Configuration.AppConfiguration
    {
        public T1B1Metadata()
        {
            MDResourceDir = "Resources\\MDInfo.xml";
            updateUDOForm = true;
        }

        public string MDResourceDir { get; set; }
        public bool updateUDOForm { get; set; }
    }

    public class T1B1Connection : Westwind.Utilities.Configuration.AppConfiguration
    {
        public T1B1Connection()
        {
            ConnectionStringCacheName = "ConnectionString";
            useCompatibilityConnection = false;
            adminInfoCacheName = "admInfo";
            useCompanyApplication = true;




        }

        public string ConnectionStringCacheName { get; set; }
        public bool useCompatibilityConnection { get; set; }
        public string adminInfoCacheName { get; set; }
        public bool useCompanyApplication { get; set; }
        

    }

    public class T1CacheManager : Westwind.Utilities.Configuration.AppConfiguration
    {
        public T1CacheManager()
        {
            useAppDomain = true;
        }

        public bool useAppDomain { get; set; }
    }

    

}
