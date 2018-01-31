using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using log4net;
using Newtonsoft.Json;

namespace T1.B1.Base.InstallInfo
{
    public class InstallInfo
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, T1.B1.Base.InstallInfo.Settings._Main.logLevel);
        private static InstallInfo objInstallInfo = null;
       

        private static T1.B1.Base.InstallInfo.InstallationInformation _config = null;
        

        public static T1.B1.Base.InstallInfo.InstallationInformation Config
        {
            get
            {
                if(objInstallInfo == null)
                {
                    objInstallInfo = new InstallInfo();
                }
                return _config;
            }
        }

        


        

        private InstallInfo()
        {
            try
            {
                _config = new Base.InstallInfo.InstallationInformation();
                _config.configurationBaseFolder = T1.B1.Base.InstallInfo.Settings._Main.configurationBaseFolder;
                _config.debugLevel = T1.B1.Base.InstallInfo.Settings._Main.logLevel;
                _config.nancyLocalAddress = T1.B1.Base.InstallInfo.Settings._Main.nancyLocalAddress;
                
            }
            catch(Exception er)
            {
                
                _Logger.Error("", er);
            }

        }
    }
}
