using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using log4net;
using Newtonsoft.Json;
using System.Runtime.Remoting;

namespace T1.InstallInfo
{
    public class InstallInfo
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private static readonly Lazy<InstallInfo> lazy =
           new Lazy<InstallInfo>(() => new InstallInfo());

        private static T1.Shared.Classes.InstallationInformation _config = null;
        private static bool _configFound = false;

        public T1.Shared.Classes.InstallationInformation Config
        {
            get
            {
                return _config;
            }
        }

        public bool configFound
        {
            get
            {
                return _configFound;
            }
        }

        public static InstallInfo Instance
        {
            get
            {
                return lazy.Value;
            }
        }

        private InstallInfo()
        {
            try
            {
                _Logger.Info("Searching for configuration file InstallInfo.json");
                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "config//InstallInfo.json"))
                {
                    _Logger.Info("File found. Reading file content");
                    string strJsonContent = "";
                    using (StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "config//InstallInfo.json"))
                    {
                        strJsonContent = sr.ReadToEnd();
                        sr.Close();
                    }
                    _Logger.Info("Read operation complete. Deserialization started");
                    if (strJsonContent.Trim().Length > 0)
                    {
                        _config = JsonConvert.DeserializeObject<T1.Shared.Classes.InstallationInformation>(strJsonContent);
                        _configFound = true;
                        _Logger.Info("Deserialization ended succesfully.");
                    }
                    else
                    {
                        _Logger.Error("The content of the file is invalid. The server will be terminated");
                    }


                }
                else
                {
                    _Logger.Error("No InstallInfo.json present. The server will be terminated");
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

        }

        public void createInstance(string FileName)
        {
            string strPath = "";
            string strFileContent = "";
            List<T1.Shared.Classes.NancyModuleLibraryInfo> objLibraryList = null;
            try
            {
                strPath = AppDomain.CurrentDomain.BaseDirectory + "config\\" + FileName;
                using (StreamReader sr = new StreamReader(strPath))
                {
                    strFileContent = sr.ReadToEnd();
                    sr.Close();
                }

                objLibraryList = JsonConvert.DeserializeObject<List<T1.Shared.Classes.NancyModuleLibraryInfo>>(strFileContent);
                if (objLibraryList != null && objLibraryList.Count > 0)
                {
                    foreach (T1.Shared.Classes.NancyModuleLibraryInfo oModule in objLibraryList)
                    {
                        try
                        {
                            ObjectHandle _handle = AppDomain.CurrentDomain.CreateInstance(oModule.ExternalLibrary, oModule.NameSpace + "." + oModule.Contructor);
                        }
                        catch (Exception er)
                        {
                            _Logger.Error("", er);
                        }

                    }
                }


                //var oTemp = JsonConvert.DeserializeObject




            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }
    }
}

