using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Config;
using System.Runtime.InteropServices;
using System.Xml;
using System.IO;
using Newtonsoft.Json;
using System.Data.Odbc;
using System;
using System.Collections.Generic;

namespace T1.B1.Connection
{
    class LoadConnections
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        public LoadConnections()
        {


        }

        public void cacheConnections()
        {
            string strErrorMessage = "";

            try
            {
                _Logger.Debug("Starting Connection confiuration reading.");
                if (Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + Settings._Main.connectionDirectory))
                {
                    string[] configFiles = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory + Settings._Main.connectionDirectory, "*.json");
                    if (configFiles.Length > 0)
                    {
                        _Logger.Debug(configFiles.Length.ToString() + " configuration files found");
                        Dictionary<string, string> conectionNameToIdDictionary = new Dictionary<string, string>();
                        Dictionary<string, ConfigurationInformation> conectionInformationDictionary = new Dictionary<string, ConfigurationInformation>();
                        for (int i = 0; i < configFiles.Length; i++)
                        {
                            string strJsonFile = "";
                            using (StreamReader sr = new StreamReader(configFiles[i]))
                            {
                                strJsonFile = sr.ReadToEnd();
                            }
                            if (strJsonFile.Length > 0)
                            {

                                ConfigurationInformation b1Configuration = JsonConvert.DeserializeObject<ConfigurationInformation>(strJsonFile);
                                _Logger.Debug("Connecting configuration " + i.ToString());
                                if (!conectionInformationDictionary.ContainsKey(b1Configuration.ConnectionId))
                                {
                                    conectionInformationDictionary.Add(b1Configuration.ConnectionId, b1Configuration);
                                }
                                for(int j=0; j < b1Configuration.Type.Length; j++)
                                {
                                    if(b1Configuration.Type[j]=="ODBC")
                                    {
                                        connectToDB(b1Configuration);
                                    }
                                    else if(b1Configuration.Type[j] == "B1")
                                    {
                                        connectToB1(b1Configuration);
                                    }
                                }
                              
                                _Logger.Debug("Adding configuration " + i.ToString() + " to Cache");
                                if (!conectionNameToIdDictionary.ContainsKey(b1Configuration.ConnectionId))
                                {
                                    conectionNameToIdDictionary.Add(b1Configuration.ConnectionName, b1Configuration.ConnectionId);
                                }
                                _Logger.Debug("Connection operation finished successfully form config file " + i.ToString());

                            }


                        }
                        T1.CacheManager.CacheManager.Instance.addToCache(Settings._Main.conectionNameToIdDictionary, conectionNameToIdDictionary, T1.CacheManager.CacheManager.objCachePriority.NotRemovable);
                        T1.CacheManager.CacheManager.Instance.addToCache(Settings._Main.conectionInformationDictionary, conectionInformationDictionary, T1.CacheManager.CacheManager.objCachePriority.NotRemovable);
                    }
                    else
                    {
                        _Logger.Fatal("No connection file found on configuration folder");

                    }
                }
                else
                {
                    _Logger.Fatal("Configuration folder does not exist. Please add the configuration information before running the service");
                }
            }
            catch (COMException comEx)
            {
                strErrorMessage = "COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace;
                _Logger.Error(strErrorMessage, comEx);
            }
            catch (Exception er)
            {
                _Logger.Error(er.Message, er);
            }
        }

        private void connectToB1(ConfigurationInformation b1ConnectionInfo)
        {
            SAPbobsCOM.Company objCompany = null;
            int intResult = -1;
            string strErrorMessage = "";
            bool blConnect = false;
            try
            {
                _Logger.Debug("Retreiving connection information from Cache: " + b1ConnectionInfo.ConnectionName);
                objCompany = T1.CacheManager.CacheManager.Instance.getFromCache(b1ConnectionInfo.ConnectionName);
                if (objCompany != null)
                {
                    if (!objCompany.Connected)
                    {
                        blConnect = true;
                        T1.CacheManager.CacheManager.Instance.removeFromCache(b1ConnectionInfo.ConnectionName);
                    }
                }
                else
                {
                    blConnect = true;
                }

                if (blConnect)
                {
                    _Logger.Debug("Starting to Connect to B1 Company");
                    objCompany = new SAPbobsCOM.Company();
                    objCompany.Server = b1ConnectionInfo.Server;
                    objCompany.LicenseServer = b1ConnectionInfo.LicenseServer;
                    objCompany.DbServerType = b1ConnectionInfo.B1DBServerType;
                    objCompany.CompanyDB = b1ConnectionInfo.CompanyDB;
                    objCompany.DbUserName = b1ConnectionInfo.DBUserName;
                    objCompany.DbPassword = b1ConnectionInfo.DBUserPassword;
                    objCompany.UserName = b1ConnectionInfo.B1UserName;
                    objCompany.Password = b1ConnectionInfo.B1Password;
                    intResult = objCompany.Connect();
                    if (intResult == 0)
                    {
                        _Logger.Debug("Conected to: " + objCompany.CompanyName);
                        T1.CacheManager.CacheManager.Instance.addToCache(b1ConnectionInfo.ConnectionName, objCompany, T1.CacheManager.CacheManager.objCachePriority.NotRemovable);
                    }
                    else
                    {
                        strErrorMessage = objCompany.GetLastErrorDescription();
                        _Logger.Error(strErrorMessage);
                    }
                }

            }
            catch (COMException comEx)
            {
                strErrorMessage = "COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace;
                _Logger.Error(strErrorMessage, comEx);
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

        }

        private void connectToDB(ConfigurationInformation b1ConnectionInfo)
        {

            bool blConnect = false;
            try
            {
                
                    _Logger.Debug("Retreiving connection information from Cache: " + b1ConnectionInfo.ConnectionId);


                    OdbcConnection objODBCCOnn = null;
                    objODBCCOnn = T1.CacheManager.CacheManager.Instance.getFromCache(b1ConnectionInfo.ConnectionId);
                    if (objODBCCOnn != null)
                    {
                        _Logger.Debug(b1ConnectionInfo.ConnectionId + "found. Checking Status");
                        if (objODBCCOnn.State != System.Data.ConnectionState.Open)
                        {
                            _Logger.Debug("Status " + objODBCCOnn.State.ToString() + " found. Starting connection mechanism");

                            blConnect = true;
                            T1.CacheManager.CacheManager.Instance.removeFromCache(b1ConnectionInfo.ConnectionId);
                        }
                    }
                    else
                    {
                        blConnect = true;
                    }

                    if (blConnect)
                    {
                        _Logger.Debug("Connecting configuration file " + b1ConnectionInfo.ConnectionId);
                        OdbcConnectionStringBuilder objODBC = new OdbcConnectionStringBuilder();

                    if (Settings._Main.isHANA)
                    {
                        objODBC.Driver = Settings._Main.HANADriver;
                        objODBC.Add("UID", b1ConnectionInfo.UserName);
                        objODBC.Add("PWD", b1ConnectionInfo.Password);
                        objODBC.Add("SERVERNODE", b1ConnectionInfo.Instance);
                    }
                    else
                    {
                        if(b1ConnectionInfo.B1DBServerType == SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008)
                        {
                            objODBC.Driver = "SQL Server Native Client 10.0";
                            objODBC.Add("MultipleActiveResultSets", "True");


                        }
                        else if (b1ConnectionInfo.B1DBServerType == SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012)
                        {
                            objODBC.Driver = "SQL Server Native Client 11.0";
                        }
                        else if (b1ConnectionInfo.B1DBServerType == SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014)
                        {
                            objODBC.Driver = "SQL Server Native Client 11.0";
                        }
                        
                        objODBC.Add("Uid", b1ConnectionInfo.UserName);
                        objODBC.Add("Pwd", b1ConnectionInfo.Password);
                        objODBC.Add("Server", b1ConnectionInfo.Instance);
                        objODBC.Add("Database", b1ConnectionInfo.DefaultSchema);
                        
                    }
                        objODBCCOnn = new OdbcConnection(objODBC.ConnectionString);
                    
                        objODBCCOnn.Open();
                        if (objODBCCOnn.State == System.Data.ConnectionState.Open)
                        {
                            _Logger.Debug("Adding connection id " + b1ConnectionInfo.ConnectionId + " to cache");

                            T1.CacheManager.CacheManager.Instance.addToCache(b1ConnectionInfo.ConnectionId, objODBCCOnn, T1.CacheManager.CacheManager.objCachePriority.NotRemovable);
                        }
                        else
                        {
                            _Logger.Error("Error connecting connection Id " + b1ConnectionInfo.ConnectionId);

                        }
                    }


                





            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }

        }

    }

    

}
