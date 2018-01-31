using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Runtime.Remoting;
using log4net;
using System.Data;
using System.Data.Odbc;

namespace T1.DBManager
{
    public class Instance
    {
        static private Instance _Instance = null;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static OdbcConnection objConnection = null;

        private Instance()
        {
            if(objConnection == null)
            {
                getConnection();
            }
        }

        private static void getConnection()
        {
            try
            {
                string strConnectionName = Settings._Main.connectionName;
                Dictionary<string, string> conectionNameToIdDictionary = T1.CacheManager.CacheManager.Instance.getFromCache(T1.B1.Connection.Settings._Main.conectionNameToIdDictionary);
                foreach (string strName in conectionNameToIdDictionary.Keys)
                {
                    if (strName == strConnectionName)
                    {
                        objConnection = T1.CacheManager.CacheManager.Instance.getFromCache(conectionNameToIdDictionary[strName]);
                        break;
                    }
                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }

        public static List<T1.Shared.Classes.ObjectCron> getObjectCronList()
        {
            List<T1.Shared.Classes.ObjectCron> objResult = null;
            
            OdbcCommand objCommand = null;
            string strSql = "";
            try
            {
                if(_Instance == null)
                {
                    _Instance = new Instance();
                }
                if (objConnection == null)
                {
                    getConnection();
                }
                if (objConnection != null && objConnection.State == ConnectionState.Open)
                {
                    strSql = Settings._SQL.getObjectCronListQuery;

                    objCommand = new OdbcCommand();
                    objCommand.Connection = objConnection;
                    objCommand.CommandText = strSql;
                    using (OdbcDataReader oDR = objCommand.ExecuteReader())
                    {
                        if (oDR.HasRows)
                        {
                            objResult = new List<Shared.Classes.ObjectCron>();
                            while (oDR.Read())
                            {
                                T1.Shared.Classes.ObjectCron objCron = new Shared.Classes.ObjectCron();
                                objCron.Id = !Convert.IsDBNull(oDR["ID"]) ? (int)oDR["ID"] : -1;
                                objCron.JobId = !Convert.IsDBNull(oDR["JOBID"]) ? (string)oDR["JOBID"] : "";
                                objCron.GroupId = !Convert.IsDBNull(oDR["GROUPID"]) ? (string)oDR["GROUPID"] : "";
                                objCron.ObjectDefId = !Convert.IsDBNull(oDR["OBJECTDEFID"]) ? (int)oDR["OBJECTDEFID"] : -1;
                                objCron.Cron = !Convert.IsDBNull(oDR["CRON"]) ? (string)oDR["CRON"] : "";
                                objCron.TriggerId = !Convert.IsDBNull(oDR["TRIGGERID"]) ? (string)oDR["TRIGGERID"] : "";
                                objCron.Instance = !Convert.IsDBNull(oDR["INSTANCE"]) ? (string)oDR["INSTANCE"] : "";
                                objCron.Library = !Convert.IsDBNull(oDR["LIBRARY"]) ? (string)oDR["LIBRARY"] : "";
                                objResult.Add(objCron);
                            }
                        }
                        oDR.Close();
                    }
                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
                objResult = new List<Shared.Classes.ObjectCron>(); ;
            }
            return objResult;
        }

        public static bool upsertObjectCron(T1.Shared.Classes.ObjectCron objCron, DateTime oDate, string Key,string DateTimeColumn, string KeyColumn)
        {
            bool blResult = false;
            DataTable objDT = null;
            OdbcDataAdapter objAdapter = null;
            DataRow objRow = null;
            string strSql = "";
            try
            {
                if (_Instance == null)
                {
                    _Instance = new Instance();
                }
                if (objConnection == null)
                {
                    getConnection();
                }
                if (objConnection != null && objConnection.State == ConnectionState.Open)
                {
                    strSql = Settings._SQL.getObjectControlQuery.Replace("[--ObjectCronId--]", objCron.Id.ToString());
                    objAdapter = new OdbcDataAdapter(strSql, objConnection);
                    objAdapter.Fill(objDT);
                    if(objDT.Rows.Count > 0)
                    {
                        objDT.Rows[0][KeyColumn] = Key;
                        objDT.Rows[0][DateTimeColumn] = oDate;

                    }
                    else
                    {
                        objRow = objDT.NewRow();
                        objRow["OBJECTCRONID"] = objCron.Id;
                        objRow[KeyColumn] = Key;
                        objRow[DateTimeColumn] = oDate;
                        objDT.Rows.Add(objRow);


                    }
                    objAdapter.Update(objDT);
                    blResult = true;


                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
            return blResult;
        }

    }
}
