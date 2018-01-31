using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Runtime.Remoting;
using log4net;
using System.Reflection;

namespace T1.InstanceLoader
{
    public class InstanceLoader
    {
        static private InstanceLoader _InstanceLoader = null;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);


        private InstanceLoader()
        {

        }

        static public void loadInstances()
        {
            if (_InstanceLoader == null)
            {
                _InstanceLoader = new InstanceLoader();
            }

            try
            {
                _Logger.Debug("Starting library loading.");
                XmlDocument objDocument = new XmlDocument();
                string strFile = AppDomain.CurrentDomain.BaseDirectory + Settings._Main.configFile;
                objDocument.Load(strFile);
                if (objDocument != null)
                {
                    _Logger.Debug("Found " + objDocument.DocumentElement.ChildNodes.Count.ToString() + " libraries to load.");
                    foreach (XmlNode xn in objDocument.DocumentElement.ChildNodes)
                    {
                        if (xn.NodeType != XmlNodeType.Comment)
                        {
                            string strAssemblyName = xn.SelectSingleNode("assemblyName").InnerText;
                            string strClassString = xn.SelectSingleNode("classString").InnerText;
                            string strType = xn.SelectSingleNode("type").InnerText;
                            string strStaticMethod = xn.SelectSingleNode("staticMethod").InnerText;

                            

                            if (strType == "static")
                            {
                                Assembly oTest =  Assembly.Load(strAssemblyName);
                                Type type = null;
                                foreach (TypeInfo oType in oTest.DefinedTypes)
                                {
                                    
                                    if (strClassString == oType.Name)
                                    {
                                        type = Type.GetType(oType.AssemblyQualifiedName);
                                        break;
                                    }
                                }
                                if(type != null)
                                {
                                    MethodInfo info = type.GetMethod(strStaticMethod);
                                    info.Invoke(null, null);
                                }

                                //Type type = Type.GetType(oTest.FullName);
                                //Type type = Type.GetType(strAssemblyName + ", AssemblyName");
                                
                            }
                            else
                            {
                                ObjectHandle _handle = Activator.CreateInstance(strAssemblyName, strClassString);
                                _handle = null;
                                _Logger.Debug("Class " + strAssemblyName + " loaded");
                            }




                            

                            
                        }
                    }
                }
                else
                {
                    _Logger.Error("Configuration file not found for T1 Admin Service InstanceLoader");
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }



        }
    }
}
