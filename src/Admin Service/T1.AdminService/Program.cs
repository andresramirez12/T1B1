using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using Nancy;
using Nancy.Hosting.Self;
using System.Configuration.Install;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;
using log4net;
using log4net.Config;

namespace T1.AdminService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        static void Main(string[] args)
        {
            try
            {
                if (Environment.UserInteractive)
                {
                    string parameter = string.Concat(args);
                    switch (parameter)
                    {
                        case "--install":

                            ManagedInstallerClass.InstallHelper(new[] { Assembly.GetExecutingAssembly().Location });
                            break;
                        case "--uninstall":
                            ManagedInstallerClass.InstallHelper(new[] { "/u", Assembly.GetExecutingAssembly().Location });
                            break;
                        default:
                            T1Service objT1Service = new T1Service();
                            objT1Service.startNancy();
                            System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
                            break;
                    }
                }
                else
                {
                    ServiceBase[] servicesToRun = new ServiceBase[]
                                      {
                              new T1Service()
                                      };
                    ServiceBase.Run(servicesToRun);
                }
            }
            catch (Exception er)
            {
                _Logger.Fatal("", er);
            }
        }
    }
}
