using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Config;
using Nancy;
using Nancy.Hosting.Self;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;

namespace T1.AdminService
{
    public partial class T1Service : ServiceBase
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        NancyHost host;
        BackgroundWorker nancyThread = null;
        public T1Service()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            startNancy();
        }

        protected override void OnStop()
        {
            stopNancy();
        }

        public void startNancy()
        {
            try
            {
                nancyThread = new BackgroundWorker();
                nancyThread.WorkerSupportsCancellation = true;
                nancyThread.WorkerReportsProgress = true;
                nancyThread.DoWork += NancyThread_DoWork;
                nancyThread.Disposed += NancyThread_Disposed;
                nancyThread.RunWorkerAsync();


            }
            catch (Exception er)
            {
                _Logger.Fatal("", er);
            }



        }

        public void stopNancy()
        {

            try
            {
                if (host != null)
                {

                    host.Stop();
                    if (nancyThread != null)
                    {
                        nancyThread.CancelAsync();
                        nancyThread.Dispose();
                    }
                    T1.Cron.StopCron();
                }
            }
            catch (Exception er)
            {
                _Logger.Fatal("", er);
            }

        }

        private void NancyThread_Disposed(object sender, EventArgs e)
        {
            if (host != null)
            {
                host.Stop();
                
            }
        }

        private void NancyThread_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                

                #region Automatic URL Registration
                var hostConfiguration = new HostConfiguration()
                {
                    UrlReservations = new UrlReservations() { CreateAutomatically = true }
                };


                #endregion
                InstallInfo.InstallInfo.Instance.createInstance(InstallInfo.InstallInfo.Instance.Config.T1Server.moduleListFile);

                BackgroundWorker worker = sender as BackgroundWorker;
                host = new NancyHost(hostConfiguration, new Uri(T1.InstallInfo.InstallInfo.Instance.Config.nancyLocalAddress));
                
                host.Start();
                T1.Cron.StartCron();
                T1.InstanceLoader.InstanceLoader.loadInstances();
            }
            catch (Exception er)
            {
                _Logger.Fatal("", er);
            }


        }
    }
}
