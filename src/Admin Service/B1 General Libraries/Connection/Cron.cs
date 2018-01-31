using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Quartz;
using Quartz.Impl;
using log4net;

namespace T1.B1.Connection
{
    public class Cron
    {
        static private Cron _cron = null;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private Cron()
        {

        }

        static public void loadCron()
        {
            if (_cron == null)
            {
                _cron = new Cron();
            }

            try
            {
                LoadConnections objLoadConn = new LoadConnections();
                objLoadConn.cacheConnections();
                LoadStartUpJob();
            }
            catch (Exception er)
            {
                _Logger.Fatal("", er);
            }




        }

        static private void LoadStartUpJob()
        {
            try
            {
                _Logger.Debug("Creating B1Connection cron job");
                IJobDetail _Job = JobBuilder.Create<StartUpJob>()
                    .WithIdentity(Settings._Main.jobId, Settings._Main.groupId)
                    .Build();

                ITrigger _Trigger = TriggerBuilder.Create()
                    .WithIdentity(Settings._Main.triggerId, Settings._Main.groupId)
                    .StartNow()
                    .WithCronSchedule(Settings._Main.cron)
                    .Build();

                if (!T1.Cron.isJobRegistered(_Job.Key))
                {

                    T1.Cron.addJob(_Job, _Trigger);
                    _Logger.Debug("B1Connection cron job added. Refreshing based on cron string " + Settings._Main.cron);
                }
                else
                {
                    _Logger.Debug("Job Already created. Ignoring creation");
                }
            }
            catch (Exception er)
            {
                _Logger.Fatal("", er);
            }

        }




        private class StartUpJob : IJob
        {
            public void Execute(IJobExecutionContext context)
            {
                try
                {

                    LoadConnections objLoadConn = new LoadConnections();
                    objLoadConn.cacheConnections();
                    if (Settings._Main.createMD)
                    {
                        if (T1.Cron.getTriggerStatus(Settings._Main.triggerId, Settings._Main.groupId) == TriggerState.Normal)
                        {
                            T1.Cron.pauseTrigger(Settings._Main.triggerId, Settings._Main.groupId);
                            //SuSo.B1.Base.MetaData.Instance objInstance = new B1.Base.MetaData.Instance();
                            //objInstance.CreateMD();
                            Settings._Main.createMD = false;
                            Settings._Main.Write();
                            T1.Cron.continueTrigger(Settings._Main.triggerId, Settings._Main.groupId);
                        }
                    }
                }
                catch (Exception er)
                {
                    _Logger.Fatal("", er);
                    

                }
                finally
                {
                    if (T1.Cron.getTriggerStatus(Settings._Main.triggerId, Settings._Main.groupId) != TriggerState.Normal)
                    {
                        T1.Cron.continueTrigger(Settings._Main.triggerId, Settings._Main.groupId);
                    }
                }

            }
        }




    }
}
