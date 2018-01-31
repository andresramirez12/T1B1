using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Runtime.Remoting;
using log4net;
using Quartz;
using Quartz.Impl;

namespace T1.B1.ObjectReader
{
    public class ObjectReader
    {
        static private ObjectReader _ObjectReader = null;
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);
        private static List<T1.Shared.Classes.ObjectCron> objList = null;

        private ObjectReader()
        {
            try
            {
                objList = T1.DBManager.Instance.getObjectCronList();
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }
        }

        static public void loadObjects()
        {
            if (_ObjectReader == null)
            {
                _ObjectReader = new ObjectReader();
            }

            try
            {
                if (objList != null)
                {
                    foreach (T1.Shared.Classes.ObjectCron objCron in objList)
                    {
                        LoadStartUpJob(objCron);
                    }
                }
                else
                {
                    _Logger.Error("No object information found");
                }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }



        }

        static private void LoadStartUpJob(T1.Shared.Classes.ObjectCron objCron)
        {
            try
            {
                
                IJobDetail _Job = JobBuilder.Create<StartUpJob>()
                    .WithIdentity(objCron.JobId, objCron.GroupId)
                    .Build();

                ITrigger _Trigger = TriggerBuilder.Create()
                    .WithIdentity(objCron.TriggerId, objCron.GroupId)
                    .StartNow()
                    .WithCronSchedule(objCron.Cron)
                    .Build();

                if (!T1.Cron.isJobRegistered(_Job.Key))
                {

                    T1.Cron.addJob(_Job, _Trigger);
                    _Logger.Debug("B1Connection cron job added. Refreshing based on cron string " + objCron.Cron);
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
                if (_ObjectReader == null)
                {
                    _ObjectReader = new ObjectReader();
                }

                try
                {
                    if (objList != null)
                    {
                        foreach (T1.Shared.Classes.ObjectCron objCron in objList)
                        {
                            if (T1.Cron.getTriggerStatus(objCron.TriggerId, objCron.GroupId) == TriggerState.Normal)
                            {
                                T1.Cron.pauseTrigger(objCron.TriggerId, objCron.GroupId);
                                //Implement Background Worker in each call.
                            }
                        }
                    }
                    else
                    {
                        _Logger.Error("No object information found");
                    }

                }
                catch (Exception er)
                {
                    _Logger.Fatal("", er);


                }
            }
        }
    }
}
