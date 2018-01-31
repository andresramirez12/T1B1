using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Quartz;
using Quartz.Impl;

namespace T1
{
    public class Cron
    {
        static private Cron _Cron = null;
        static private IScheduler _scheduler = null;


        private Cron()
        {

        }

        static public void StartCron()
        {
            if (_Cron == null)
            {
                _Cron = new Cron();
            }

            _scheduler = StdSchedulerFactory.GetDefaultScheduler();
            _scheduler.Start();
        }

        static public void StopCron()
        {
            _scheduler.Shutdown();
        }

        static public void addJob(IJobDetail Job, ITrigger Trigger)
        {
            _scheduler.ScheduleJob(Job, Trigger);
        }

        static public void pauseTrigger(string TriggerName, string TriggerGroup)
        {

            TriggerKey objTrigerKey = new TriggerKey(TriggerName, TriggerGroup);
            TriggerState objState = _scheduler.GetTriggerState(objTrigerKey);
            _scheduler.PauseTrigger(objTrigerKey);

        }

        static public void continueTrigger(string TriggerName, string TriggerGroup)
        {

            TriggerKey objTrigerKey = new TriggerKey(TriggerName, TriggerGroup);
            TriggerState objState = _scheduler.GetTriggerState(objTrigerKey);
            _scheduler.ResumeTrigger(objTrigerKey);

        }

        static public TriggerState getTriggerStatus(string TriggerName, string TriggerGroup)
        {

            TriggerKey objTrigerKey = new TriggerKey(TriggerName, TriggerGroup);
            return _scheduler.GetTriggerState(objTrigerKey);
        }

        static public bool isJobRegistered(JobKey objJobKey)
        {
            return _scheduler.CheckExists(objJobKey);
        }













    }
}
