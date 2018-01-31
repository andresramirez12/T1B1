using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Quartz;
using Quartz.Impl;
using log4net;
using log4net.Config;

namespace T1.B1.Connection
{
    public class Instance
    {

        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType,Settings._Main.logLevel);
        public Instance()
        {
            try
            {

                T1.B1.Connection.Cron.loadCron();
            }
            catch (Exception er)
            {
                _Logger.Fatal("", er);
            }
        }

    }
}
