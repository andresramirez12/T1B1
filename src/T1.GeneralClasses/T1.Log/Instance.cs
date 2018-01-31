using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Config;
using log4net.Repository.Hierarchy;
using log4net.Core;
using log4net.Appender;
using log4net.Layout;
using System.IO;

namespace T1.Log
{
    public class Instance
    {
        static private Instance _logg = null;
        private Instance()
        {
            string AppDataPath = AppDomain.CurrentDomain.BaseDirectory + "\\BYB\\T1";
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }


            Hierarchy hierarchy = (Hierarchy)LogManager.GetRepository();

            PatternLayout patternLayout = new PatternLayout();
            patternLayout.ConversionPattern = T1.Log.Settings._Main.pattern;
            patternLayout.ActivateOptions();

            RollingFileAppender roller = new RollingFileAppender();
            roller.AppendToFile = false;
            roller.File = AppDataPath + T1.Log.Settings._Main.logFolder + T1.Log.Settings._Main.masterLogName;
            roller.Layout = patternLayout;
            roller.MaxSizeRollBackups = 5;
            roller.MaximumFileSize = T1.Log.Settings._Main.masterSize;
            roller.RollingStyle = RollingFileAppender.RollingMode.Size;
            roller.AppendToFile = true;
            log4net.Filter.LevelRangeFilter filter = new log4net.Filter.LevelRangeFilter();
            filter.LevelMin = log4net.Core.Level.Error;
            filter.ActivateOptions();
            roller.AddFilter(filter);
            roller.ActivateOptions();
            hierarchy.Root.AddAppender(roller);
            hierarchy.Configured = true;
        }

        public static ILog GetLogger(Type tBaseType, string levelName)
        {
            if (_logg == null)
            {
                _logg = new Instance();
            }
            ILog logResult = null;
            string strTypeName = tBaseType.FullName;
            log4net.Appender.IAppender[] objAppenders = log4net.LogManager.GetRepository().GetAppenders();
            bool blFound = false;
            if (objAppenders.Length > 0)
            {
                for (int i = 0; i < objAppenders.Length; i++)
                {
                    log4net.Appender.IAppender objA = objAppenders[i];
                    if (strTypeName + T1.Log.Settings._Main.appenderSufix == objA.Name)
                    {
                        blFound = true;
                        break;
                    }
                }
            }
            if (blFound)
            {
                logResult = LogManager.GetLogger(strTypeName);
            }
            else
            {
                SetLevel(strTypeName, levelName);
                AddAppender(strTypeName, CreateFileAppender(strTypeName, strTypeName, levelName));
                logResult = LogManager.GetLogger(strTypeName);
            }
            return logResult;
        }

        public static log4net.Appender.IAppender CreateFileAppender(string name, string fileName, string levelName)
        {
            string AppDataPath = AppDomain.CurrentDomain.BaseDirectory + "\\BYB\\T1";
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }


            log4net.Appender.RollingFileAppender appender = new
            log4net.Appender.RollingFileAppender();
            appender.Name = name + T1.Log.Settings._Main.appenderSufix;
            appender.File = AppDataPath + T1.Log.Settings._Main.logFolder + fileName + ".log";
            appender.AppendToFile = true;
            appender.RollingStyle = log4net.Appender.RollingFileAppender.RollingMode.Size;
            appender.MaxSizeRollBackups = -1;
            appender.MaximumFileSize = T1.Log.Settings._Main.masterSize;
            appender.CountDirection = 1;

            log4net.Layout.PatternLayout layout = new
            log4net.Layout.PatternLayout();
            layout.ConversionPattern = T1.Log.Settings._Main.pattern;
            layout.ActivateOptions();

            log4net.Filter.LevelRangeFilter filter = new log4net.Filter.LevelRangeFilter();
            switch (levelName)
            {
                case "All":
                    filter.LevelMin = log4net.Core.Level.All;
                    break;
                case "Alert":
                    filter.LevelMin = log4net.Core.Level.Alert;
                    break;
                case "Debug":
                    filter.LevelMin = log4net.Core.Level.Debug;
                    break;
                case "Critical":
                    filter.LevelMin = log4net.Core.Level.Critical;
                    break;
                case "Error":
                    filter.LevelMin = log4net.Core.Level.Error;
                    break;
                case "Fatal":
                    filter.LevelMin = log4net.Core.Level.Fatal;
                    break;
                case "Info":
                    filter.LevelMin = log4net.Core.Level.Info;
                    break;
                case "Warn":
                    filter.LevelMin = log4net.Core.Level.Warn;
                    break;
                default:
                    filter.LevelMin = log4net.Core.Level.All;
                    break;

            }





            filter.ActivateOptions();

            appender.Layout = layout;
            appender.AddFilter(filter);

            appender.ActivateOptions();

            return appender;
        }

        public static void SetLevel(string loggerName, string levelName)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(loggerName);
            log4net.Repository.Hierarchy.Logger l = (log4net.Repository.Hierarchy.Logger)log.Logger;

            l.Level = l.Hierarchy.LevelMap[levelName];
        }

        // Add an appender to a logger
        public static void AddAppender(string loggerName, log4net.Appender.IAppender appender)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(loggerName);
            log4net.Repository.Hierarchy.Logger l = (log4net.Repository.Hierarchy.Logger)log.Logger;
            l.AddAppender(appender);
        }
    }
}
