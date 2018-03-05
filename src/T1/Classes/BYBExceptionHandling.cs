using System;
using System.Diagnostics;
using System.Collections.Generic;
using Mindscape.Raygun4Net;
using System.Xml;
using System.Threading;
using System.Windows.Forms;
using System.IO;


namespace T1.Classes
{
    public class BYBExceptionHandling
    {
        static private BYBExceptionHandling objException = null;
        
        
        static private RaygunClient webExceptionHandler = null;

        static private Exception raygunException = null;
        

        private BYBExceptionHandling()
        {
            if (T1.Settings.ApplicationConfiguration.Default.useRaygun)
            {
                webExceptionHandler = new RaygunClient("+NUBMzSURXBariwzJrkfmw==");
            }
            
        }

        static public void reportException(string message, string messageLocation, Exception ex, int eventId, EventLogEntryType eventLogEntryType)
        {
            if (objException == null)
            { 
                objException = new BYBExceptionHandling();
            }

            string strLogPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + T1.Settings.ApplicationConfiguration.Default.logPath;
            string strFullLogFilePath = strLogPath + T1.Settings.ApplicationConfiguration.Default.logName;

            try
            {

                if (T1.Settings.ApplicationConfiguration.Default.useEventLog)
                {
                    using (EventLog eventLog = new EventLog())
                    {
                        if (!System.Diagnostics.EventLog.SourceExists(T1.Settings.ApplicationConfiguration.Default.eventSourceName))
                        {
                            System.Diagnostics.EventLog.CreateEventSource(T1.Settings.ApplicationConfiguration.Default.eventSourceName, T1.Settings.ApplicationConfiguration.Default.eventLogName);
                        }
                        eventLog.Source = T1.Settings.ApplicationConfiguration.Default.eventSourceName;

                        eventLog.WriteEntry(message + "::\r\n" + ex.StackTrace, eventLogEntryType, eventId);
                        eventLog.Close();
                    }
                }
                else
                {
                    if (!Directory.Exists(strLogPath))
                    {
                        Directory.CreateDirectory(strLogPath);
                    }
                    using (StreamWriter sr = new StreamWriter(string.Format(strFullLogFilePath, messageLocation, DateTime.Now.ToString("yyyyMMdd"))))
                    {
                        sr.Write(message + "::\r\n" + ex.StackTrace + eventLogEntryType.ToString() + eventId.ToString());
                    }
                }

                if (T1.Settings.ApplicationConfiguration.Default.useRaygun)
                {
                    raygunException = new Exception(ex.Message);
                    raygunException.Source = Convert.ToString(eventId);
                    Thread rungunThread = new Thread(sendRaygunException);
                    rungunThread.Start();
                }
                if(T1.Properties.Settings.Default.showB1Message)
                {
                    if(BYBB1MainObject.Instance.B1Company.Connected)
                    {
                        BYBB1MainObject.Instance.B1Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
            }
            catch(Exception er)
            {
                MessageBox.Show(er.Message);
            }

            
        }

        static private void sendRaygunException()
        {
            webExceptionHandler.Send(raygunException);
        }
    }
}