using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Collections;
using System.Reflection;
using log4net;

namespace T1
{
    static class Program
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                _Logger.Debug("Adding ConnectionString to cache");
                T1.CacheManager.CacheManager.Instance.addToCache(T1.CacheManager.Settings._Main.connStringCacheName, (string)Environment.GetCommandLineArgs().GetValue(1), CacheManager.CacheManager.objCachePriority.NotRemovable);
                _Logger.Debug("Starting Connection to SAP Business One");
                T1.B1.Connection.Class objConnClass = new B1.Connection.Class();
                objConnClass.B1Connect(false);
                _Logger.Debug("Connection status " + objConnClass.Connected.ToString());
                if (objConnClass.Connected)
                {
                    _Logger.Debug("Starting Meta data creation.");
                    bool blMD = T1.B1.MetaData.Operations.blCreateMD(Settings._Main.createMD);
                    _Logger.Debug("Meta data creation ended.");
                    if (blMD)
                    {
                        #region loadFirstTimeData
                        if (Settings._Main.loadInitialData)
                        {
                            //Expand the logic of this methods in the future to create also all the transaction codes needed
                            T1.B1.MetaData.Operations.loadMuni();
                            T1.B1.MetaData.Operations.loadDepto();
                            T1.B1.MetaData.Operations.loadGenericUDO();
                        }
                        
                        #endregion

                        T1.B1.EventFilter.Operations objEventFilter = new B1.EventFilter.Operations();
                        T1.B1.EventManager.Operations objEventManager = new B1.EventManager.Operations();
                        

                        if (objEventManager.Status)
                        {
                            //Add Logic to menu that finds if the menu was created or not before. Find a quicker way to compare the existing menus instead of com object.
                            T1.B1.MenuManager.Operations.addMenu();
                            GC.KeepAlive(objEventManager);
                            GC.KeepAlive(objEventFilter);
                            Application.Run();
                        }
                        else
                        {
                            _Logger.Error("There was an error adding the EventListeners for the AddOn. Please check the log.");
                            T1.B1.MainObject.Instance.B1Application.SetStatusBarMessage("T1: There was an error adding the EventListeners for the AddOn. The execution will be halted.", SAPbouiCOM.BoMessageTime.bmt_Short);
                            Application.Exit();
                        }
                    }
                    else
                    {
                        _Logger.Error("There was an error creating the MetaData for the AddOn. Please check the log.");
                        T1.B1.MainObject.Instance.B1Application.SetStatusBarMessage("T1: An error was found creating the MD for the AddOn. The execution will be halted.", SAPbouiCOM.BoMessageTime.bmt_Short);
                        Application.Exit();
                    }

                    //Add Cache logic for Withholding taxes and WT Item Info. Also add cache for project and dimension lists. Add not also the object to cache but also a ready to use valid values to drop down. Run everything on background workers.
                    
                    //if (objFilterClass.Status && objEventManagerClass.Status)
                    //{
                    //    ///Load all Items WT configuration into memory
                    //    T1.B1.Helpers.WithHoldingTaxes.cacheAllItemsWTInfo("WTCodesByItem");
                    //    ///Load WT System Defintions
                    //    T1.B1.Helpers.WithHoldingTaxes.cacheAllWTDefinitions("WTCodesDefinitions");

                    //    T1.Classes.BYBMenuManager objMenuManager = new Classes.BYBMenuManager();
                    
                    //    GC.KeepAlive(objMenuManager);
                    //    Application.Run();
                    //}

                }
                else
                {
                    _Logger.Error("Could not connect to SAP Business One. T1 is terminating.");
                    Application.Exit();
                }


            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                
            }
            
            
        }
    }
}
