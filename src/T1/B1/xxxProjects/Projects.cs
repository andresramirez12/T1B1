using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using T1.Classes;
using System.Runtime.InteropServices;

namespace T1.B1.Projects
{
    public class Projects
    {
        private static Projects objProjects = null;
        private static SAPbobsCOM.ProjectsService objProjectsService = null;
        
        private Projects()
        {
            SAPbobsCOM.CompanyService objCompanyService = null;
        
            try
            {
                objCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                objProjectsService = objCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService);
                
                
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        static private SAPbobsCOM.ProjectsParams getProjectList()
        {

            SAPbobsCOM.ProjectsParams objProjectList = null;

            try
            {
                objProjectList = objProjectsService.GetProjectList();
                return objProjectList;
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
                return null;
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
                return null;
            }
            

        }

        static public void fillValidValues(ref SAPbouiCOM.ComboBox oCombo)
        {
            SAPbobsCOM.ProjectsParams objProjectsParams = null;
            XmlDocument strXMLProjectList = null;
            
            try
            {
                if (objProjects == null)
                    objProjects = new Projects();

                strXMLProjectList = new XmlDocument();

                objProjectsParams = getProjectList();
                if(objProjectsParams != null)
                {
                    strXMLProjectList.LoadXml(objProjectsParams.ToXMLString());
                    if(oCombo.ValidValues.Count > 0)
                    {
                        while(oCombo.ValidValues.Count > 0)
                        {
                            oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }

                    }
                    else
                    {

                    }
                }

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException("", "ProductionUnits.MenuEvent", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        

    }
}
