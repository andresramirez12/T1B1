using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Runtime.InteropServices;
using System.IO;

namespace T1.Classes
{

    public enum Status
    {
        installed = 1,
        failed = 2,
        pending = 3
    }

    public class ModuleRegistrationClass
    {
        public string Code { get; set; }
        
        private XmlDocument oXMLRegistration = new XmlDocument();

        public XmlDocument ModuleInformation
        {
            get
            {
                return oXMLRegistration;
            }
        }

        public ModuleRegistrationClass()
        {
            
                XmlNode docNode = oXMLRegistration.CreateXmlDeclaration("1.0", "UTF-8", null);
                oXMLRegistration.AppendChild(docNode);
                
                    XmlNode moduleNode = oXMLRegistration.CreateElement("Module");
                    oXMLRegistration.AppendChild(moduleNode);
                    
                
        }
        public bool AddModuleInformation(string Code, string Name, string Version)
        {
            bool blResult = false;
            

            XmlNode adminfoNode = oXMLRegistration.CreateElement("AdminInfo");
            
            XmlNode CodeNode = oXMLRegistration.CreateElement("Code");
            CodeNode.InnerText = Code;
            this.Code = Code;

            XmlNode NameNode = oXMLRegistration.CreateElement("Name");
            NameNode.InnerText = Name;

            XmlNode VersionNode = oXMLRegistration.CreateElement("Version");
            VersionNode.InnerText = Version;

            adminfoNode.AppendChild(CodeNode);
            adminfoNode.AppendChild(NameNode);
            adminfoNode.AppendChild(VersionNode);


            XmlNode ModuleNode = oXMLRegistration.SelectSingleNode("/Module");
            ModuleNode.AppendChild(adminfoNode);

            return blResult;
        }

        public bool AddMDInformation(string Type, string Name, Status Status)
        {
            bool blResult = false;
            bool blAddInstallationNode = false;
            bool blAddUDTNode = false;
            bool blAddUDFNode = false;
            bool blAddUDONode = false;

            XmlAttribute nameAttribute = oXMLRegistration.CreateAttribute("name");
            nameAttribute.Value = Name;
            XmlAttribute statusAttribute = oXMLRegistration.CreateAttribute("status");
            statusAttribute.Value = Status.ToString();
            XmlNode ItemNode = oXMLRegistration.CreateElement("Item");
            ItemNode.Attributes.Append(nameAttribute);
            ItemNode.Attributes.Append(statusAttribute);


            XmlNode installationNode = oXMLRegistration.SelectSingleNode("/Module/Installation");
            if (installationNode == null)
            {
                installationNode = oXMLRegistration.CreateElement("Installation");
                blAddInstallationNode = true;

            }

            switch(Type)
            {
                case "UDT":
                    XmlNode udtNode = installationNode.SelectSingleNode("./UDT");
                    if (udtNode == null)
                    {
                        udtNode = oXMLRegistration.CreateElement("UDT");
                        blAddUDTNode = true;
                    }
                    udtNode.AppendChild(ItemNode);
                    if (blAddUDTNode)
                        installationNode.AppendChild(udtNode);
                    break;
                case "UDF":
                    XmlNode udfNode = installationNode.SelectSingleNode("./UDF");
                    if (udfNode == null)
                    {
                        udfNode = oXMLRegistration.CreateElement("UDF");
                        blAddUDFNode = true;
                    }
                    udfNode.AppendChild(ItemNode);
                    if (blAddUDFNode)
                        installationNode.AppendChild(udfNode);
                    break;
                case "UDO":
                    XmlNode udoNode = installationNode.SelectSingleNode("./UDO");
                    if (udoNode == null)
                    {
                        udfNode = oXMLRegistration.CreateElement("UDO");
                        blAddUDONode = true;
                    }
                    udoNode.AppendChild(ItemNode);
                    if (blAddUDONode)
                        installationNode.AppendChild(udoNode);
                    break;
            }

            if (blAddInstallationNode)
            {
                XmlNode moduleNode = oXMLRegistration.SelectSingleNode("/Module");
                moduleNode.AppendChild(installationNode);

            }
            blResult = true;

            return blResult;
        }

        public bool AddReportTypeInformation(string TypeName, string TypeCode, string AddOnFormType, Status Status)
        {
            bool blResult = false;

            
            bool blAddInstallationNode = false;
            bool blAddReportTypeNode = false;
            

            XmlAttribute nameAttribute = oXMLRegistration.CreateAttribute("name");
            nameAttribute.Value = TypeName;
            XmlAttribute statusAttribute = oXMLRegistration.CreateAttribute("status");
            statusAttribute.Value = Status.ToString();
            XmlAttribute codeAttribute = oXMLRegistration.CreateAttribute("code");
            codeAttribute.Value = TypeCode;
            XmlAttribute formIdAttribute = oXMLRegistration.CreateAttribute("formId");
            formIdAttribute.Value = AddOnFormType;

            XmlNode ItemNode = oXMLRegistration.CreateElement("Item");
            ItemNode.Attributes.Append(nameAttribute);
            ItemNode.Attributes.Append(statusAttribute);
            ItemNode.Attributes.Append(codeAttribute);
            ItemNode.Attributes.Append(formIdAttribute);


            XmlNode installationNode = oXMLRegistration.SelectSingleNode("/Module/Installation");
            if (installationNode == null)
            {
                installationNode = oXMLRegistration.CreateElement("Installation");
                blAddInstallationNode = true;

            }

            
                    XmlNode reportTypeNode = installationNode.SelectSingleNode("./ReportType");
                    if (reportTypeNode == null)
                    {
                        reportTypeNode = oXMLRegistration.CreateElement("ReportType");
                        blAddReportTypeNode = true;
                    }
                    reportTypeNode.AppendChild(ItemNode);
                    if (blAddReportTypeNode)
                        installationNode.AppendChild(reportTypeNode);
                    

            if (blAddInstallationNode)
            {
                XmlNode moduleNode = oXMLRegistration.SelectSingleNode("/Module");
                moduleNode.AppendChild(installationNode);

            }

            
            blResult = true;

            return blResult;


        }

        public bool AddReportInformation(string TypeCode, string ReportName, string ReportCode, Status Status)
        {
            bool blResult = false;


            bool blAddInstallationNode = false;
            bool blAddReportNode = false;


            XmlAttribute nameAttribute = oXMLRegistration.CreateAttribute("name");
            nameAttribute.Value = ReportName;
            XmlAttribute statusAttribute = oXMLRegistration.CreateAttribute("status");
            statusAttribute.Value = Status.ToString();
            XmlAttribute strTypeCodeAttribute = oXMLRegistration.CreateAttribute("typeCode");
            strTypeCodeAttribute.Value = TypeCode;
            XmlAttribute codeAttribute = oXMLRegistration.CreateAttribute("code");
            codeAttribute.Value = ReportCode;

            XmlNode ItemNode = oXMLRegistration.CreateElement("Item");
            ItemNode.Attributes.Append(nameAttribute);
            ItemNode.Attributes.Append(statusAttribute);
            ItemNode.Attributes.Append(codeAttribute);
            ItemNode.Attributes.Append(strTypeCodeAttribute);


            XmlNode installationNode = oXMLRegistration.SelectSingleNode("/Module/Installation");
            if (installationNode == null)
            {
                installationNode = oXMLRegistration.CreateElement("Installation");
                blAddInstallationNode = true;

            }


            XmlNode reportNode = installationNode.SelectSingleNode("./Report");
            if (reportNode == null)
            {
                reportNode = oXMLRegistration.CreateElement("Report");
                blAddReportNode = true;
            }
            reportNode.AppendChild(ItemNode);
            if (blAddReportNode)
                installationNode.AppendChild(reportNode);


            if (blAddInstallationNode)
            {
                XmlNode moduleNode = oXMLRegistration.SelectSingleNode("/Module");
                moduleNode.AppendChild(installationNode);

            }


            blResult = true;

            return blResult;


        }

        public string AddReportType(string ModuleCode, string TypeName, string AddOnName, string AddOnFormType)
        {
            
            string strTypeCode = "";
            XmlDocument oCacheInformation = null;
            SAPbobsCOM.ReportTypesService rptTypeService = null;
            SAPbobsCOM.ReportType oReport = null;
            SAPbobsCOM.ReportTypeParams newTypeParam = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCacheInformation = BYBCache.Instance.getFromCache(T1.Properties.Settings.Default.VersionControlCacheName);
                if(oCacheInformation != null)
                {
                    XmlNode oReportTypeNode = oCacheInformation.SelectSingleNode("/VersionControl/Module[./AdminInfo/Code/text() = '" + ModuleCode + "']/Installation/ReportType/Item[@name='" + TypeName + "']");
                    bool blCreateType = false;
                    if(oReportTypeNode != null)
                    {
                        if (oReportTypeNode.Attributes["status"].Value == Status.installed.ToString())
                        {
                            strTypeCode = oReportTypeNode.Attributes["code"].Value;
                            AddReportTypeInformation(TypeName, strTypeCode, AddOnFormType, Status.installed);
                        }
                        else
                        {
                            blCreateType = true;
                        }
                    }
                    else
                    {
                        blCreateType = true;
                    }


                    if (blCreateType)
                    {
                        oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                        rptTypeService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
                        oReport = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType);

                        oReport.TypeName = TypeName;
                        oReport.AddonName = AddOnName;
                        oReport.AddonFormType = AddOnFormType;
                        newTypeParam = rptTypeService.AddReportType(oReport);
                        strTypeCode = newTypeParam.TypeCode;
                        AddReportTypeInformation(TypeName, strTypeCode, AddOnFormType, Status.installed);
                    }
                }
                else
                {
                    Exception er = new Exception(Convert.ToString("SAPIFRS Error::No Config File Found in cache"));
                    BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.AddReportType", er, 1, System.Diagnostics.EventLogEntryType.Error);
                    AddReportTypeInformation(TypeName, strTypeCode, AddOnFormType,Status.failed);
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.AddReportType", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.AddReportType", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

            return strTypeCode;


        }

        public string AddReport(string ModuleCode, string Author, string Name, string TypeCode, string FileName)
        {

            string strReportCode = "";
            XmlDocument oCacheInformation = null;
            //SAPbobsCOM.ReportTypesService rptTypeService = null;
            //SAPbobsCOM.ReportType oReport = null;
            //SAPbobsCOM.ReportTypeParams newTypeParam = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            
            
            SAPbobsCOM.ReportLayoutsService rptService = null;
            SAPbobsCOM.ReportLayout oReportLayout = null;
            SAPbobsCOM.ReportLayoutParams newReportParam = null;
            SAPbobsCOM.BlobParams oBlobParams = null;


            try
            {
                oCacheInformation = BYBCache.Instance.getFromCache(T1.Properties.Settings.Default.VersionControlCacheName);
                if (oCacheInformation != null)
                {
                    XmlNode oReportNode = oCacheInformation.SelectSingleNode("/VersionControl/Module[./AdminInfo/Code/text() = '" + ModuleCode + "']/Installation/Report/Item[@name='" + Name + "' and @typeCode='" + TypeCode + "']");
                    bool blCreateType = false;
                    if (oReportNode != null)
                    {
                        if (oReportNode.Attributes["status"].Value == Status.installed.ToString())
                        {
                            strReportCode = oReportNode.Attributes["code"].Value;
                            AddReportInformation(TypeCode, Name, strReportCode, Status.installed);
                            
                        }
                        else
                        {
                            blCreateType = true;
                        }
                    }
                    else
                    {
                        blCreateType = true;
                    }


                    if (blCreateType)
                    {

                        oCompanyService = BYBB1MainObject.Instance.B1Company.GetCompanyService();
                        //rptTypeService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);

                        //newTypeParam = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportTypeParams);
                        //newTypeParam.TypeCode = TypeCode;
                        //oReport = rptTypeService.GetReportType(newTypeParam);

                        rptService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
                        oReportLayout = rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);
                        oReportLayout.Author = T1.Properties.Settings.Default.ReportAuthor;
                        oReportLayout.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal;
                        oReportLayout.Name = Name;
                        oReportLayout.TypeCode = TypeCode;
                        newReportParam = rptService.AddReportLayout(oReportLayout);

                        strReportCode = newReportParam.LayoutCode;

                        //oReport = rptTypeService.GetReportType(newTypeParam);
                        //oReport.DefaultReportLayout = newReportParam.LayoutCode;
                        //rptTypeService.UpdateReportType(oReport);


                        oBlobParams = oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams); 
                        oBlobParams.Table = "RDOC"; 
                        oBlobParams.Field = "Template"; 
                        SAPbobsCOM.BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add(); 
                        oKeySegment.Name = "DocCode"; 
                        oKeySegment.Value = newReportParam.LayoutCode;
                        FileStream oFile = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @FileName , System.IO.FileMode.Open); 
                        int fileSize = (int)oFile.Length; 
                        byte[] buf = new byte[fileSize]; 
                        oFile.Read(buf, 0, fileSize);
                        oFile.Dispose();
                        SAPbobsCOM.Blob oBlob = (SAPbobsCOM.Blob)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob); 
                        oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);
                        oCompanyService.SetBlob(oBlobParams, oBlob);


                        AddReportInformation(TypeCode, Name, strReportCode, Status.installed);
                    }
                }
                else
                {
                    Exception er = new Exception(Convert.ToString("SAPIFRS Error::No Config File Found in cache"));
                    BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.AddReport", er, 1, System.Diagnostics.EventLogEntryType.Error);
                    AddReportInformation(TypeCode, Name, strReportCode, Status.failed);
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.AddReport", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.AddReport", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }

            return strReportCode;


        }

        public bool updateModuleRegistration(string ModuleCode)
        {
            bool blResult = false;
            XmlDocument objDocument = null;
            XmlNode oModuleNode = null;
            try
            {
                objDocument = BYBCache.Instance.getFromCache(T1.Properties.Settings.Default.VersionControlCacheName);
                if (objDocument != null)
                {
                    oModuleNode = objDocument.SelectSingleNode("/VersionControl/Module[./AdminInfo/Code/text() = '" + ModuleCode + "']");
                    if (oModuleNode != null)
                    {
                        objDocument.SelectSingleNode("/VersionControl").RemoveChild(oModuleNode);
                    }
                }
                XmlDocumentFragment fragment = objDocument.CreateDocumentFragment();
                fragment.InnerXml = oXMLRegistration.SelectSingleNode("/Module").OuterXml;

                objDocument.SelectSingleNode("/VersionControl").AppendChild(fragment);

                BYBCache.Instance.addToCache(T1.Properties.Settings.Default.VersionControlCacheName, objDocument, BYBCache.objCachePriority.NotRemovable);
                updateB1InstallationInformation();

            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.updateModuleRegistration", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.updateModuleRegistration", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            return blResult;
        }

        private void updateB1InstallationInformation()
        {
            SAPbobsCOM.UserTables oUserTables = null;
            SAPbobsCOM.UserTable oConfigTable = null;
            XmlDocument oDocument = null;

            try
            {
                oUserTables = BYBB1MainObject.Instance.B1Company.UserTables;
                oDocument = BYBCache.Instance.getFromCache(T1.Properties.Settings.Default.VersionControlCacheName);
                
                try
                {
                    oConfigTable = oUserTables.Item(T1.Properties.Settings.Default.versionControlB1Table);
                    oConfigTable.GetByKey(T1.Properties.Settings.Default.versionControlKey);
                    string strConfigValue = oDocument.OuterXml;


                    oConfigTable.UserFields.Fields.Item(T1.Properties.Settings.Default.versionControlField).Value = strConfigValue;
                    if (oConfigTable.Update() != 0)
                    {
                        Exception er = new Exception(Convert.ToString("COM Error::" + BYBB1MainObject.Instance.B1Company.GetLastErrorCode().ToString() + "::" + BYBB1MainObject.Instance.B1Company.GetLastErrorDescription() + "::" ));

                    }
                    

                }
                catch (COMException comEx)
                {
                    Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                    BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.updateB1InstallationInformation", er, 1, System.Diagnostics.EventLogEntryType.Error);
                }
                catch (Exception er)
                {
                    BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.updateB1InstallationInformation", er, 1, System.Diagnostics.EventLogEntryType.Error);
                }
            }
            catch (COMException comEx)
            {
                Exception er = new Exception(Convert.ToString("COM Error::" + comEx.ErrorCode + "::" + comEx.Message + "::" + comEx.StackTrace));
                BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.updateB1InstallationInformation", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
            catch (Exception er)
            {
                BYBExceptionHandling.reportException(er.Message, "ModuleRegistrationClass.updateB1InstallationInformation", er, 1, System.Diagnostics.EventLogEntryType.Error);
            }
        }

        




    }
}
