using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.Shared.Classes
{
    public class ObjectCron
    {
        public int Id { get; set; }
        public int ObjectDefId { get; set; }
        public string Cron { get; set; }
        public string JobId { get; set; }
        public string TriggerId { get; set; }
        public string GroupId { get; set; }
        public string Instance { get; set; }
        public string Library { get; set; }
    }

    public class ObjectDef
    {
        public int Id { get; set; }
        public string B1Object { get; set; }
        public string Tables { get; set; }
        
    }

    public class InstallationInformation
    {

        public string nancyLocalAddress { get; set; }

        public T1Server T1Server { get; set; }
    }

    public class T1Server
    {
        public string logLevel { get; set; }
        public string moduleListFile { get; set; }
    }

    public class NancyModuleLibraryInfo
    {
        public string ExternalLibrary { get; set; }
        public string NameSpace { get; set; }
        public string Contructor { get; set; }
        public string LibraryType { get; set; }
    }
}
