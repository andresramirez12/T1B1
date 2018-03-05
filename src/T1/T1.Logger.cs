using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Config;

[assembly: log4net.Config.XmlConfigurator(ConfigFile = "logConfig.xml", Watch = true)]
namespace T1
{
    
    public class Logger
    {

    }
}
