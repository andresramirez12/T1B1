using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nancy;
using Nancy.Helpers;
using Nancy.ModelBinding;
using Nancy.Extensions;
using Newtonsoft.Json;
using System.Xml;
using System.IO;

namespace T1.Nancy.Modules.MediosMagneticos
{
    public class MesiosMagneticosModule : NancyModule
    {
        public MesiosMagneticosModule()
        {
            Get["MM/execute"] = x =>
            {
                _MediosMagneticosClass objMMClass = new _MediosMagneticosClass();
                objMMClass.executeMM();

                var response = (Response)"";
                response.ContentType = "application/txt";
                response.StatusCode = HttpStatusCode.OK;
                return response;
            }
        }
    }
}
