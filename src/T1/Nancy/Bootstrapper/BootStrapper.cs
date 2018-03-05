using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nancy;
using Nancy.Bootstrapper;
using Nancy.TinyIoc;
using Nancy.Conventions;

namespace T1.Nancy.Bootstrapper
{
    public class BootStrapper : DefaultNancyBootstrapper
    {
        protected override void ConfigureConventions(NancyConventions nancyConventions)
        {
            base.ConfigureConventions(nancyConventions);
            nancyConventions.StaticContentsConventions.Clear();
            nancyConventions.StaticContentsConventions.Add
            (StaticContentConventionBuilder.AddDirectory("/", "/Nancy/Content"));
            nancyConventions.StaticContentsConventions.Add
            (StaticContentConventionBuilder.AddDirectory("/json", "/Nancy/Content/json"));
            nancyConventions.StaticContentsConventions.Add
            (StaticContentConventionBuilder.AddDirectory("/ajax", "/Nancy/Content/ajax"));
            nancyConventions.StaticContentsConventions.Add
            (StaticContentConventionBuilder.AddDirectory("/pages", "/Nancy/Content/pages"));
            nancyConventions.StaticContentsConventions.Add
            (StaticContentConventionBuilder.AddDirectory("/admin", "/Nancy/Content/Admin"));
        }
    }
}
