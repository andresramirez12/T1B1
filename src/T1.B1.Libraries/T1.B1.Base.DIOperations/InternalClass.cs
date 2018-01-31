using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.Base.DIOperations
{
    public class projectInfo
    {
        public string ProjectCode { get; set; }
        public string ProjectName { get; set; }
    }

    public class costCenterInfo
    {
        public string CostCenterCode { get; set; }
        public string CostCenterName { get; set; }
        public int DimensionCode { get; set; }
    }

    public class dimensionInfo
    {
        public int DimentionCode { get; set; }
        public string DimensionName { get; set; }
        public bool isActive { get; set; }
    }
}
