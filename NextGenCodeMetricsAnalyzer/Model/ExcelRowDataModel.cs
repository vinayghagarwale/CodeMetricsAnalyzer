using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NextGenCodeMetricsAnalyzer.Model
{
    public enum scope
    {
        Project,
        Namespace,
        Type,
        Member
    }
    public class ExcelRowDataModel
    {
        public string Scope {get; set; }
        public string Project { get; set; }
        public string Namespace { get; set; }
        public string Type { get; set; }
        public string Member { get; set; }
        public string MaintainabilityIndex { get; set; }
        public string CyclomaticComplexity { get; set; }
        public string DepthofInheritance { get; set; }
        public string ClassCoupling { get; set; }
        public string LinesofCode { get; set; }
    }
    public class ExcelRowDataModelList : List<ExcelRowDataModel>
    {
        public ExcelRowDataModelList()
        {
            this.Add(new ExcelRowDataModel());
        }
    }
}
