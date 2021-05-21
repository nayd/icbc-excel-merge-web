using IcbcExcelMergeWeb.DotNetCore.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IcbcExcelMergeWeb.DotNetCore.Services
{
    public class ReportService : IReportService
    {
        public string MergeReport(ISheet sheet, Reports report)
        {
            StringBuilder sb = new StringBuilder();

            return sb.ToString();
        }
    }
}
