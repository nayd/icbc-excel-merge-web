using IcbcExcelMergeWeb.DotNetCore.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace IcbcExcelMergeWeb.DotNetCore.Services
{
    interface IReportService
    {
        string MergeReport(ISheet sheet, Reports report);
    }
}
