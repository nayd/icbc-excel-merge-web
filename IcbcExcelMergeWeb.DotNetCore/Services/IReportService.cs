using IcbcExcelMergeWeb.DotNetCore.Models;
using Microsoft.AspNetCore.Http;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace IcbcExcelMergeWeb.DotNetCore.Services
{
    public interface IReportService
    {
        string MergeReport(IFormFile file, string uploadPath, string xmlReportsPath);
    }
}
