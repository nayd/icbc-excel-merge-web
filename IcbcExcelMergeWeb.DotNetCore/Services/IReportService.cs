using IcbcExcelMergeWeb.DotNetCore.Models;
using Microsoft.AspNetCore.Http;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace IcbcExcelMergeWeb.DotNetCore.Services
{
    public interface IReportService
    {
        string BuildReport(IFormFile file, string uploadPath, string xmlReportsPath);
        string MergeReport(Reports reportResults, XSSFWorkbook workbook, string result);
        string RenderWorkbook(XSSFWorkbook workbook, string result);
    }
}
