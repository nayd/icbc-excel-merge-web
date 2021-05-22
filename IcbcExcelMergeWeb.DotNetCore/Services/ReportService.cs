using IcbcExcelMergeWeb.DotNetCore.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace IcbcExcelMergeWeb.DotNetCore.Services
{
    public class ReportService : IReportService
    {
        private readonly IConfiguration configuration;
        private readonly ILogger<ReportService> logger;

        public ReportService(IConfiguration configuration, ILogger<ReportService> logger)
        {
            this.configuration = configuration;
            this.logger = logger;
        }

        public string MergeReport(IFormFile file, string uploadPath, string xmlReportsPath)
        {
            StringBuilder sb = new StringBuilder();

            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Reports));
            Reports reports = (Reports)serializer.Deserialize(XmlReader.Create(xmlReportsPath));

            if (file.Length > 0)
            {
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }
                string fileExtension = Path.GetExtension(file.FileName).ToLower();
                string uploadedFilePath = Path.Combine(uploadPath, file.FileName);

                ISheet sheet;
                int sheetIndex = GetSheetIndex();

                using (var stream = new FileStream(uploadedFilePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                    stream.Position = 0;

                    if (fileExtension == ".xls")
                    {
                        HSSFWorkbook hssfwb = new HSSFWorkbook(stream); //This will read the Excel 97-2000 formats  
                        sheet = hssfwb.GetSheetAt(sheetIndex);
                    }
                    else
                    {
                        XSSFWorkbook hssfwb = new XSSFWorkbook(stream); //This will read 2007 Excel format  
                        sheet = hssfwb.GetSheetAt(sheetIndex);
                    }
                    IRow headerRow = sheet.GetRow(0); //Get Header Row
                    int cellCount = headerRow.LastCellNum;
                    sb.Append("<table class='table table-bordered'><tr>");
                    for (int j = 0; j < cellCount; j++)
                    {
                        NPOI.SS.UserModel.ICell cell = headerRow.GetCell(j);
                        if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                        sb.Append("<th>" + cell.ToString() + "</th>");
                    }
                    sb.Append("</tr>");
                    sb.AppendLine("<tr>");
                    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++) //Read Excel File
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                                sb.Append("<td>" + row.GetCell(j).ToString() + "</td>");
                        }
                        sb.AppendLine("</tr>");
                    }
                    sb.Append("</table>");
                }
            }

            return sb.ToString();
        }

        public int GetSheetIndex()
        {
            int defaultSheetIndex = 0;

            if (!int.TryParse(this.configuration["SheetIndex"], out int sheetIndex))
            {
                logger.LogWarning("sheet index is defaulting to 0, please add SheetIndex in appsettings.json if you need to change this");
                return defaultSheetIndex;
            }

            return sheetIndex;
        }
    }
}
