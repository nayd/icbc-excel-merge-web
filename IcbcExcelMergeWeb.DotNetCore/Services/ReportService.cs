using IcbcExcelMergeWeb.DotNetCore.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using NPOI.HSSF.UserModel;
using NPOI.SS.Converter;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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
            string result = "Could not generate a report. Please check the logs for more information.";

            // TODO: this can be abstracted to a report results service which can validate report name, etc. 
            // and get results from web service instead of xml file
            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Reports));
            Reports reportResults = (Reports)serializer.Deserialize(XmlReader.Create(xmlReportsPath));
            string sheetName = GetSheetName();
            
            // validate report name with report data
            //TODO: add test
            if (reportResults.Report.Name.ToUpper() != sheetName.ToUpper())
            {
                result = "Could not generate the report. Report data did not match report name.";
                logger.LogWarning(result);
                return result;
            }

            if (file.Length > 0)
            {
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }
                string fileExtension = Path.GetExtension(file.FileName).ToLower();
                string uploadedFilePath = Path.Combine(uploadPath, file.FileName);

                // process sheet, merge data, and output formatted html table
                //TODO: add test
                using (var stream = new FileStream(uploadedFilePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                    stream.Position = 0;
                    int sheetIndex = 0;

                    if (fileExtension == ".xls")
                    {
                        HSSFWorkbook workbookOldFormat = new HSSFWorkbook(stream); //This will read the Excel 97-2000 formats
                        ISheet sheet = workbookOldFormat.GetSheetAt(sheetIndex);
                    }
                    else
                    {
                        // we assume here the file is .xlsx but this could probably be probed from the stream
                        XSSFWorkbook workbook = new XSSFWorkbook(stream); //This will read 2007 Excel format

                        for (int i = workbook.NumberOfSheets - 1; i >= 0; i--)
                        {
                            if (workbook[i].SheetName != sheetName)
                            {
                                workbook.RemoveAt(i);
                            }
                            else
                            {
                                sheetIndex = i;
                            }
                        }

                        if (workbook.NumberOfSheets != 1)
                        {
                            result = $"Sheet name '{sheetName}' was not found.";
                            logger.LogWarning(result);
                            return result;
                        }

                        // the workbook now contains only the sheet that we need
                        ISheet sheet = workbook.GetSheetAt(sheetIndex);
                        for (int i = 0; i < reportResults.Report.ReportVal.Length; i++)
                        {
                            var reportValue = reportResults.Report.ReportVal[i];
                            int rowNum = reportValue.ReportRow;

                            int rowOffset = 1;                               // TODO: add to config
                            int colOffset = sheet.GetRow(0).Cells.Count - 1; // TODO: we depend on NPOI to determine the max cells correctly, which it currently does.
                            ICell rowCell = GetCellByValue(sheet, reportResults.Report.ReportVal[i].ReportRow, rowOffset);
                            ICell colCell = GetCellByValue(sheet, reportResults.Report.ReportVal[i].ReportCol, colOffset);

                            //TODO: add test
                            if (rowCell == null || colCell == null)
                            {
                                result = $"Could not generate the report. Unable to determine rowCell or colCell for report value {reportResults.Report.ReportVal[i].Val}.";
                                logger.LogWarning(result);
                                return result;
                            }

                            sheet.GetRow(rowCell.RowIndex).GetCell(colCell.ColumnIndex).SetCellValue(reportResults.Report.ReportVal[i].Val);

                            // audit
                            // TODO: make configurable
                            var cellReference = new CellReference(rowCell.RowIndex, colCell.ColumnIndex);
                            logger.LogInformation($"Sheet '{sheetName}': value '{reportResults.Report.ReportVal[i].Val,10}' -> {cellReference.FormatAsString()}");
                        }

                        // the NPOI converter will keep the formatting in the resulting HTML table
                        ExcelToHtmlConverter converter = new ExcelToHtmlConverter();
                        converter.OutputRowNumbers = false;
                        converter.OutputColumnHeaders = false;
                        converter.ProcessWorkbook(workbook);

                        return converter.Document.InnerXml;
                    }                    
                }
            }

            return result;
        }

        /// <summary>
        /// Find a cell in the given sheet by its value.
        /// </summary>
        /// <param name="sheet">The sheet we are searching in</param>
        /// <param name="dataValue">The value to search for</param>
        /// <param name="columnOffset">The column offset to stop searching</param>
        /// <returns></returns>
        private static ICell GetCellByValue(ISheet sheet, int dataValue, int columnOffset)
        {
            //TODO: assumption report rows and columns indications on the sheet are parsable to int
            //TODO: added test

            foreach (IRow row in sheet)
            {
                foreach (ICell cell in row.Cells)
                {
                    // lookup up to the column offset and skip to the next row
                    if (columnOffset + 1 == cell.ColumnIndex)
                    {
                        break;
                    }

                    if (cell.StringCellValue != "" && int.TryParse(cell.StringCellValue, out int value) && dataValue == value)
                    {
                        return cell;
                    }
                }
            }

            return null;
        }

        [Obsolete("First cut by iterating each row and cell. Replaced with converter which keeps the Excel styles")]
        public static string MergeSheet(ISheet sheet, Reports reports)
        {
            StringBuilder sb = new StringBuilder();

            IRow headerRow = sheet.GetRow(0); //Get Header Row
            int cellCount = headerRow.LastCellNum;
            sb.AppendLine("<table class='table table-bordered'>");
            sb.AppendLine("<tr>");
            for (int j = 0; j < cellCount; j++)
            {
                ICell cell = headerRow.GetCell(j);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                sb.Append("<th>" + cell.ToString() + "</th>");
            }
            sb.AppendLine("</tr>");
            sb.AppendLine("<tr>");
            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++) //Read Excel File
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                        sb.AppendLine("<td>" + (row.GetCell(j).StringCellValue == "" ? "&nbsp;" : row.GetCell(j).StringCellValue) + "</td>");
                }
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</table>");

            return sb.ToString();
        }

        public string GetSheetName()
        {
            //TODO: create config repository where all configuration options are defined, and provide reusable parsers for config data types

            string defaultSheetName = "Sheet1";
            string configName = "SheetName";

            if (configuration[configName] == "")
            {
                logger.LogWarning($"sheet name is defaulting to {defaultSheetName}, please add {configName} in appsettings.json if you need to change this");
                return defaultSheetName;
            }

            return configuration[configName];
        }
    }
}
