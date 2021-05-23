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

        /// <summary>
        /// Store an excel file from the upload and build a report using server-stored XML file with data
        /// </summary>
        /// <param name="file">The web input file</param>
        /// <param name="uploadPath">The upload path to store the Excel file</param>
        /// <param name="xmlReportsPath">The server-stored XML data file path</param>
        /// <returns></returns>
        public string BuildReport(IFormFile file, string uploadPath, string xmlReportsPath)
        {
            string output = "Could not generate a report. Please check the logs for more information.";

            // TODO: this can be abstracted to a report results service which can validate report name, etc. 
            // and get results from web service instead of xml file
            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Reports));
            Reports reportResults = (Reports)serializer.Deserialize(XmlReader.Create(xmlReportsPath));
            
            if (file.Length > 0)
            {
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }
                string fileExtension = Path.GetExtension(file.FileName).ToLower();
                string uploadedFilePath = Path.Combine(uploadPath, file.FileName);

                using var stream = new FileStream(uploadedFilePath, FileMode.Create);
                file.CopyTo(stream);
                stream.Position = 0;

                //TODO: abstract with strategy for file extension
                if (fileExtension == ".xls")
                {
                    //read Excel 97-2000 formats
                    HSSFWorkbook workbookOldFormat = new HSSFWorkbook(stream);
                }
                else
                {
                    // read 2007 Excel format
                    // we assume here the file is .xlsx but this could probably be probed from the stream
                    XSSFWorkbook workbook = new XSSFWorkbook(stream);

                    //TODO: abstract with report builder
                    output = MergeReport(reportResults, workbook, output);
                    output = RenderWorkbook(workbook, output);
                   
                    return output;
                }
            }

            return output;
        }

        /// <summary>
        /// Merge report data already parsed in <paramref name="reportResults"/> into an NPOI <paramref name="workbook"/>
        /// </summary>
        /// <param name="reportResults">The report data</param>
        /// <param name="workbook">The NPOI workbook</param>
        /// <param name="result">Current user messages</param>
        /// <returns>Output any useful messages to the user</returns>
        public string MergeReport(Reports reportResults, XSSFWorkbook workbook, string result)
        {
            string sheetName = GetSheetName("SheetName", "Sheet1");

            // validate report name with report data
            // OTODO: add test
            if (reportResults.Report.Name.ToUpper() != sheetName.ToUpper())
            {
                result = "Could not generate the report. Report data did not match report name.";
                logger.LogWarning(result);
                return result;
            }

            for (int i = workbook.NumberOfSheets - 1; i >= 0; i--)
            {
                if (workbook[i].SheetName != sheetName)
                {
                    workbook.RemoveAt(i);
                }
            }

            if (workbook.NumberOfSheets != 1)
            {
                result = $"Sheet name '{sheetName}' was not found.";
                logger.LogWarning(result);
                return result;
            }

            // the workbook now contains only the sheet that we need
            ISheet sheet = workbook.GetSheet(sheetName);
            for (int i = 0; i < reportResults.Report.ReportVal.Length; i++)
            {
                int rowColumnOffset = GetReportYAxisColumnOffset("ReportYAxisColumnOffset", 1);
                int colColumnOffset = sheet.GetRow(0).Cells.Count - 1; // TODO: we depend on NPOI to determine the max cells correctly, which it currently does.
                ICell rowCell = GetCellByValue(sheet, reportResults.Report.ReportVal[i].ReportRow, rowColumnOffset);
                ICell colCell = GetCellByValue(sheet, reportResults.Report.ReportVal[i].ReportCol, colColumnOffset);

                // TODO: add test
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

            return result;
        }

        /// <summary>
        /// Render the final HTML table using an NPOI Excel converter which preserves the Excel formatting
        /// </summary>
        /// <param name="workbook">The NPOI workbook containing the final report</param>
        /// <param name="result">Current user messages</param>
        /// <returns></returns>
        /// <remarks>See <see cref="ExcelToHtmlConverter"/></remarks>
        public string RenderWorkbook(XSSFWorkbook workbook, string result)
        {
            ExcelToHtmlConverter converter = new ExcelToHtmlConverter();
            converter.OutputRowNumbers = false;
            converter.OutputColumnHeaders = false;
            converter.ProcessWorkbook(workbook);

            return converter.Document.InnerXml;
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
            //TODO: assumption report results are x-y-axis based, or always a rectangular range
            //TODO: add test

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

        public string GetSheetName(string configName, string defaultSheetName)
        {
            //TODO: create config repository where all configuration options are defined, and provide reusable parsers for config data types

            if (configuration[configName] == "")
            {
                logger.LogWarning($"sheet name is defaulting to {defaultSheetName}, please add {configName} in appsettings.json if you " +
                    $"need to change this");
                return defaultSheetName;
            }

            return configuration[configName];
        }

        public int GetReportYAxisColumnOffset(string configName, int defaultReportYAxisColumnOffset)
        {
            //TODO: create config repository where all configuration options are defined, and provide reusable parsers for config data types

            if (!int.TryParse(this.configuration[configName], out int reportYAxisColumnOffset))
            {
                logger.LogWarning($"report Y Axis column offset is defaulting to {defaultReportYAxisColumnOffset}, please add {configName} " +
                    $"in appsettings.json if you need to change this");
                reportYAxisColumnOffset = defaultReportYAxisColumnOffset;
            }

            return reportYAxisColumnOffset;
        }
    }
}
