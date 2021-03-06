using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Hosting;
using System.Text;
using System.IO;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using IcbcExcelMergeWeb.DotNetCore.Services;

namespace IcbcExcelMergeWeb.DotNetCore.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> logger;
        private IWebHostEnvironment hostingEnvironment;
        private readonly IConfiguration configuration;
        private readonly IReportService reportService;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment, IConfiguration configuration, IReportService reportService)
        {
            this.logger = logger;
            this.hostingEnvironment = hostingEnvironment;
            this.configuration = configuration;
            this.reportService = reportService;
        }

        public IActionResult Index()
        {
            return View();
        }
      
        public ActionResult Import()
        {
            //TODO: add validation
            IFormFile file = Request.Form.Files[0];

            string folderName = "UploadExcel";
            string webRootPath = hostingEnvironment.WebRootPath;
            string uploadPath = Path.Combine(webRootPath, folderName);
            string reportDataFileName = GetReportDataFileName("ReportDataFileName", "Reports.xml");

            string reportDataFilePath = Path.Combine(webRootPath, reportDataFileName);

            return this.Content(reportService.BuildReport(file, uploadPath, reportDataFilePath));
        }

        public string GetReportDataFileName(string configName, string defaultFileName)
        {
            //TODO: create config repository where all configuration options are defined, and provide reusable parsers for config data types

            if (configuration[configName] == "")
            {
                logger.LogWarning($"{configName} not specified in configuration, defaulting to {defaultFileName}");
                return defaultFileName;
            }

            return configuration[configName];
        }
    }
}
