using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit.Sdk;
using Microsoft.Extensions.Logging;
using Moq;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using IcbcExcelMergeWeb.DotNetCore.Services;

namespace IcbcExcelMergeWeb.DotNetCore.Controllers.Tests
{
    [TestClass()]
    public class HomeControllerTests
    {
        [TestMethod()]
        public void GetReportDataFileNameFromConfigTest()
        {
            var logger = new Mock<ILogger<HomeController>>();

            var hostingEnvironment = new Mock<IWebHostEnvironment>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["ReportDataFileName"]).Returns("ReportsSchema.xml");

            var reportService = new Mock<IReportService>();

            HomeController homeController = 
                new HomeController(
                    logger.Object, 
                    hostingEnvironment.Object,
                    configuration.Object,
                    reportService.Object);

            string reportDataFileName = homeController.GetReportDataFileName("ReportDataFileName", "Reports.xml");
            Assert.AreEqual("ReportsSchema.xml", reportDataFileName);
        }

        [TestMethod()]
        public void GetReportDataFileNameDefaultTest()
        {
            var logger = new Mock<ILogger<HomeController>>();

            var hostingEnvironment = new Mock<IWebHostEnvironment>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["ReportDataFileName"]).Returns("");

            var reportService = new Mock<IReportService>();

            HomeController homeController =
                new HomeController(
                    logger.Object,
                    hostingEnvironment.Object,
                    configuration.Object,
                    reportService.Object);

            string reportDataFileName = homeController.GetReportDataFileName("ReportDataFileName", "Reports.xml");
            Assert.AreEqual("Reports.xml", reportDataFileName);
        }
    }
}