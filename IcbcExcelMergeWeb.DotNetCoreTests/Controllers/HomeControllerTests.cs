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
        public void GetReportsXmlFileNameFromConfigTest()
        {
            var logger = new Mock<ILogger<HomeController>>();

            var hostingEnvironment = new Mock<IWebHostEnvironment>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["ReportsXml"]).Returns("ReportsSchema.xml");

            var reportService = new Mock<IReportService>();

            HomeController homeController = 
                new HomeController(
                    logger.Object, 
                    hostingEnvironment.Object,
                    configuration.Object,
                    reportService.Object);

            string reportsXmlFileName = homeController.GetReportsXmlFile();
            Assert.AreEqual("ReportsSchema.xml", reportsXmlFileName);
        }

        [TestMethod()]
        public void GetReportsXmlFileNameDefaultTest()
        {
            var logger = new Mock<ILogger<HomeController>>();

            var hostingEnvironment = new Mock<IWebHostEnvironment>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["ReportsXml"]).Returns("");

            var reportService = new Mock<IReportService>();

            HomeController homeController =
                new HomeController(
                    logger.Object,
                    hostingEnvironment.Object,
                    configuration.Object,
                    reportService.Object);

            string reportsXmlFileName = homeController.GetReportsXmlFile();
            Assert.AreEqual("Reports.xml", reportsXmlFileName);
        }
    }
}