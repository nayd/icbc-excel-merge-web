using IcbcExcelMergeWeb.DotNetCore.Controllers;
using IcbcExcelMergeWeb.DotNetCore.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Text;

namespace IcbcExcelMergeWeb.DotNetCoreTests.Controllers
{
    [TestClass()]
    public class ReportServiceTests
    {
        [TestMethod()]
        public void GetSheetNameFromConfigTest()
        {
            var logger = new Mock<ILogger<ReportService>>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["SheetName"]).Returns("F 20.04");

            ReportService reportService =
                new ReportService(
                    configuration.Object,
                    logger.Object);

            string sheetName = reportService.GetSheetName("Sheet1");
            Assert.AreEqual("F 20.04", sheetName);
        }

        [TestMethod()]
        public void GetSheetNameDefaultTest()
        {
            var logger = new Mock<ILogger<ReportService>>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["SheetName"]).Returns(string.Empty);

            ReportService reportService =
                new ReportService(
                    configuration.Object,
                    logger.Object);

            string sheetName = reportService.GetSheetName("Sheet1");
            Assert.AreEqual("Sheet1", sheetName);
        }
    }
}
