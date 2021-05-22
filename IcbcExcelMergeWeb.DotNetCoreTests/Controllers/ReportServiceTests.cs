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
        public void GetSheetIndexFromConfigTest()
        {
            var logger = new Mock<ILogger<ReportService>>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["SheetIndex"]).Returns("1");

            ReportService reportService =
                new ReportService(
                    configuration.Object,
                    logger.Object);

            int sheetIndex = reportService.GetSheetIndex();
            Assert.AreEqual(1, sheetIndex);
        }

        [TestMethod()]
        public void GetSheetIndexDefaultTest()
        {
            var logger = new Mock<ILogger<ReportService>>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["SheetIndex"]).Returns("");

            ReportService reportService =
                new ReportService(
                    configuration.Object,
                    logger.Object);

            int sheetIndex = reportService.GetSheetIndex();
            Assert.AreEqual(0, sheetIndex);
        }
    }
}
