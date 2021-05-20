using Microsoft.VisualStudio.TestTools.UnitTesting;
using IcbcExcelMergeWeb.DotNetCore.Controllers;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit.Sdk;
using Microsoft.Extensions.Logging;
using Moq;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;

namespace IcbcExcelMergeWeb.DotNetCore.Controllers.Tests
{
    [TestClass()]
    public class HomeControllerTests
    {
        [TestMethod()]
        public void GetSheetIndexFromConfigTest()
        {
            var logger = new Mock<ILogger<HomeController>>();

            var hostingEnvironment = new Mock<IHostingEnvironment>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["SheetIndex"]).Returns("1");

            HomeController homeController = 
                new HomeController(
                    logger.Object, 
                    hostingEnvironment.Object,
                    configuration.Object);

            int sheetIndex = homeController.GetSheetIndex();
            Assert.AreEqual(1, sheetIndex);
        }

        [TestMethod()]
        public void GetSheetIndexDefaultTest()
        {
            var logger = new Mock<ILogger<HomeController>>();

            var hostingEnvironment = new Mock<IHostingEnvironment>();

            var configuration = new Mock<IConfiguration>();
            configuration.Setup(x => x["SheetIndex"]).Returns("");

            HomeController homeController =
                new HomeController(
                    logger.Object,
                    hostingEnvironment.Object,
                    configuration.Object);

            int sheetIndex = homeController.GetSheetIndex();
            Assert.AreEqual(0, sheetIndex);
        }
    }
}