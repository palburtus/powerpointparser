using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointParser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Moq;
using PowerPointParser.Dto;
using PowerPointParserTests.Mocks;

namespace PowerPointParser.Tests
{
    [TestClass()]
    public class ParserTests
    {
        [TestMethod()]
        [DeploymentItem("TestData/v0.1_ImprovedCodingPlanfor2022.pptx")]
        public void ParseTest()
        {
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var path = System.IO.Path.Combine(directory!, "v0.1_ImprovedCodingPlanFor2022.pptx");

            IHtmlConverter htmlConverter = new MockHtmlConverter();


            Mock<ILogger> logger = new Mock<ILogger>();

            IParser parser = new Parser(new HtmlConverter(), logger.Object);
            var slides = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, slides.Count);

            Assert.AreEqual(3, slides[2].SlidePosition);
            Assert.AreEqual("Ask devs for other examples", slides[2].SpeakerNotes);
        }
    }
}