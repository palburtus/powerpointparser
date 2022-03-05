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

namespace PowerPointParser.Tests
{
    [TestClass()]
    public class ParserTests
    {
        [TestMethod()]
        [DeploymentItem("TestData/TestDeckOne.pptx")]
        public void Parse_ParseTestDeck_ReturnsIntWrapperMap()
        {
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var path = System.IO.Path.Combine(directory!, "TestDeckOne.pptx");
            

            Mock<ILogger> logger = new Mock<ILogger>();

            IParser parser = new Parser(new HtmlConverter(), logger.Object);
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(4, map.Keys.Count);
        }

        [TestMethod()]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseNoteParagraph_ReturnsIntWrapperMap()
        {
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var path = System.IO.Path.Combine(directory!, "TestDeckParagraph.pptx");


            Mock<ILogger> logger = new Mock<ILogger>();

            IParser parser = new Parser(new HtmlConverter(), logger.Object);
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map.Keys.Count);

            var actual = map[2][0];

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.PPr);
            Assert.AreEqual("This note is just a paragraph", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
            Assert.IsNull(actual.Text);
        }
    }
}