using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Microsoft.Extensions.Logging;
using Moq;

namespace PowerPointParser.Tests
{
    [TestClass()]
    public class ParserTests
    {
        [ClassInitialize]
        public static void ClassSetup(TestContext context)
        {

        }

        [TestMethod]
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

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseNoteParagraph_ReturnsIntWrapperMap()
        {
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var path = System.IO.Path.Combine(directory!, "TestDeckParagraph.pptx");


            Mock<ILogger> logger = new Mock<ILogger>();

            IParser parser = new Parser(new HtmlConverter(), logger.Object);
            var map = parser.ParseSpeakerNotes(path);

            var actual = map[2][0];

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.PPr);
            Assert.AreEqual("This note is just a paragraph", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
            Assert.IsNull(actual.Text);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseNoteConsecutiveParagraphs_ReturnsIntWrapperMap()
        {
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var path = System.IO.Path.Combine(directory!, "TestDeckParagraph.pptx");


            Mock<ILogger> logger = new Mock<ILogger>();

            IParser parser = new Parser(new HtmlConverter(), logger.Object);
            var map = parser.ParseSpeakerNotes(path);

            var actual = map[3][1];

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.PPr);
            Assert.IsNull(actual.Text);

            Assert.AreEqual("And this is a ", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

            Assert.AreEqual("second paragraph", actual.R![1].T);
            Assert.AreEqual(0, actual.R![1].RPr!.B);
            Assert.AreEqual(0, actual.R![1].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![1].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseBoldParagraph_ReturnsIntWrapperMap()
        {
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var path = System.IO.Path.Combine(directory!, "TestDeckParagraph.pptx");


            Mock<ILogger> logger = new Mock<ILogger>();

            IParser parser = new Parser(new HtmlConverter(), logger.Object);
            var map = parser.ParseSpeakerNotes(path);

            var actual = map[4][0];

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.PPr);
            Assert.IsNull(actual.Text);

            Assert.AreEqual("This is a bold paragraph", actual.R![0].T);
            Assert.AreEqual(1, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
            
        }
    }
}