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
            Assert.AreEqual(1, actual.R!.Count);
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
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual("And this is a second paragraph", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

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
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual("This is a bold paragraph", actual.R![0].T);
            Assert.AreEqual(1, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseUnorderedList_ReturnsIntWrapperMap()
        {
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var path = System.IO.Path.Combine(directory!, "TestDeckParagraph.pptx");


            Mock<ILogger> logger = new Mock<ILogger>();

            IParser parser = new Parser(new HtmlConverter(), logger.Object);
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[5].Count);

            var actualOne = map[5][0];

            Assert.IsNull(actualOne.A);
            Assert.IsNull(actualOne.Text);
            Assert.AreEqual(1, actualOne.R!.Count);
            Assert.AreEqual(0, actualOne.PPr!.Lvl);
            Assert.AreEqual("Unordered item 1", actualOne.R![0].T);
            Assert.AreEqual(0, actualOne.R![0].RPr!.B);
            Assert.AreEqual(0, actualOne.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualOne.R![0].RPr!.Lang);

            var actualTwo = map[5][1];

            Assert.IsNull(actualTwo.A);
            Assert.IsNull(actualTwo.Text);
            Assert.AreEqual(1, actualTwo.R!.Count);
            Assert.AreEqual(0, actualTwo.PPr!.Lvl);
            Assert.AreEqual("Unordered item 2", actualTwo.R![0].T);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.B);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualTwo.R![0].RPr!.Lang);

            var actualThree = map[5][2];

            Assert.IsNull(actualThree.A);
            Assert.IsNull(actualThree.Text);
            Assert.AreEqual(1, actualThree.R!.Count);
            Assert.AreEqual(0, actualThree.PPr!.Lvl);
            Assert.AreEqual("Unordered item 3", actualThree.R![0].T);
            Assert.AreEqual(0, actualThree.R![0].RPr!.B);
            Assert.AreEqual(0, actualThree.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![0].RPr!.Lang);

        }


        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseEmbeddedUnorderedList_ReturnsIntWrapperMap()
        {
            var directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var path = System.IO.Path.Combine(directory!, "TestDeckParagraph.pptx");


            Mock<ILogger> logger = new Mock<ILogger>();

            IParser parser = new Parser(new HtmlConverter(), logger.Object);
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[6].Count);

            var actualOne = map[6][0];

            Assert.IsNull(actualOne.A);
            Assert.IsNull(actualOne.Text);
            Assert.AreEqual(1, actualOne.R!.Count);
            Assert.AreEqual(0, actualOne.PPr!.Lvl);
            Assert.AreEqual("Unordered item 1", actualOne.R![0].T);
            Assert.AreEqual(0, actualOne.R![0].RPr!.B);
            Assert.AreEqual(0, actualOne.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualOne.R![0].RPr!.Lang);

            var actualTwo = map[6][1];

            Assert.IsNull(actualTwo.A);
            Assert.IsNull(actualTwo.Text);
            Assert.AreEqual(1, actualTwo.R!.Count);
            Assert.AreEqual(1, actualTwo.PPr!.Lvl);
            Assert.AreEqual("Unordered item 2", actualTwo.R![0].T);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.B);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualTwo.R![0].RPr!.Lang);

            var actualThree = map[6][2];

            Assert.IsNull(actualThree.A);
            Assert.IsNull(actualThree.Text);
            Assert.AreEqual(1, actualThree.R!.Count);
            Assert.AreEqual(2, actualThree.PPr!.Lvl);
            Assert.AreEqual("Unordered item 3", actualThree.R![0].T);
            Assert.AreEqual(0, actualThree.R![0].RPr!.B);
            Assert.AreEqual(0, actualThree.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![0].RPr!.Lang);

        }
    }
}