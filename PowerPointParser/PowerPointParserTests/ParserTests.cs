using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

// ReSharper disable once CheckNamespace - Test namespaces should match production
namespace Aaks.PowerPointParser.Tests
{
    [TestClass]
    public class ParserTests
    {
        private static string? _directory;
        private static string? _path;

        [ClassInitialize]
        public static void ClassSetup(TestContext context)
        {
            _directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            _path = Path.Combine(_directory!, "TestDeckParagraph.pptx");
            
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckOne.pptx")]
        public void Parse_ParseTestDeck_ReturnsIntWrapperMap()
        {
            var path = Path.Combine(_directory!, "TestDeckOne.pptx");

            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(4, map.Keys.Count);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseNoteParagraph_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(_path!);

            var actual = map[2][0]!;

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
        [DeploymentItem("TestData/TestThree.pptx")]
        public void Parse_ParseAlternativeFormat_ReturnsIntWrapperMap()
        {
            var path = Path.Combine(_directory!, "TestThree.pptx");
            IParser parser = new Parser();

            var map = parser.ParseSpeakerNotes(path);

            var actual = map[8][0]!;

            Assert.AreEqual("Caricas", actual.R![0].T);
            
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseNoteConsecutiveParagraphs_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(_path!);

            var actual = map[3][1]!;

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
            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(_path!);

            var actual = map[4][0]!;

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
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseItalicParagraph_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            var actual = map[1][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNotNull(actual.PPr);
            Assert.IsNull(actual.Text);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual("This is an italic paragraph", actual.R![0].T);
            Assert.AreEqual(1, actual.R![0].RPr!.I);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseUnderlinedParagraph_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            var actual = map[2][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.PPr);
            Assert.IsNull(actual.Text);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual("This is underlined", actual.R![0].T);
            Assert.AreEqual("sng", actual.R![0].RPr!.U);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseStrikeThroughParagraph_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            var actual = map[3][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.PPr);
            Assert.IsNull(actual.Text);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual("This text has a strike through", actual.R![0].T);
            Assert.AreEqual("sngStrike", actual.R![0].RPr!.Strike);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseCenterAlignedParagraph_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            var actual = map[4][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.AreEqual("ctr", actual.PPr!.Algn);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual("This text is center aligned", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

        }
     
        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseRightAlignedParagraph_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            var actual = map[5][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.AreEqual("r", actual.PPr!.Algn);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual("This text is right aligned", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseJustifiedAlignedParagraph_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            var actual = map[6][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.AreEqual("just", actual.PPr!.Algn);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual("This text is aligned justified", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseConsecutiveEmptySpaces_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            var actualOne = map[7][0]!;

            Assert.IsNull(actualOne.A);
            Assert.IsNull(actualOne.Text);
            Assert.IsNull(actualOne.PPr!.Algn);
            Assert.AreEqual(1, actualOne.R!.Count);
            Assert.AreEqual("Paragraph One", actualOne.R![0].T);
            Assert.AreEqual(0, actualOne.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualOne.R![0].RPr!.Lang);

            var actualTwo = map[7][1]!;

            Assert.IsNull(actualTwo.A);
            Assert.IsNull(actualTwo.Text);
            Assert.IsNull(actualTwo.PPr!.Algn);
            Assert.AreEqual(0, actualTwo.R!.Count);

            var actualThree = map[7][2]!;

            Assert.IsNull(actualThree.A);
            Assert.IsNull(actualThree.Text);
            Assert.IsNull(actualThree.PPr!.Algn);
            Assert.AreEqual(0, actualThree.R!.Count);

            var actualFour = map[7][3]!;

            Assert.IsNull(actualFour.A);
            Assert.IsNull(actualFour.Text);
            Assert.IsNull(actualFour.PPr!.Algn);
            Assert.AreEqual(1, actualFour.R!.Count);
            Assert.AreEqual("Paragraph Two after two spaces", actualFour.R![0].T);
            Assert.AreEqual(0, actualFour.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualFour.R![0].RPr!.Lang);

            var actualFive = map[7][4]!;

            Assert.IsNull(actualFive.A);
            Assert.IsNull(actualFive.Text);
            Assert.IsNull(actualFive.PPr!.Algn);
            Assert.AreEqual(0, actualFive.R!.Count);

            var actualSix = map[7][5]!;

            Assert.IsNull(actualSix.A);
            Assert.IsNull(actualSix.Text);
            Assert.IsNull(actualSix.PPr!.Algn);
            Assert.AreEqual(0, actualSix.R!.Count);

            var actualSeven = map[7][6]!;

            Assert.IsNull(actualSeven.A);
            Assert.IsNull(actualSeven.Text);
            Assert.IsNull(actualSeven.PPr!.Algn);
            Assert.AreEqual(1, actualSeven.R!.Count);
            Assert.AreEqual("Paragraph Three after three spaces", actualSeven.R![0].T);
            Assert.AreEqual(0, actualSeven.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualSeven.R![0].RPr!.Lang);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseUnorderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(3, map[5].Count);

            var actualOne = map[5][0]!;

            Assert.IsNull(actualOne.A);
            Assert.IsNull(actualOne.Text);
            Assert.AreEqual(1, actualOne.R!.Count);
            Assert.AreEqual(0, actualOne.PPr!.Lvl);
            Assert.AreEqual("•", actualOne.PPr!.BuChar!.Char);
            Assert.AreEqual("Unordered item 1", actualOne.R![0].T);
            Assert.AreEqual(0, actualOne.R![0].RPr!.B);
            Assert.AreEqual(0, actualOne.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualOne.R![0].RPr!.Lang);

            var actualTwo = map[5][1]!;

            Assert.IsNull(actualTwo.A);
            Assert.IsNull(actualTwo.Text);
            Assert.AreEqual(1, actualTwo.R!.Count);
            Assert.AreEqual(0, actualTwo.PPr!.Lvl);
            Assert.AreEqual("•", actualTwo.PPr!.BuChar!.Char);
            Assert.AreEqual("Unordered item 2", actualTwo.R![0].T);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.B);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualTwo.R![0].RPr!.Lang);

            var actualThree = map[5][2]!;

            Assert.IsNull(actualThree.A);
            Assert.IsNull(actualThree.Text);
            Assert.AreEqual(1, actualThree.R!.Count);
            Assert.AreEqual(0, actualThree.PPr!.Lvl);
            Assert.AreEqual("•", actualTwo.PPr!.BuChar!.Char);
            Assert.AreEqual("Unordered item 3", actualThree.R![0].T);
            Assert.AreEqual(0, actualThree.R![0].RPr!.B);
            Assert.AreEqual(0, actualThree.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![0].RPr!.Lang);

        }


        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseEmbeddedUnorderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(3, map[6].Count);

            var actualOne = map[6][0]!;

            Assert.IsNull(actualOne.A);
            Assert.IsNull(actualOne.Text);
            Assert.AreEqual(1, actualOne.R!.Count);
            Assert.AreEqual(0, actualOne.PPr!.Lvl);
            Assert.AreEqual("•", actualOne.PPr!.BuChar!.Char);
            Assert.AreEqual("Indent Unordered item 1", actualOne.R![0].T);
            Assert.AreEqual(0, actualOne.R![0].RPr!.B);
            Assert.AreEqual(0, actualOne.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualOne.R![0].RPr!.Lang);

            var actualTwo = map[6][1]!;

            Assert.IsNull(actualTwo.A);
            Assert.IsNull(actualTwo.Text);
            Assert.AreEqual(1, actualTwo.R!.Count);
            Assert.AreEqual(1, actualTwo.PPr!.Lvl);
            Assert.AreEqual("•", actualTwo.PPr!.BuChar!.Char);
            Assert.AreEqual("Indent Unordered item 2", actualTwo.R![0].T);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.B);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualTwo.R![0].RPr!.Lang);

            var actualThree = map[6][2]!;

            Assert.IsNull(actualThree.A);
            Assert.IsNull(actualThree.Text);
            Assert.AreEqual(1, actualThree.R!.Count);
            Assert.AreEqual(2, actualThree.PPr!.Lvl);
            Assert.AreEqual("•", actualThree.PPr!.BuChar!.Char);
            Assert.AreEqual("Indent Unordered item 3", actualThree.R![0].T);
            Assert.AreEqual(0, actualThree.R![0].RPr!.B);
            Assert.AreEqual(0, actualThree.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseOrderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(3, map[7].Count);

            var actualOne = map[7][0]!;

            Assert.IsNull(actualOne.A);
            Assert.IsNull(actualOne.Text);
            Assert.IsNull(actualOne.PPr!.BuChar);
            Assert.AreEqual("arabicPeriod", actualOne.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(1, actualOne.R!.Count);
            Assert.AreEqual(0, actualOne.PPr!.Lvl);
            Assert.AreEqual("Ordered one", actualOne.R![0].T);
            Assert.AreEqual(0, actualOne.R![0].RPr!.B);
            Assert.AreEqual(0, actualOne.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualOne.R![0].RPr!.Lang);

            var actualTwo = map[7][1]!;

            Assert.IsNull(actualTwo.A);
            Assert.IsNull(actualTwo.Text);
            Assert.IsNull(actualTwo.PPr!.BuChar);
            Assert.AreEqual("arabicPeriod", actualTwo.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(1, actualTwo.R!.Count);
            Assert.AreEqual(0, actualTwo.PPr!.Lvl);
            Assert.AreEqual("Ordered two", actualTwo.R![0].T);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.B);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualTwo.R![0].RPr!.Lang);

            var actualThree = map[7][2]!;

            Assert.IsNull(actualThree.A);
            Assert.IsNull(actualThree.Text);
            Assert.IsNull(actualThree.PPr!.BuChar);
            Assert.AreEqual("arabicPeriod", actualThree.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(1, actualThree.R!.Count);
            Assert.AreEqual(0, actualThree.PPr!.Lvl);
            Assert.AreEqual("Ordered three", actualThree.R![0].T);
            Assert.AreEqual(0, actualThree.R![0].RPr!.B);
            Assert.AreEqual(0, actualThree.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseEmbeddedOrderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(4, map[8].Count);

            var actualOne = map[8][0]!;

            Assert.IsNull(actualOne.A);
            Assert.IsNull(actualOne.Text);
            Assert.IsNull(actualOne.PPr!.BuChar);
            Assert.AreEqual("arabicPeriod", actualOne.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(1, actualOne.R!.Count);
            Assert.AreEqual(0, actualOne.PPr!.Lvl);
            Assert.AreEqual("Indent Ordered One", actualOne.R![0].T);
            Assert.AreEqual(0, actualOne.R![0].RPr!.B);
            Assert.AreEqual(0, actualOne.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualOne.R![0].RPr!.Lang);

            var actualTwo = map[8][1]!;

            Assert.IsNull(actualTwo.A);
            Assert.IsNull(actualTwo.Text);
            Assert.IsNull(actualTwo.PPr!.BuChar);
            Assert.AreEqual("arabicPeriod", actualTwo.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(2, actualTwo.R!.Count);
            Assert.AreEqual(1, actualTwo.PPr!.Lvl);
            Assert.AreEqual("Indent Ordered One ", actualTwo.R![0].T);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.B);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualTwo.R![0].RPr!.Lang);

            Assert.AreEqual("One", actualTwo.R![1].T);
            Assert.AreEqual(0, actualTwo.R![1].RPr!.B);
            Assert.AreEqual(0, actualTwo.R![1].RPr!.Dirty);
            Assert.AreEqual("en-US", actualTwo.R![1].RPr!.Lang);

            var actualThree = map[8][2]!;

            Assert.IsNull(actualThree.A);
            Assert.IsNull(actualThree.Text);
            Assert.IsNull(actualThree.PPr!.BuChar);
            Assert.AreEqual("arabicPeriod", actualThree.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(4, actualThree.R!.Count);
            Assert.AreEqual(2, actualThree.PPr!.Lvl);
            Assert.AreEqual("Indent Order One ", actualThree.R![0].T);
            Assert.AreEqual(0, actualThree.R![0].RPr!.B);
            Assert.AreEqual(0, actualThree.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![0].RPr!.Lang);

            Assert.AreEqual("One", actualThree.R![1].T);
            Assert.AreEqual(0, actualThree.R![1].RPr!.B);
            Assert.AreEqual(0, actualThree.R![1].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![1].RPr!.Lang);

            Assert.AreEqual("", actualThree.R![2].T);
            Assert.AreEqual(0, actualThree.R![2].RPr!.B);
            Assert.AreEqual(0, actualThree.R![2].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![2].RPr!.Lang);

            Assert.AreEqual("One", actualThree.R![3].T);
            Assert.AreEqual(0, actualThree.R![3].RPr!.B);
            Assert.AreEqual(0, actualThree.R![3].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![3].RPr!.Lang);

            var actualFour = map[8][3]!;

            Assert.IsNull(actualFour.A);
            Assert.IsNull(actualFour.Text);
            Assert.IsNull(actualFour.PPr!.BuChar);
            Assert.AreEqual("arabicPeriod", actualFour.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(1, actualFour.R!.Count);
            Assert.AreEqual(0, actualFour.PPr!.Lvl);
            Assert.AreEqual("Indent Ordered Three", actualFour.R![0].T);
            Assert.AreEqual(0, actualFour.R![0].RPr!.B);
            Assert.AreEqual(0, actualFour.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualFour.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_HollowRoundBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[9].Count);

            var actual = map[9][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuAutoNum);
            Assert.AreEqual("§", actual.PPr!.BuChar!.Char);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Filled Square Bullets", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_HollowSquareBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[10].Count);

            var actual = map[10][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuAutoNum);
            Assert.AreEqual("q", actual.PPr!.BuChar!.Char);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Hollow Square Bullets", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_StarBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[11].Count);

            var actual = map[11][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuAutoNum);
            Assert.AreEqual("v", actual.PPr!.BuChar!.Char);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Star Bullets", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ArrowBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[12].Count);

            var actual = map[12][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuAutoNum);
            Assert.AreEqual("Ø", actual.PPr!.BuChar!.Char);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Arrow Bullets", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_CheckMarkBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[13].Count);

            var actual = map[13][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuAutoNum);
            Assert.AreEqual("ü", actual.PPr!.BuChar!.Char);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Checkmark Bullets", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_OpenParenRightOrderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[14].Count);

            var actual = map[14][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuChar);
            Assert.AreEqual("arabicParenR", actual.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(3, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Open ", actual.R![0].T);
            Assert.AreEqual("Paren", actual.R![1].T);
            Assert.AreEqual(" Right", actual.R![2].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_CapitalRomanNumeralsPeriodOrderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[15].Count);

            var actual = map[15][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuChar);
            Assert.AreEqual("romanUcPeriod", actual.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Roman Numerals", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_CapitalAlphaPeriodOrderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[16].Count);

            var actual = map[16][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuChar);
            Assert.AreEqual("alphaUcPeriod", actual.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Capital Letters", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_LowerCaseAlphaRightParenOrderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[17].Count);

            var actual = map[17][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuChar);
            Assert.AreEqual("alphaLcParenR", actual.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(2, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Lowercase Right ", actual.R![0].T);
            Assert.AreEqual("Paren", actual.R![1].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_LowerCaseAlphaPeriodOrderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[18].Count);

            var actual = map[18][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuChar);
            Assert.AreEqual("alphaLcPeriod", actual.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Lowercase Period", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_LowerCaseRomanNumeralsPeriodOrderedList_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = parser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[19].Count);

            var actual = map[19][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.IsNull(actual.PPr!.BuChar);
            Assert.AreEqual("romanLcPeriod", actual.PPr!.BuAutoNum!.Type);
            Assert.AreEqual(1, actual.R!.Count);
            Assert.AreEqual(0, actual.PPr!.Lvl);
            Assert.AreEqual("Lowercase Roman Numerals", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseHyperlink_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(1, map[9].Count);

            var actual = map[9][0]!;

            Assert.IsNull(actual.A);
            Assert.IsNull(actual.Text);
            Assert.AreEqual("Here a link: https://www.google.com/", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseIndentFollowedByOrdered_ReturnsIntWrapperMap()
        {
            IParser parser = new Parser();
            var map = parser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(6, map[10].Count);

            var actual = map[10];
            Assert.AreEqual("Un", actual[0]!.R![0].T);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_FromMemoryStreamParseUnorderedList_ReturnsIntWrapperMap()
        {
            using var memoryStream = new MemoryStream();
            using var fileStream = File.OpenRead(_path!);
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            var parser = new Parser();
            var map = parser.ParseSpeakerNotes(memoryStream);

            Assert.AreEqual(3, map[5].Count);

            var actualOne = map[5][0]!;

            Assert.IsNull(actualOne.A);
            Assert.IsNull(actualOne.Text);
            Assert.AreEqual(1, actualOne.R!.Count);
            Assert.AreEqual(0, actualOne.PPr!.Lvl);
            Assert.AreEqual("•", actualOne.PPr!.BuChar!.Char);
            Assert.AreEqual("Unordered item 1", actualOne.R![0].T);
            Assert.AreEqual(0, actualOne.R![0].RPr!.B);
            Assert.AreEqual(0, actualOne.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualOne.R![0].RPr!.Lang);

            var actualTwo = map[5][1]!;

            Assert.IsNull(actualTwo.A);
            Assert.IsNull(actualTwo.Text);
            Assert.AreEqual(1, actualTwo.R!.Count);
            Assert.AreEqual(0, actualTwo.PPr!.Lvl);
            Assert.AreEqual("•", actualTwo.PPr!.BuChar!.Char);
            Assert.AreEqual("Unordered item 2", actualTwo.R![0].T);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.B);
            Assert.AreEqual(0, actualTwo.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualTwo.R![0].RPr!.Lang);

            var actualThree = map[5][2]!;

            Assert.IsNull(actualThree.A);
            Assert.IsNull(actualThree.Text);
            Assert.AreEqual(1, actualThree.R!.Count);
            Assert.AreEqual(0, actualThree.PPr!.Lvl);
            Assert.AreEqual("•", actualTwo.PPr!.BuChar!.Char);
            Assert.AreEqual("Unordered item 3", actualThree.R![0].T);
            Assert.AreEqual(0, actualThree.R![0].RPr!.B);
            Assert.AreEqual(0, actualThree.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actualThree.R![0].RPr!.Lang);
        }
    }
}