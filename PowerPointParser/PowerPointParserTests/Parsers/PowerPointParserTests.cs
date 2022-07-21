using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Aaks.PowerPointParser.Dto;
using FluentAssertions;
using PowerPointParserTests.Html;

// ReSharper disable once CheckNamespace - Test namespaces should match production
namespace Aaks.PowerPointParser.Parsers.Tests
{
    [TestClass]
    public class PowerPointParserTests : BaseOpenXmlParserTests
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

            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(path);

            map.Keys.Count.Should().Be(4);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseNoteParagraph_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            var actual = map[2][0]!;

            var expected = BuildParagraphTextWrapper("This note is just a paragraph");

            actual.Should().BeEquivalentTo(expected);
        }
        
        [TestMethod]
        [DeploymentItem("TestData/TestThree.pptx")]
        public void Parse_ParseAlternativeFormat_ReturnsIntWrapperMap()
        {
            var path = Path.Combine(_directory!, "TestThree.pptx");
            IPowerPointParser powerPointParser = new PowerPointParser();

            var map = powerPointParser.ParseSpeakerNotes(path);

            var actual = map[8][0]!;

            actual.R![0].T.Should().Be("Caricas");
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseNoteConsecutiveParagraphs_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            var actual = map[3][1]!;
            var expected = BuildParagraphTextWrapper("And this is a second paragraph");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseBoldParagraph_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            var actual = map[4][0]!;
            var expected = BuildParagraphTextWrapper("This is a bold paragraph");
            expected.R![0].RPr!.B = 1;

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseItalicParagraph_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            var actual = map[1][0]!;
            var expected = BuildParagraphTextWrapper("This is an italic paragraph");
            expected.R![0].RPr!.I = 1;
            expected.PPr = BuildDefaultParagraphPpr();
            expected.PPr.SpcBef = BuildDefaultSpcBef();
            expected.PPr.SpcAft = BuildDefaultSpcAft();
            expected.PPr.BuNone = new object();

            actual.Should().BeEquivalentTo(expected);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseUnderlinedParagraph_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            var actual = map[2][0]!;
            var expected = BuildParagraphTextWrapper("This is underlined");
            expected.R![0].RPr!.U = "sng";

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseStrikeThroughParagraph_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            var actual = map[3][0]!;
            var expected = BuildParagraphTextWrapper("This text has a strike through");
            expected.R![0].RPr!.Strike = "sngStrike";

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseCenterAlignedParagraph_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            var actual = map[4][0]!;

            var expected = BuildParagraphTextWrapper("This text is center aligned");
            expected.PPr = BuildDefaultParagraphPpr();
            expected.PPr.Algn = "ctr";

            actual.Should().BeEquivalentTo(expected);
        }
     
        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseRightAlignedParagraph_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            var actual = map[5][0]!;
            var expected = BuildParagraphTextWrapper("This text is right aligned");
            expected.PPr = BuildDefaultParagraphPpr();
            expected.PPr.Algn = "r";

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseJustifiedAlignedParagraph_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            var actual = map[6][0]!;
            var expected = BuildParagraphTextWrapper("This text is aligned justified");
            expected.PPr = BuildDefaultParagraphPpr();
            expected.PPr.Algn = "just";

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ParseConsecutiveEmptySpaces_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            var expectedPpr = BuildDefaultParagraphPpr();
            expectedPpr.BuFont = new BuFont { Typeface = "+mj-lt" };
            expectedPpr.BuNone = new object();
            expectedPpr.Algn = null;

            var actualOne = map[7][0]!;
            var expectedOne = BuildParagraphTextWrapper("Paragraph One");
            expectedOne.PPr = expectedPpr;

            var actualTwo = map[7][1]!;
            var expectedEmpty = new OpenXmlTextWrapper
            {
                R = new List<R>(),
                PPr = expectedPpr
            };

            var actualThree = map[7][2]!;

            var actualFour = map[7][3]!;
            var expectedFour = BuildParagraphTextWrapper("Paragraph Two after two spaces");
            expectedFour.PPr = expectedPpr;

            var actualFive = map[7][4]!;

            var actualSix = map[7][5]!;

            var actualSeven = map[7][6]!;
            var expectedSeven = BuildParagraphTextWrapper("Paragraph Three after three spaces");
            expectedSeven.PPr = expectedPpr;

            actualOne.Should().BeEquivalentTo(expectedOne);
            actualTwo.Should().BeEquivalentTo(expectedEmpty);
            actualThree.Should().BeEquivalentTo(expectedEmpty);
            actualFour.Should().BeEquivalentTo(expectedFour);
            actualFive.Should().BeEquivalentTo(expectedEmpty);
            actualSix.Should().BeEquivalentTo(expectedEmpty);
            actualSeven.Should().BeEquivalentTo(expectedSeven);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseUnorderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(4, map[5].Count);

            var actualOne = map[5][0]!;
            var expectedOne = BuildUlTextWrapper("Unordered item 1");

            var actualTwo = map[5][1]!;
            var expectedTwo = BuildUlTextWrapper("Unordered item 2");

            var actualThree = map[5][2]!;
            var expectedThree = BuildUlTextWrapper("Unordered item 3");

            actualOne.Should().BeEquivalentTo(expectedOne);
            actualTwo.Should().BeEquivalentTo(expectedTwo);
            actualThree.Should().BeEquivalentTo(expectedThree);

        }
        
        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseEmbeddedUnorderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            var actualOne = map[6][0]!;
            var expectedOne = BuildUlTextWrapper("Indent Unordered item 1");

            var actualTwo = map[6][1]!;
            var expectedTwo = BuildUlTextWrapper("Indent Unordered item 2", 1);

            var actualThree = map[6][2]!;
            var expectedThree = BuildUlTextWrapper("Indent Unordered item 3", 2);

            map[6].Count.Should().Be(4);
            actualOne.Should().BeEquivalentTo(expectedOne);
            actualTwo.Should().BeEquivalentTo(expectedTwo);
            actualThree.Should().BeEquivalentTo(expectedThree);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseOrderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(4, map[7].Count);

            var actualOne = map[7][0]!;
            var expectedOne = BuildOlTextWrapper("Ordered one");
            
            var actualTwo = map[7][1]!;
            var expectedTwo = BuildOlTextWrapper("Ordered two");
            
            var actualThree = map[7][2]!;
            var expectedThree = BuildOlTextWrapper("Ordered three");

            actualOne.Should().BeEquivalentTo(expectedOne);
            actualTwo.Should().BeEquivalentTo(expectedTwo);
            actualThree.Should().BeEquivalentTo(expectedThree);
        }
        

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseEmbeddedOrderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(5, map[8].Count);

            var actualOne = map[8][0]!;
            var expectedOne = BuildOlTextWrapper("Indent Ordered One");

            

            var actualTwo = map[8][1]!;
            var expectedTwo = BuildOlTextWrapper("Indent Ordered One ", 1);
            expectedTwo.R!.Add(BuildRItem("One"));

            

            var actualThree = map[8][2]!;
            var expectedThree = BuildOlTextWrapper("Indent Order One ", 2);
            expectedThree.R!.Add(BuildRItem("One"));
            expectedThree.R!.Add(BuildRItem(string.Empty));
            expectedThree.R!.Add(BuildRItem("One"));
            
            var actualFour = map[8][3]!;
            var expectedFour = BuildOlTextWrapper("Indent Ordered Three");

            actualOne.Should().BeEquivalentTo(expectedOne);
            actualTwo.Should().BeEquivalentTo(expectedTwo);
            actualThree.Should().BeEquivalentTo(expectedThree);
            actualFour.Should().BeEquivalentTo(expectedFour);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_HollowRoundBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[9].Count);

            var actual = map[9][0]!;
            var expected = BuildUlTextWrapper("Filled Square Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("§");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_HollowSquareBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[10].Count);

            var actual = map[10][0]!;
            var expected = BuildUlTextWrapper("Hollow Square Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("q");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_StarBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[11].Count);

            var actual = map[11][0]!;
            var expected = BuildUlTextWrapper("Star Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("v");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ArrowBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[12].Count);

            var actual = map[12][0]!;
            var expected = BuildUlTextWrapper("Arrow Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("Ø");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_CheckMarkBulletsUnorderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[13].Count);

            var actual = map[13][0]!;
            var expected = BuildUlTextWrapper("Checkmark Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("ü");

            actual.Should().BeEquivalentTo(expected);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_OpenParenRightOrderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[14].Count);

            var actual = map[14][0]!;
            var expected = BuildOlTextWrapper("Open ");
            expected.R!.Add(BuildRItem("Paren"));
            expected.R!.Add(BuildRItem(" Right"));
            expected.PPr = BuildSpecialOlDefaultPPr("arabicParenR");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_CapitalRomanNumeralsPeriodOrderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[15].Count);

            var actual = map[15][0]!;

            var expected = BuildOlTextWrapper("Roman Numerals");
            expected.PPr = BuildSpecialOlDefaultPPr("romanUcPeriod");
            expected.PPr.MarL = 285750;
            expected.PPr.Indent = -285750;

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_CapitalAlphaPeriodOrderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[16].Count);

            var actual = map[16][0]!;

            var expected = BuildOlTextWrapper("Capital Letters");
            expected.PPr = BuildSpecialOlDefaultPPr("alphaUcPeriod");
            expected.PPr.MarL = 285750;
            expected.PPr.Indent = -285750;

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_LowerCaseAlphaRightParenOrderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[17].Count);

            var actual = map[17][0]!;

            var expected = BuildOlTextWrapper("Lowercase Right ");
            expected.R!.Add(BuildRItem("Paren"));
            expected.PPr = BuildSpecialOlDefaultPPr("alphaLcParenR");
            expected.PPr.MarL = 228600;
            expected.PPr.Indent = -228600;
            expected.PPr.Algn = null;
            expected.PPr.SpcAft = null;
            expected.PPr.SpcBef = null;

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_LowerCaseAlphaPeriodOrderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[18].Count);

            var actual = map[18][0]!;

            var expected = BuildOlTextWrapper("Lowercase Period");
            expected.PPr = BuildSpecialOlDefaultPPr("alphaLcPeriod");
            expected.PPr.MarL = 228600;
            expected.PPr.Indent = -228600;
            expected.PPr.Algn = null;
            expected.PPr.SpcAft = null;
            expected.PPr.SpcBef = null;

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_LowerCaseRomanNumeralsPeriodOrderedList_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(3, map[19].Count);

            var actual = map[19][0]!;

            var expected = BuildOlTextWrapper("Lowercase Roman Numerals");
            expected.PPr = BuildSpecialOlDefaultPPr("romanLcPeriod");
            expected.PPr.MarL = 285750;
            expected.PPr.Indent = -285750;
            expected.PPr.Algn = null;
            expected.PPr.SpcAft = null;
            expected.PPr.SpcBef = null;

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseHyperlink_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(2, map[9].Count);

            var actual = map[9][0]!;

            var expected = BuildParagraphTextWrapper("Here a link: https://www.google.com/");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseIndentFollowedByOrdered_ReturnsIntWrapperMap()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(7, map[10].Count);

            var actual = map[10];
            
            actual[0]!.R![0].T.Should().Be("Un");
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_FromMemoryStreamParseUnorderedList_ReturnsIntWrapperMap()
        {
            using var memoryStream = new MemoryStream();
            using var fileStream = File.OpenRead(_path!);
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            var parser = new PowerPointParser();
            var map = parser.ParseSpeakerNotes(memoryStream);

            Assert.AreEqual(4, map[5].Count);

            var actualOne = map[5][0]!;
            var expectedOne = BuildUlTextWrapper("Unordered item 1");

           
            var actualTwo = map[5][1]!;
            var expectedTwo = BuildUlTextWrapper("Unordered item 2");

          
            var actualThree = map[5][2]!;
            var expectedThree = BuildUlTextWrapper("Unordered item 3");

            actualOne.Should().BeEquivalentTo(expectedOne);
            actualTwo.Should().BeEquivalentTo(expectedTwo);
            actualThree.Should().BeEquivalentTo(expectedThree);
        }

        [TestMethod]
        [DeploymentItem("TestData/Malformed.pptx")]
        public void ParseSlide_ParseSlidesMailformedUrl_ReturnsOpenXmlLineItem()
        {
            var path = Path.Combine(_directory!, "Malformed.pptx");
            using var memoryStream = new MemoryStream();
            using var fileStream = File.OpenRead(path!);
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(memoryStream);

            map.Keys.Count.Should().Be(2);

            var actual = map[1];

            actual[0]!.R![0].T.Should().Be("Figure ");
            actual[0]!.R![1].T.Should().Be("Legend");
            actual[0]!.R![2].T.Should().Be(":");

        }

        [TestMethod]
        [DeploymentItem("TestData/ExtraNoteObject.pptx")]
        public void ParseSpeakerNotes_ParsesSpeakerNotesInPresentationWithExtraUnusuedNoteObject_ReturnsOpenLineItemXml()
        {
            var path = Path.Combine(_directory!, "ExtraNoteObject.pptx");
            using var memoryStream = new MemoryStream();
            using var fileStream = File.OpenRead(path!);
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(memoryStream);

            map.Keys.Count.Should().Be(7);

            var actual = map[1];

            actual[0]!.R!.Count.Should().Be(0);
            actual[1]!.R![0].T.Should().Be("The");
            actual[1]!.R![1].T.Should().Be(".");
        }

        [TestMethod]
        [DeploymentItem("TestData/SpecialCharacters.pptx")]
        public void ParseSpeakerNotes_ParseSpecialCharacters_ReturnsOpenLineItemXml()
        {
            var path = Path.Combine(_directory!, "SpecialCharacters.pptx");
            using var memoryStream = new MemoryStream();
            using var fileStream = File.OpenRead(path!);
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(memoryStream);

            map.Keys.Count.Should().Be(1);

            var actual = map[1];

            //actual[0]!.R!.Count.Should().Be(37);
            actual[12]!.R![0].T.Should().Be("Google");
            actual[12]!.R![0].RPr!.Strike.Should().Be("dblStrike");
        }

        [TestMethod]
        [DeploymentItem("TestData/SpecialCharacters.pptx")]
        public void ParseSpeakerNotes_ParseDoubleStrikeThrough_ReturnsOpenLineItemXml()
        {
            var path = Path.Combine(_directory!, "SpecialCharacters.pptx");
            using var memoryStream = new MemoryStream();
            using var fileStream = File.OpenRead(path!);
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(memoryStream);

            map.Keys.Count.Should().Be(1);

            var actual = map[1];
            
            actual[12]!.R![0].T.Should().Be("Google");
            actual[12]!.R![0].RPr!.Strike.Should().Be("dblStrike");
        }
        
        [TestMethod]
        [DeploymentItem("TestData/Indentation.pptx")]
        public void ParseSpeakerNotes_Indentation_ReturnsOpenLineItemXml()
        {
            var path = Path.Combine(_directory!, "Indentation.pptx");
            using var memoryStream = new MemoryStream();
            using var fileStream = File.OpenRead(path!);
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(memoryStream);

            map.Keys.Count.Should().Be(1);

            var actual = map[1];

            //actual[12]!.R![0].T.Should().Be("Google");
            //actual[12]!.R![0].RPr!.Strike.Should().Be("dblStrike");
        }
    }
}