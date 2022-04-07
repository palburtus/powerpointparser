﻿using System.Collections.Generic;
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

        #region Notes Parsing

        [TestMethod]
        [DeploymentItem("TestData/TestDeckOne.pptx")]
        public void Parse_ParseTestDeck_ReturnsOpenXmlLineItem()
        {
            var path = Path.Combine(_directory!, "TestDeckOne.pptx");

            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(path);

            map.Keys.Count.Should().Be(4);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseNoteParagraph_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            var actual = map[2][0]!;

            var expected = BuildParagraphTextWrapper("This note is just a paragraph");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestThree.pptx")]
        public void Parse_ParseAlternativeFormat_ReturnsOpenXmlLineItem()
        {
            var path = Path.Combine(_directory!, "TestThree.pptx");
            IPowerPointParser powerPointParser = new PowerPointParser();

            var map = powerPointParser.ParseSpeakerNotes(path);

            var actual = map[8][0]!;

            actual.R![0].T.Should().Be("Caricas");
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseNoteConsecutiveParagraphs_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            var actual = map[3][1]!;
            var expected = BuildParagraphTextWrapper("And this is a second paragraph");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseBoldParagraph_ReturnsOpenXmlLineItem()
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
        public void Parse_ParseItalicParagraph_ReturnsOpenXmlLineItem()
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
        public void Parse_ParseUnderlinedParagraph_ReturnsOpenXmlLineItem()
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
        public void Parse_ParseStrikeThroughParagraph_ReturnsOpenXmlLineItem()
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
        public void Parse_ParseCenterAlignedParagraph_ReturnsOpenXmlLineItem()
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
        public void Parse_ParseRightAlignedParagraph_ReturnsOpenXmlLineItem()
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
        public void Parse_ParseJustifiedAlignedParagraph_ReturnsOpenXmlLineItem()
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
        public void Parse_ParseConsecutiveEmptySpaces_ReturnsOpenXmlLineItem()
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
            var expectedEmpty = new OpenXmlLineItem
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
        public void Parse_ParseUnorderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(3, map[5].Count);

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
        public void Parse_ParseEmbeddedUnorderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            var actualOne = map[6][0]!;
            var expectedOne = BuildUlTextWrapper("Indent Unordered item 1");

            var actualTwo = map[6][1]!;
            var expectedTwo = BuildUlTextWrapper("Indent Unordered item 2", 1);

            var actualThree = map[6][2]!;
            var expectedThree = BuildUlTextWrapper("Indent Unordered item 3", 2);

            map[6].Count.Should().Be(3);
            actualOne.Should().BeEquivalentTo(expectedOne);
            actualTwo.Should().BeEquivalentTo(expectedTwo);
            actualThree.Should().BeEquivalentTo(expectedThree);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseOrderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(3, map[7].Count);

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
        public void Parse_ParseEmbeddedOrderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(4, map[8].Count);

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
        public void Parse_HollowRoundBulletsUnorderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[9].Count);

            var actual = map[9][0]!;
            var expected = BuildUlTextWrapper("Filled Square Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("§");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_HollowSquareBulletsUnorderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[10].Count);

            var actual = map[10][0]!;
            var expected = BuildUlTextWrapper("Hollow Square Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("q");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_StarBulletsUnorderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[11].Count);

            var actual = map[11][0]!;
            var expected = BuildUlTextWrapper("Star Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("v");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_ArrowBulletsUnorderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[12].Count);

            var actual = map[12][0]!;
            var expected = BuildUlTextWrapper("Arrow Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("Ø");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_CheckMarkBulletsUnorderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[13].Count);

            var actual = map[13][0]!;
            var expected = BuildUlTextWrapper("Checkmark Bullets");
            expected.PPr = BuildSpecialUlCharacterDefaultPPr("ü");

            actual.Should().BeEquivalentTo(expected);

        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_OpenParenRightOrderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[14].Count);

            var actual = map[14][0]!;
            var expected = BuildOlTextWrapper("Open ");
            expected.R!.Add(BuildRItem("Paren"));
            expected.R!.Add(BuildRItem(" Right"));
            expected.PPr = BuildSpecialOlDefaultPPr("arabicParenR");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_CapitalRomanNumeralsPeriodOrderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[15].Count);

            var actual = map[15][0]!;

            var expected = BuildOlTextWrapper("Roman Numerals");
            expected.PPr = BuildSpecialOlDefaultPPr("romanUcPeriod");
            expected.PPr.MarL = 285750;
            expected.PPr.Indent = -285750;

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_CapitalAlphaPeriodOrderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[16].Count);

            var actual = map[16][0]!;

            var expected = BuildOlTextWrapper("Capital Letters");
            expected.PPr = BuildSpecialOlDefaultPPr("alphaUcPeriod");
            expected.PPr.MarL = 285750;
            expected.PPr.Indent = -285750;

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestFour.pptx")]
        public void Parse_LowerCaseAlphaRightParenOrderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[17].Count);

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
        public void Parse_LowerCaseAlphaPeriodOrderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[18].Count);

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
        public void Parse_LowerCaseRomanNumeralsPeriodOrderedList_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var path = Path.Combine(_directory!, "TestFour.pptx");
            var map = powerPointParser.ParseSpeakerNotes(path);

            Assert.AreEqual(2, map[19].Count);

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
        public void Parse_ParseHyperlink_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(1, map[9].Count);

            var actual = map[9][0]!;

            var expected = BuildParagraphTextWrapper("Here a link: https://www.google.com/");

            actual.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_ParseIndentFollowedByOrdered_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser powerPointParser = new PowerPointParser();
            var map = powerPointParser.ParseSpeakerNotes(_path!);

            Assert.AreEqual(6, map[10].Count);

            var actual = map[10];

            actual[0]!.R![0].T.Should().Be("Un");
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void Parse_FromMemoryStreamParseUnorderedList_ReturnsOpenXmlLineItem()
        {
            using var memoryStream = new MemoryStream();
            using var fileStream = File.OpenRead(_path!);
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            var parser = new PowerPointParser();
            var map = parser.ParseSpeakerNotes(memoryStream);

            Assert.AreEqual(3, map[5].Count);

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

        #endregion

        #region Slide Parsing

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void ParseSlide_ParseTitleSlideSubtitle_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser parser = new PowerPointParser();
            var slideDictionary = parser.ParseSlide(_path!);

            slideDictionary.Should().NotBeNull();

            var actual = slideDictionary[2][0]!.CSld.SpTree.Sp[2].TxBody.P.R.T;

            actual.Should().Be("Sit Dolor Amet");
        }

        [TestMethod]
        [DeploymentItem("TestData/TestDeckParagraph.pptx")]
        public void ParseSlide_ParseTitleSlideTitle_ReturnsOpenXmlLineItem()
        {
            IPowerPointParser parser = new PowerPointParser();
            var slideDictionary = parser.ParseSlide(_path!);

            slideDictionary.Should().NotBeNull();

            var actual = slideDictionary[2][0]!.CSld.SpTree.Sp[1].TxBody.P.R.T;

            actual.Should().Be("Title Lorem Ipsum");
        }

        #endregion


    }
}