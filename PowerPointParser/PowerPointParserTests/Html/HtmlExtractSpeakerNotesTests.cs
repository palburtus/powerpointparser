using System.Collections.Generic;
using System.IO;
using Aaks.PowerPointParser.Extensions;
using Aaks.PowerPointParser.Html;
using Aaks.PowerPointParser.Parsers;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PowerPointParserTests.Html;

[TestClass]
public class HtmlExtractSpeakerNotesTests 
{
    private readonly string _directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? string.Empty;
    protected const string ExpectedTestDeckParagraph = @"<p>This note is just a paragraph</p><p>This note is just a paragraph</p><p>And this is a second paragraph</p><p><strong>This is a bold paragraph</strong></p><ul><li>Unordered item 1</li><li>Unordered item 2</li><li>Unordered item 3</li></ul><ul><li>Indent Unordered item 1<ul><li>Indent Unordered item 2<ul><li>Indent Unordered item 3</li></ul></li></ul></li></ul><ol><li>Ordered one</li><li>Ordered two</li><li>Ordered three</li><li>Indent Ordered One<ol><li>Indent Ordered One One<ol><li>Indent Order One OneOne</li></ol></li></ol></li><li>Indent Ordered Three</li></ol><p>Here a link: https://www.google.com/</p><ul><li>Un</li><li>Order</li><li>List</li></ul><ol><li>Followed </li><li>by </li><li>Ordered</li></ol>";
    protected Dictionary<int, string> ExpectedTestDeckOneDict = new()
    {
        {1, "<p>Intro Slide </p><p><strong>Test</strong></p><p><strong>One</strong></p><p><strong>Two</strong></p><ol><li><strong>Order one</strong></li><li><strong>Order two</strong></li><li><strong>Order three</strong></li></ol><p><strong>Here is a note that is </strong>not bold</p>" },
        {2, "<p>Ul 1</p><p>Ul 2</p><p>ul3</p>"},
        {3, "<p>Ask devs for other examples</p>"},
        {4, "<ol><li>Ol 1</li><li>Ol2</li><li>ol3</li></ol>"}
    };
    protected Dictionary<int, string> ExpectedTestDeckParagraphDict = new()
    {
        {1,  ""},
        {2,  "<p>This note is just a paragraph</p>"},
        {3,  "<p>This note is just a paragraph</p><p>And this is a second paragraph</p>"},
        {4,  "<p><strong>This is a bold paragraph</strong></p>"},
        {5,  "<ul><li>Unordered item 1</li><li>Unordered item 2</li><li>Unordered item</li></ul>"},
        {6,  "<ul><li>Indent Unordered item 1<ul><li>Indent Unordered item 2<ul><li>Indent Unordered item 3</li></ul></li></ul></li></ul>" },
        {7,  "<ol><li>Ordered one</li><li>Ordered two</li><li>Ordered three</li></ol>"},
        {8,  "<ol><li>Indent Ordered One<ol><li>Indent Ordered One One<ol><li>Indent Order One OneOne</li></ol></li></ol></li><li>Indent Ordered Three</li></ol>" },
        {9,  "<p>Here a link: https://www.google.com/</p>"},
        {10, "<ul><li>Un</li><li>Order</li><li>List</li></ul><ol><li>Followed </li><li>by </li><li>Ordered</li></ol>"}
    };

   
    [TestMethod]
    [DeploymentItem("TestData")]
    [DataRow("TestDeckParagraph.pptx", ExpectedTestDeckParagraph)]
    public void Test_ExtractSpeakerNotesTest(string fileName, string expected)
    {
        var filePath = Path.Combine(_directory, fileName);
        File.Exists(filePath).Should().Be(true);

        var parser = new PowerPointParser();

        var items = parser.ParseSpeakerNotes(filePath);

        var innerBuilder = new InnerHtmlBuilder();

        var htmlBuilder = new HtmlBuilder(new HtmlListBuilder(innerBuilder), innerBuilder);
        var openXmlParagraphWrappers = items!.ToQueue();


        var htmlStringActual = htmlBuilder.ConvertOpenXmlParagraphWrapperToHtml(openXmlParagraphWrappers!);
        

        htmlStringActual.Should().NotBeEmpty();
        htmlStringActual.Should().Be(expected);
    }
    [TestMethod]
    [DeploymentItem("TestData")]
    [DataRow("TestDeckParagraph.pptx")]
    [DataRow("TestDeckOne.pptx")]
    public void Test_ExtractSpeakerNotesTestAsDictionary(string fileNameTest)
    {

        // Arrange
        var filePath = Path.Combine(_directory, fileNameTest);
        File.Exists(filePath).Should().Be(true); 
        var expectedDict = new Dictionary<string, Dictionary<int, string>>()
        {
            {"TestDeckOne.pptx", ExpectedTestDeckOneDict },
            {"TestDeckParagraph.pptx", ExpectedTestDeckParagraphDict}
        };
        


        var parser = new PowerPointParser();

        // Act
        var items = parser.ParseSpeakerNotes(filePath);
        var innerBuilder = new InnerHtmlBuilder();
        var htmlBuilder = new HtmlBuilder(new HtmlListBuilder(innerBuilder), innerBuilder);
        var htmlPayloadActual = htmlBuilder.ConvertOpenXmlParagraphWrapperToHtml(items);

        // Assert

        htmlPayloadActual.Should().NotBeEmpty();
        
        htmlPayloadActual.Should().BeEquivalentTo(expectedDict[fileNameTest]);
    }

    

}