using System.Collections.Generic;
using System.IO;
using Aaks.PowerPointParser;
using Aaks.PowerPointParser.Html;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PowerPointParserTests.Html;

[TestClass]
public class HtmlExtractSpeakerNotesTests 
{
    private readonly string _directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? string.Empty;
    const string ExpectedTestDeckParagraph = @"<p>This note is just a paragraph</p><p>This note is just a paragraph</p><p>And this is a second paragraph</p><p><strong>This is a bold paragraph</strong></p><li>Unordered item 1</li><li>Unordered item 2</li><li>Unordered item 3</li><li>Indent Unordered item 1</li><ul><li>Indent Unordered item 2</li><ul><li>Indent Unordered item 3</li></ul></ul></ul><ol><li>Ordered one</li><li>Ordered two</li><li>Ordered three</li><li>Indent Ordered One</li><ol><li>Indent Ordered One One</li><ol><li>Indent Order One OneOne</li></ol></ol><li>Indent Ordered Three</li><p>Here a link: https://www.google.com/</p><li>Un</li><li>Order</li><li>List</li></ul><ol><li>Followed </li><li>by </li><li>Ordered</li></ol>";

    [TestMethod]
    [DeploymentItem("TestData")]
    [DataRow("TestDeckParagraph.pptx", ExpectedTestDeckParagraph)]
    public void Test_ExtractSpeakerNotesTest(string fileName, string expected)
    {
        var filePath = Path.Combine(_directory, fileName);
        File.Exists(filePath).Should().Be(true);

        var parser = new Parser();

        var items = parser.ParseSpeakerNotes(filePath);

        var innerBuilder = new InnerHtmlBuilder();

        var htmlBuilder = new HtmlBuilder(new HtmlListBuilder(innerBuilder), innerBuilder);
        var openXmlParagraphWrappers = items!.ToQueue();


        var htmlStringActual = htmlBuilder.ConvertOpenXmlParagraphWrapperToHtml(openXmlParagraphWrappers!);
        

        htmlStringActual.Should().NotBeEmpty();
        htmlStringActual.Should().Be(expected);
    }
}