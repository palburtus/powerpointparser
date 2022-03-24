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
    const string ExpectedTestThree = @"<p>Harbeck N, Penault-Llorca F, Cortes J, et al. Breast cancer. Nat Rev Dis Primers. 2019 Sep 23;5(1):66. doi: 10.1038/s41572-019-0111-2. PMID: 31548545.</p><p>Harbeck N. Risk-adapted adjuvant therapy of luminal early breast cancer in 2020. CurrOpinObstet Gynecol. 2021;33:53-58. doi: 10.1097/GCO.0000000000000679. PMID: 33337614.</p><p>Loibl S, Marmé f, Martin, M, et al. Phase III study of palbociclib combined with endocrine therapy (ET) in patients with hormone-receptor-positive (HR+), HER2-negative primary breast cancer and with high relapse risk after neoadjuvant chemotherapy (NACT): First results from PENELOPE-B . San Antonio Breast Cancer Symposium;December 8-11, 2020; San Antonio, TX. Abstract GS01-02.</p><p>Denkert C, Marmé F, Martin M, et al. Subgroup of post-neoadjuvant luminal-B tumors assessed by HTG in PENELOPE-B investigating palbociclib in high risk HER2-/HR+ breast cancer with residual disease. Presented at: ASCO 2021 Annual Meeting; June 4-8, 2021. Abstract 519.</p><p>Marmé F, Martin M, Untch M, et al. Palbociclib combined with endocrine treatment in breast cancer patients with high relapse risk after neoadjuvant chemotherapy: Subgroup analyses of premenopausal patients in PENELOPE-B. ASCO 2021 Annual Meeting; June 4-8, 2021. Abstract 518. </p><p>Mayer E, et al. PALLAS: A randomized phase III trial of adjuvant palbociclib with endocrine therapy versus endocrine therapy alone for HR+/HER2- early breast cancer. Presented at: European Society for Medical Oncology (ESMO) congress. Virtual 2020. Abstract LBA12.</p><p>DRFS = distant recurrence-free survival</p><p>O’Shaughnessy JA, Johnston S, Harbeck N, et al. Primary outcome analysis of invasive disease-free survival for monarchE: abemaciclib combined with adjuvant endocrine therapy for high risk early breast cancer. San Antonio Breast Cancer Symposium;December 8-11, 2020; San Antonio, TX. Abstract GS01-01.</p>";

    [TestMethod]
    [DeploymentItem("TestData")]
    [DataRow("TestDeckParagraph.pptx")]
    [DataRow("TestThree.pptx")]
    public void Test_ExtractSpeakerNotesTest(string fileName)
    {
        var expected = new Dictionary<string, string>
        {
            { "TestDeckParagraph.pptx", ExpectedTestDeckParagraph },
            { "TestThree.pptx", ExpectedTestThree },
        };
        var filePath = Path.Combine(_directory, fileName);
        File.Exists(filePath).Should().Be(true);

        var parser = new Parser();

        var items = parser.ParseSpeakerNotes(filePath);

        var innerBuilder = new InnerHtmlBuilder();

        var htmlBuilder = new HtmlBuilder(new HtmlListBuilder(innerBuilder), innerBuilder);
        var openXmlParagraphWrappers = items!.ToQueue();


        var htmlStringActual = htmlBuilder.ConvertOpenXmlParagraphWrapperToHtml(openXmlParagraphWrappers!);
        

        htmlStringActual.Should().NotBeEmpty();
        htmlStringActual.Should().Be(expected[fileName]);
    }
}