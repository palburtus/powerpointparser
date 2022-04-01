using Microsoft.VisualStudio.TestTools.UnitTesting;
using Aaks.PowerPointParser.Dto;
using FluentAssertions;
using PowerPointParserTests.Html;

namespace Aaks.PowerPointParser.Html.Tests
{
    [TestClass]
    public class InnerHtmlBuilderTests : BaseHtmlTests
    {
        [TestMethod]
        public void BuildInnerHtmlParagraph_PlainText_ReturnsString()
        {
            IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
            var current = BuildParagraphLine("hello world");
            var result = innerHtmlBuilder.BuildInnerHtmlParagraph(current);

            result.Should().Be("<p>hello world</p>");
        }

        [TestMethod]
        public void BuildInnerHtmlListItem_PlainText_ReturnsString()
        {
            IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
            var current = BuildParagraphLine("hello world", new RPr { B = 1 });
            var result = innerHtmlBuilder.BuildInnerHtmlListItem(current);

            result.Should().Be("<li><strong>hello world</strong></li>");
        }

        [TestMethod]
        public void BuildInnerHtmlParagraph_BoldText_ReturnsString()
        {
            IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
            var current = BuildParagraphLine("hello world", new RPr { B = 1 });
            var result = innerHtmlBuilder.BuildInnerHtmlParagraph(current);

            result.Should().Be("<p><strong>hello world</strong></p>");
        }

        [TestMethod]
        public void BuildInnerHtmlListItem_BoldText_ReturnsString()
        {
            IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
            var current = BuildParagraphLine("hello world");
            var result = innerHtmlBuilder.BuildInnerHtmlListItem(current);

            result.Should().Be("<li>hello world</li>");
        }


        [TestMethod]
        public void BuildInnerHtmlParagraph_EncodingCharacters_ReturnsString()
        {
            IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
            var current = BuildParagraphLine("\"hello & world <one> 'two' \"");
            var result = innerHtmlBuilder.BuildInnerHtmlParagraph(current);

            result.Should().Be("<p>&quot;hello &amp; world &lt;one&gt; &#39;two&#39; &quot;</p>");
        }

        [TestMethod]
        public void BuildInnerHtmlListItem_EncodingCharacters_ReturnsString()
        {
            IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
            var current = BuildParagraphLine("\"hello & world <one> 'two' \"");
            var result = innerHtmlBuilder.BuildInnerHtmlListItem(current);

            result.Should().Be("<li>&quot;hello &amp; world &lt;one&gt; &#39;two&#39; &quot;</li>");
        }
    }
}