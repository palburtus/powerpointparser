using Microsoft.VisualStudio.TestTools.UnitTesting;
using Aaks.PowerPointParser.Html;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aaks.PowerPointParser.Dto;
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

            Assert.AreEqual("<p>hello world</p>", result);
        }

        [TestMethod]
        public void BuildInnerHtmlListItem_PlainText_ReturnsString()
        {
            IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
            var current = BuildParagraphLine("hello world", new RPr { B = 1 });
            var result = innerHtmlBuilder.BuildInnerHtmlListItem(current);

            Assert.AreEqual("<li><strong>hello world</strong></li>", result);
        }

        [TestMethod]
        public void BuildInnerHtmlParagraph_BoldText_ReturnsString()
        {
            IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
            var current = BuildParagraphLine("hello world", new RPr { B = 1 });
            var result = innerHtmlBuilder.BuildInnerHtmlParagraph(current);

            Assert.AreEqual("<p><strong>hello world</strong></p>", result);
        }

        [TestMethod]
        public void BuildInnerHtmlListItem_BoldText_ReturnsString()
        {
            IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
            var current = BuildParagraphLine("hello world");
            var result = innerHtmlBuilder.BuildInnerHtmlListItem(current);

            Assert.AreEqual("<li>hello world</li>", result);
        }
    }
}