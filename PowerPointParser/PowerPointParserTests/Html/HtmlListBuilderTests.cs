using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointParser.Html;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParserTests.Html;

namespace PowerPointParser.Html.Tests
{
    [TestClass]
    public class HtmlListBuilderTests : BaseHtmlTests
    {
        private static IHtmlListBuilder _builder;

        [ClassInitialize]
        public static void ClassSetup(TestContext context)
        {
            var innerHtmlBuilder = new InnerHtmlBuilder();
            _builder = new HtmlListBuilder(innerHtmlBuilder);


        }

        [TestMethod]
        public void IsListItem_IsUnorderedList_ReturnsTrue()
        {
            var listItem = BuildUnorderedListItem("hello world");
            bool actual = _builder.IsListItem(listItem);

            Assert.IsTrue(actual);
        }

        [TestMethod]
        public void IsListItem_IsOrderedList_ReturnsTrue()
        {
            var listItem = BuildOrderListItem("hello world");
            bool actual = _builder.IsListItem(listItem);

            Assert.IsTrue(actual);
        }

        [TestMethod]
        public void IsListItem_IsParagraph_ReturnsFalse()
        {
            var listItem = BuildParagraphLine("hello world");
            bool actual = _builder.IsListItem(listItem);

            Assert.IsFalse(actual);
        }
    }
}