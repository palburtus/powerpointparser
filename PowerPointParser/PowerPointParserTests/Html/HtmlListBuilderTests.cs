using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointParser.Html;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser.Dto;
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

        [TestMethod]
        public void BuildList_OrderedPreviousNestedCurrentUnorderedNormalLastNormal_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var three = BuildOrderListItem("three", level: 1);
            var four = BuildUnorderedListItem("hello world");
            var five = BuildUnorderedListItem("goodbye world");

            var actual = _builder.BuildList(three, four, five);

            Assert.AreEqual("</ol><ul><li>hello world</li>", actual);
        }

        [TestMethod]
        public void BuildList_UnOrderedPreviousCurrentLastNormal_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var four = BuildUnorderedListItem("hello world");
            var five = BuildUnorderedListItem("goodbye world");
            var six = BuildUnorderedListItem("test world");

            var actual = _builder.BuildList(four, five, six);

            Assert.AreEqual("<li>goodbye world</li>", actual);
        }

        [TestMethod]
        public void BuildList_UnOrderedPreviousCurrentNormalLastNull_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var five = BuildUnorderedListItem("goodbye world");
            var six = BuildUnorderedListItem("test world");

            var actual = _builder.BuildList(five, six, null);

            Assert.AreEqual("<li>test world</li></ol>", actual);
        }
        
        [TestMethod]
        public void BuildList_OrderedPreviousNullCurrentNormalLastNormal_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var one = BuildOrderListItem("one");
            var two = BuildOrderListItem("two");
            
            var actual = _builder.BuildList(null, one, two);

            Assert.AreEqual("<ol><li>one</li>", actual);
        }

        [TestMethod]
        public void BuildList_OrderedPreviousNormalCurrentNormalLastNested_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var one = BuildOrderListItem("one");
            var two = BuildOrderListItem("two");
            var three = BuildOrderListItem("three", level: 1);
            
            var actual = _builder.BuildList(one, two, three);

            Assert.AreEqual("<li>two</li>", actual);
        }

        [TestMethod]
        public void BuildList_OrderedPreviousNormalCurrentNestedLastNormal_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var two = BuildOrderListItem("two");
            var three = BuildOrderListItem("three", level: 1);
            var four = BuildUnorderedListItem("hello world");
            
            var actual = _builder.BuildList(two, three, four);

            Assert.AreEqual("<ol><li>three</li></ol>", actual);
        }

        [TestMethod]
        public void BuildList_NullPreviousOrderedCurrentNestedUnorderedLast_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var one = BuildOrderListItem("three", level: 1);
            var two = BuildUnorderedListItem("hello world");

            var actual = _builder.BuildList(one, two, null);

            Assert.AreEqual("<li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void BuildList_NullPreviousNestedUnorderedCurrentLast_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var two = BuildOrderListItem("one");
            var three = BuildOrderListItem("two");


            var actual = _builder.BuildList(null, two, three);

            Assert.AreEqual("<ol><li>one</li>", actual);
        }

        [TestMethod]
        public void BuildList_OrderedPreviousCurrentNestedUnorderedLast_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var one = BuildOrderListItem("one");
            var two = BuildOrderListItem("two");
            var three = BuildUnorderedListItem("hello world", level: 1);


            var actual = _builder.BuildList(one, two, three);

            Assert.AreEqual("<li>two</li>", actual);
        }

        [TestMethod]
        public void BuildList_OrderedPreviousNestedUnorderedCurrentLast_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var one = BuildOrderListItem("two");
            var two = BuildUnorderedListItem("hello world", level: 1);
            var three = BuildUnorderedListItem("goodbye world", level: 1);

            var actual = _builder.BuildList(one, two, three);

            Assert.AreEqual("<ul><li>hello world</li>", actual);
        }

        [TestMethod]
        public void BuildList_NestedUnOrderedPreviousCurrentLast_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var one = BuildUnorderedListItem("hello world", level: 1);
            var two = BuildUnorderedListItem("goodbye world", level: 1);
            var three = BuildUnorderedListItem("test world", level: 1);

            var actual = _builder.BuildList(one, two, three);

            Assert.AreEqual("<li>goodbye world</li>", actual);
        }

        [TestMethod]
        public void BuildList_NestedUnOrderedPreviousCurrentUnorderedLast_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var one = BuildUnorderedListItem("goodbye world", level: 1);
            var two = BuildUnorderedListItem("test world", level: 1);
            var three = BuildOrderListItem("three");

            var actual = _builder.BuildList(one, two, three);

            Assert.AreEqual("<li>test world</li>", actual);
        }

        [TestMethod]
        public void BuildList_NestedUnOrderedPreviousCurrentNullLast_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            var one = BuildUnorderedListItem("test world", level: 1);
            var two = BuildOrderListItem("three");
            

            var actual = _builder.BuildList(one, two, null);

            Assert.AreEqual("<li>three</li>", actual);
        }

       
    }
}