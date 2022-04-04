using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointParserTests.Html;

namespace Aaks.PowerPointParser.Html.Tests
{
    [TestClass]
    public class HtmlListBuilderTests : BaseHtmlTests
    {
        private static IHtmlListBuilder _builder = null!;

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

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void IsListItem_IsOrderedList_ReturnsTrue()
        {
            var listItem = BuildOrderListItem("hello world");
            bool actual = _builder.IsListItem(listItem);

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void IsListItem_IsParagraph_ReturnsFalse()
        {
            var listItem = BuildParagraphLine("hello world");
            bool actual = _builder.IsListItem(listItem);

            actual.Should().BeFalse();
        }

        [TestMethod]
        public void BuildList_UnOrderedPreviousCurrentLastNormal_ReturnsString()
        {
            var four = BuildUnorderedListItem("hello world");
            var five = BuildUnorderedListItem("goodbye world");
            var six = BuildUnorderedListItem("test world");

            var actual = _builder.BuildList(four, five, six);

            actual.Should().Be("<li>goodbye world</li>");
        }

        [TestMethod]
        public void BuildList_UnOrderedPreviousCurrentNormalLastNull_ReturnsString()
        {
            var five = BuildUnorderedListItem("goodbye world");
            var six = BuildUnorderedListItem("test world");

            var actual = new HtmlListBuilder(new InnerHtmlBuilder()).BuildList(five, six, null);

            actual.Should().Be("<li>test world</li>");
        }
        
        [TestMethod]
        public void BuildList_OrderedPreviousNullCurrentNormalLastNormal_ReturnsString()
        {
            var one = BuildOrderListItem("one");
            var two = BuildOrderListItem("two");
            
            var actual = _builder.BuildList(null, one, two);

            actual.Should().Be("<ol><li>one</li>");
        }

        [TestMethod]
        public void BuildList_OrderedPreviousNormalCurrentNormalLastNested_ReturnsString()
        {
            var one = BuildOrderListItem("one");
            var two = BuildOrderListItem("two");
            var three = BuildOrderListItem("three", level: 1);
            
            var actual = _builder.BuildList(one, two, three);

            actual.Should().Be("<li>two");
        }

        [TestMethod]
        public void BuildList_OrderedPreviousNormalCurrentNestedLastNormal_ReturnsString()
        {
            var two = BuildOrderListItem("two");
            var three = BuildOrderListItem("three", level: 1);
            var four = BuildUnorderedListItem("hello world");
            
            var actual = _builder.BuildList(two, three, four);

            actual.Should().Be("<ol><li>three</li></ol></li>");
        }

        [TestMethod]
        public void BuildList_NullPreviousOrderedCurrentNestedUnorderedLast_ReturnsString()
        {
            var one = BuildOrderListItem("three", level: 1);
            var two = BuildUnorderedListItem("hello world");

            var actual = _builder.BuildList(one, two, null);

            actual.Should().Be("<li>hello world</li></ol>");
        }

        [TestMethod]
        public void BuildList_NullPreviousNestedUnorderedCurrentLast_ReturnsString()
        {
            var two = BuildOrderListItem("one");
            var three = BuildOrderListItem("two");


            var actual = _builder.BuildList(null, two, three);

            actual.Should().Be("<ol><li>one</li>");
        }

        [TestMethod]
        public void BuildList_OrderedPreviousCurrentNestedUnorderedLast_ReturnsString()
        {
            var one = BuildOrderListItem("one");
            var two = BuildOrderListItem("two");
            var three = BuildUnorderedListItem("hello world", level: 1);


            var actual = _builder.BuildList(one, two, three);

            actual.Should().Be("<li>two");
        }

        [TestMethod]
        public void BuildList_OrderedPreviousNestedUnorderedCurrentLast_ReturnsString()
        {
            var one = BuildOrderListItem("two");
            var two = BuildUnorderedListItem("hello world", level: 1);
            var three = BuildUnorderedListItem("goodbye world", level: 1);

            var actual = _builder.BuildList(one, two, three);

            actual.Should().Be("<ul><li>hello world</li>");
        }

        [TestMethod]
        public void BuildList_NestedUnOrderedPreviousCurrentLast_ReturnsString()
        {
            var one = BuildUnorderedListItem("hello world", level: 1);
            var two = BuildUnorderedListItem("goodbye world", level: 1);
            var three = BuildUnorderedListItem("test world", level: 1);

            var actual = _builder.BuildList(one, two, three);

            actual.Should().Be("<li>goodbye world</li>");
        }
    }
}