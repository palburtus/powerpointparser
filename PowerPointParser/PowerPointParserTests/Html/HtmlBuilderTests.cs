using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Text;
using Aaks.PowerPointParser.Dto;
using FluentAssertions;
using PowerPointParserTests.Html;

// ReSharper disable once CheckNamespace - Test namespaces should match production
namespace Aaks.PowerPointParser.Html.Tests
{
    [TestClass]
    public class HtmlBuilderTests : BaseHtmlTests
    {
        private static IHtmlBuilder _htmlConverter;

        [ClassInitialize]
        public static void ClassSetup(TestContext context)
        {
            var innerHtmlBuilder = new InnerHtmlBuilder();
            _htmlConverter = new HtmlBuilder(new HtmlListBuilder(innerHtmlBuilder), innerHtmlBuilder);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_WrapperListNull_ReturnsNull()
        {
            Queue<OpenXmlParagraphWrapper?>? wrapperList = null;

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(wrapperList);

            actual.Should().BeEmpty();
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_WrapperNull_ReturnsNull()
        {
            OpenXmlParagraphWrapper? wrapper = null;
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            actual.Should().BeEmpty();
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_RNull_ReturnsNull()
        {
            OpenXmlParagraphWrapper wrapper = new();
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            actual.Should().BeEmpty();
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_REmptyNull_ReturnsNull()
        {
            OpenXmlParagraphWrapper wrapper = new()
            {
                R = new List<R>()
            };

            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            actual.Should().BeEmpty();
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ParagraphTag_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p>hello world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_BoldTag_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", new RPr {B = 1}));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong>hello world</strong></p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_BoldAndNonBuildMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {B = 1}),
                BuildR(" world")
            };

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr {BuNone = new object()},
                R = rs
            };

            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong>hello:</strong> world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ConsecutiveParagraph_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello"));

            Queue<OpenXmlParagraphWrapper?> queue2 = new();
            queue2.Enqueue(BuildParagraphLine("world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);
            actual += _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue2);

            Assert.AreEqual("<p>hello</p><p>world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OneLineInTwoRRecords_ReturnsString()
        {
            var rs = new List<R>();
            var rOne = new R {T = "hello "};
            var rTwo = new R {T = "world"};
            rs.Add(rOne);
            rs.Add(rTwo);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr {BuNone = new object()},
                R = rs
            };

            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p>hello world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_UnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", new RPr {B = 1}));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li><strong>hello world</strong></li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_UnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><li>goodbye world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedListItem_ReturnsString()
        {
            var wrapper = BuildOrderListItem("hello world");

            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_EmbeddedOrderedListItem_ReturnsString()
        {
            var rs = new List<R>();
            var rOne = BuildR("hello");
            var rTwo = BuildR(" world");
            var rThree = BuildR(" ");
            var rFour = BuildR("test");

            rs.Add(rOne);
            rs.Add(rTwo);
            rs.Add(rThree);
            rs.Add(rFour);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr {BuAutoNum = new BuAutoNum {Type = "arabicPeriod"}},
                R = rs
            };

            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world test</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world"));
            queue.Enqueue(BuildOrderListItem("goodbye world"));
            queue.Enqueue(BuildOrderListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li><li>goodbye world</li><li>test world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><ul><li>nested item</li></ul><li>test world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TwiceNestedUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual(
                "<ul><li>hello world</li><ul><li>nested item</li><ul><li>nested double</li></ul></ul><li>test world</li></ul>",
                actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TwiceTwoNestedUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("nested double two", level: 2));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual(
                "<ul><li>hello world</li><ul><li>nested item</li><ul><li>nested double</li><li>nested double two</li></ul></ul><li>test world</li></ul>",
                actual);
        }


        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TwiceNestedFollowedBySingleNested_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("nested double two", level: 2));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual(
                "<ul><li>hello world</li><ul><li>nested item</li><ul><li>nested double</li><li>nested double two</li></ul><li>test world</li></ul></ul>",
                actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlternateEveryItem_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li></ol><ul><li>one one</li></ul><ol><li>one one one</li></ol><ul><li>one one one one</li></ul>", actual);

        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlternateEveryNestItem_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li></ol><ul><li>one one</li></ul><ol><li>one one one</li></ol><ul><li>one one one one</li><ol><li>two</li></ol><ul><li>two two</li></ul><ol><li>two two two</li></ol><ul><li>two two two two</li></ul></ul>", actual);

        }



        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TripleNested_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildOrderListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("nested double two", level: 2));
            queue.Enqueue(BuildOrderListItem("three nested one", level: 3));
            queue.Enqueue(BuildOrderListItem("three nested two", level: 3));
            queue.Enqueue(BuildOrderListItem("test world", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><ol><li>nested item</li><ul><li>nested double</li><li>nested double two</li><ol><li>three nested one</li><li>three nested two</li></ol></ul><li>test world</li></ol></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedNestUnorderedNestOrdered_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("nested double two", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three nested one", level: 3));
            queue.Enqueue(BuildUnorderedListItem("three nested two", level: 3));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><ul><li>nested item</li><ul><li>nested double</li><li>nested double two</li><ul><li>three nested one</li><li>three nested two</li></ul></ul><li>test world</li></ul></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedLastItemUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item"));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><li>nested item</li><ul><li>test world</li></ul></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedUnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("two"));
            queue.Enqueue(BuildUnorderedListItem("hello world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("goodbye world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three"));


            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>one</li><li>two</li><ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul><li>three</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_UnorderedFollowedByOrderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));
            queue.Enqueue(BuildUnorderedListItem("test world"));
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildOrderListItem("three"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul><ol><li>one</li><li>two</li><li>three</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedFollowedByUnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildOrderListItem("three"));
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));
            queue.Enqueue(BuildUnorderedListItem("test world"));
            
            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><li>three</li></ol><ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedOrderedFollowedByUnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildOrderListItem("three", level: 1));
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><ol><li>three</li></ol></ol><ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedFollowedByNestedUnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildOrderListItem("three"));
            queue.Enqueue(BuildUnorderedListItem("hello world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><li>three</li></ol><ul><ul><li>hello world</li></ul><li>goodbye world</li><li>test world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedTwoUnorderedInsideOrderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));
            queue.Enqueue(BuildOrderListItem("three"));


            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><li>hello world</li><ul><li>goodbye world</li><li>test world</li></ul><li>three</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AllNestedTwoUnorderedInsideOrderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one", level: 1));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("hello world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("goodbye world", level: 2));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 2));
            queue.Enqueue(BuildOrderListItem("three", level: 1));


            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><ol><li>one</li><li>two</li><li>hello world</li><ul><li>goodbye world</li><li>test world</li></ul><li>three</li></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedUnorderedInsideOrderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildUnorderedListItem("hello world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("goodbye world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));
            queue.Enqueue(BuildOrderListItem("three"));
            

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul><li>three</li></ol>", actual);
        }


        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingAlternatingOrderEveryTwo_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><ol><li>two</li><ul><li>three</li><ul><li>four</li><ol><li>five</li><ol><li>six</li><ul><li>seven</li><ul><li>eight</li><ol><li>nine</li><ol><li>ten</li></ol></ol></ul></ul></ol></ol></ul></ul></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingAlternatingTwoOrderEveryTwo_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>one one</li><ol><li>two</li><li>two two</li><ul><li>three</li><li>three three</li>" +
                            "<ul><li>four</li><li>four four</li><ol><li>five</li><li>five five</li>" +
                            "<ol><li>six</li><li>six six</li><ul><li>seven</li><li>seven seven</li>" +
                            "<ul><li>eight</li><li>eight eight</li><ol><li>nine</li><li>nine nine</li>" +
                            "<ol><li>ten</li><li>ten ten</li></ol></ol></ul></ul></ol></ol></ul></ul></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingAlternatingThreeOrderEveryTwo_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>one one</li><li>one one one</li><ol><li>two</li><li>two two</li><li>two two two</li><ul><li>three</li><li>three three</li><li>three three three</li><ul><li>four</li><li>four four</li><li>four four four</li><ol><li>five</li><li>five five</li><li>five five five</li><ol><li>six</li><li>six six</li><li>six six six six</li><ul><li>seven</li><li>seven seven</li><li>seven seven seven</li><ul><li>eight</li><li>eight eight</li><li>eight eight eight</li><ol><li>nine</li><li>nine nine</li><li>nine nine nine</li><ol><li>ten</li><li>ten ten</li><li>ten ten ten</li></ol></ol></ul></ul></ol></ol></ul></ul></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingAlternatingFourOrderEveryTwo_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildOrderListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li><ol><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li><ul><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li><ul><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li><ol><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li><ol><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li><ul><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li><ul><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li><ol><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li><ol><li>ten</li><li>ten ten</li><li>ten ten ten</li><li>ten ten ten ten</li></ol></ol></ul></ul></ol></ol></ul></ul></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingFourUnordered_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));
            queue.Enqueue(BuildUnorderedListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li><ul><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li><ul><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li><ul><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li><ul><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li><ul><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li><ul><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li><ul><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li><ul><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li><ul><li>ten</li><li>ten ten</li><li>ten ten ten</li><li>ten ten ten ten</li></ul></ul></ul></ul></ul></ul></ul></ul></ul></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingFourOrdered_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildOrderListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li><ol><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li><ol><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li><ol><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li><ol><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li><ol><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li><ol><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li><ol><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li><ol><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li><ol><li>ten</li><li>ten ten</li><li>ten ten ten</li><li>ten ten ten ten</li></ol></ol></ol></ol></ol></ol></ol></ol></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingFourAlternateEachOrdered_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li></ol><ul><li>one one</li></ul><ol><li>one one one</li></ol><ul><li>one one one one</li><ol><li>two</li></ol><ul><li>two two</li></ul><ol><li>two two two</li></ol><ul><li>two two two two</li><ol><li>three</li></ol><ul><li>three three</li></ul><ol><li>three three three</li></ol><ul><li>three three three three</li><ol><li>four</li></ol><ul><li>four four</li></ul><ol><li>four four four</li></ol><ul><li>four four four four</li><ol><li>five</li></ol><ul><li>five five</li></ul><ol><li>five five five</li></ol><ul><li>five five five five</li><ol><li>six</li></ol><ul><li>six six</li></ul><ol><li>six six six six</li></ol><ul><li>six six six six six</li><ol><li>seven</li></ol><ul><li>seven seven</li></ul><ol><li>seven seven seven</li></ol><ul><li>seven seven seven seven</li><ol><li>eight</li></ol><ul><li>eight eight</li></ul><ol><li>eight eight eight</li></ol><ul><li>eight eight eight eight</li><ol><li>nine</li></ol><ul><li>nine nine</li></ul><ol><li>nine nine nine</li></ol><ul><li>nine nine nine nine</li><ol><li>ten</li></ol><ul><li>ten ten</li></ul><ol><li>ten ten ten</li></ol><ul><li>ten ten ten ten</li></ul></ul></ul></ul></ul></ul></ul></ul></ul></ul>", actual);
        }
    }
}