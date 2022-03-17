using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointParser.Dto;
using PowerPointParserTests.Html;

// ReSharper disable once CheckNamespace - Test namespaces should match production
namespace PowerPointParser.Html.Tests
{
    [TestClass]
    public class HtmlBuilderTests : BaseHtmlTests
    {
        private static IHtmlBuilder _htmlConverter;

        [ClassInitialize]
        public static void ClassSetup(TestContext context)
        {
            _htmlConverter = new HtmlBuilder(new InnerHtmlBuilder());
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_WrapperListNull_ReturnsNull()
        {
            Queue<OpenXmlParagraphWrapper?>? wrapperList = null;

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(wrapperList);

            Assert.IsNull(actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_WrapperNull_ReturnsNull()
        {
            OpenXmlParagraphWrapper? wrapper = null;
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.IsNull(actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_RNull_ReturnsNull()
        {
            OpenXmlParagraphWrapper wrapper = new();
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.IsNull(actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_REmptyNull_ReturnsNull()
        {
            OpenXmlParagraphWrapper wrapper = new()
            {
                R = new List<R>()
            };

            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.IsNull(actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_ParagraphTag_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p>hello world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_BoldTag_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", new RPr { B = 1 }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong>hello world</strong></p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_BoldAndNonBuildMix_ReturnsString()
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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_ConsecutiveParagraph_ReturnsString()
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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_OneLineInTwoRRecords_ReturnsString()
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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_UnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", new RPr { B = 1 }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li><strong>hello world</strong></li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_UnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><li>goodbye world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_OrderedListItem_ReturnsString()
        {
            var wrapper = BuildOrderListItem("hello world");

            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_EmbeddedOrderedListItem_ReturnsString()
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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_OrderedListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world"));
            queue.Enqueue(BuildOrderListItem("goodbye world"));
            queue.Enqueue(BuildOrderListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li><li>goodbye world</li><li>test world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_NestedUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level:1));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><ul><li>nested item</li></ul><li>test world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_NestedLastItemUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlParagraphWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item"));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><li>nested item</li><ul><li>test world</li></ul></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_UnorderedFollowedByOrderedListItems_ReturnsString()
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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_OrderedFollowedByUnorderedListItems_ReturnsString()
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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_NestedOrderedFollowedByUnorderedListItems_ReturnsString()
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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_OrderedFollowedByNestedUnorderedListItems_ReturnsString()
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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_NestedUnorderedInsideOrderedListItems_ReturnsString()
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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_NestedUnorderedListItems_ReturnsString()
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

        
    }
}