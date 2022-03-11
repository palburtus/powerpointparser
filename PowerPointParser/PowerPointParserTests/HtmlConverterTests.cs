using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointParser.Dto;

// ReSharper disable once CheckNamespace - Test namespaces should match production
namespace PowerPointParser.Tests
{
    [TestClass()]
    public class HtmlConverterTests
    {

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_WrapperListNull_ReturnsNull()
        {
            IHtmlConverter converter = new HtmlConverter();

            Stack<OpenXmlParagraphWrapper?>? wrapperList = null;

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapperList);

            Assert.IsNull(actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_WrapperNull_ReturnsNull()
        {
            IHtmlConverter converter = new HtmlConverter();

            OpenXmlParagraphWrapper? wrapper = null;
            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.IsNull(actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_RNull_ReturnsNull()
        {
            IHtmlConverter converter = new HtmlConverter();

            OpenXmlParagraphWrapper wrapper = new ();
            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.IsNull(actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_REmptyNull_ReturnsNull()
        {
            IHtmlConverter converter = new HtmlConverter();

            OpenXmlParagraphWrapper wrapper = new()
            {
                R = new List<R>()
            };

            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.IsNull(actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_ParagraphTag_ReturnsString()
        {
            IHtmlConverter converter = new HtmlConverter();

            var rs = new List<R>();
            var r = new R
            {
                T = "hello world"
            };
            rs.Add(r);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.AreEqual("<p>hello world</p>", actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_BoldTag_ReturnsString()
        {
            IHtmlConverter converter = new HtmlConverter();

            var rs = new List<R>();
            var r = new R
            {
                RPr = new RPr { B = 1 },
                T = "hello world"
            };
            rs.Add(r);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.AreEqual("<p><strong>hello world</strong></p>", actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_BoldAndNonBuildMix_ReturnsString()
        {
            IHtmlConverter converter = new HtmlConverter();

            var rs = new List<R>();
            var r1 = new R
            {
                RPr = new RPr { B = 1 },
                T = "hello:"
            };
            var r2 = new R {T = " world"};
            rs.Add(r1);
            rs.Add(r2);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.AreEqual("<p><strong>hello:</strong> world</p>", actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_ConsecutiveParagraph_ReturnsString()
        {
            IHtmlConverter converter = new HtmlConverter();

            var rs = new List<R>();
            var r = new R
            {
                T = "hello"
            };
            rs.Add(r);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            var rs2 = new List<R>();
            var r2 = new R
            {
                T = "world"
            };
            rs2.Add(r2);

            OpenXmlParagraphWrapper wrapper2 = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs2
            };

            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            Stack<OpenXmlParagraphWrapper?> stack2 = new();
            stack2.Push(wrapper2);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);
            actual += converter.ConvertOpenXmlParagraphWrapperToHtml(stack2);

            Assert.AreEqual("<p>hello</p><p>world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_OneLineInTwoRRecords_ReturnsString()
        {
            IHtmlConverter converter = new HtmlConverter();

            var rs = new List<R>();
            var rOne = new R { T = "hello " };
            var rTwo = new R { T = "world" };
            rs.Add(rOne);
            rs.Add(rTwo);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.AreEqual("<p>hello world</p>", actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_UnorderedListItem_ReturnsString()
        {
            IHtmlConverter converter = new HtmlConverter();

            var rs = new List<R>();
            var r = new R
            {
                RPr = new RPr { B = 1 },
                T = "hello world"
            };
            rs.Add(r);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr {BuNone = null, BuChar = new BuChar{ Char =  "•" }},
                R = rs
            };

            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.AreEqual("<ul><li><strong>hello world</strong></li></ul>", actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_OrderedListItem_ReturnsString()
        {
            IHtmlConverter converter = new HtmlConverter();

            var rs = new List<R>();
            var r = new R { T = "hello world" };
            rs.Add(r);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr { BuAutoNum = new BuAutoNum{ Type = "arabicPeriod"}},
                R = rs
            };

            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_EmbeddedOrderedListItem_ReturnsString()
        {
            IHtmlConverter converter = new HtmlConverter();

            var rs = new List<R>();
            var rOne = new R { T = "hello " };
            var rTwo = new R { T = "world" };
            var rThree = new R { T = " " };
            var rFour = new R { T = "test" };

            rs.Add(rOne);
            rs.Add(rTwo);
            rs.Add(rThree);
            rs.Add(rFour);

            OpenXmlParagraphWrapper wrapper = new()
            {
                PPr = new PPr { BuAutoNum = new BuAutoNum { Type = "arabicPeriod" } },
                R = rs
            };

            Stack<OpenXmlParagraphWrapper?> stack = new();
            stack.Push(wrapper);

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(stack);

            Assert.AreEqual("<ol><li>hello world test</li></ol>", actual);
        }
    }
}