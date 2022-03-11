using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointParser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser.Dto;

namespace PowerPointParser.Tests
{
    [TestClass()]
    public class HtmlConverterTests
    {
        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_WrapperNull_ReturnsNull()
        {
            IHtmlConverter converter = new HtmlConverter();

            OpenXmlParagraphWrapper? wrapper = null;

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

            Assert.IsNull(actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_RNull_ReturnsNull()
        {
            IHtmlConverter converter = new HtmlConverter();

            OpenXmlParagraphWrapper? wrapper = new ();

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

            Assert.IsNull(actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_REmptyNull_ReturnsNull()
        {
            IHtmlConverter converter = new HtmlConverter();

            OpenXmlParagraphWrapper? wrapper = new()
            {
                R = new List<R>()
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

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

            OpenXmlParagraphWrapper? wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

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

            OpenXmlParagraphWrapper? wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

            Assert.AreEqual("<p><strong>hello world</strong></p>", actual);
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

            OpenXmlParagraphWrapper? wrapper = new()
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

            OpenXmlParagraphWrapper? wrapper2 = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs2
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);
            actual += converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper2);

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

            OpenXmlParagraphWrapper? wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

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

            OpenXmlParagraphWrapper? wrapper = new()
            {
                PPr = new PPr {BuNone = null, BuChar = new BuChar{ Char =  "•" }},
                R = rs
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

            Assert.AreEqual("<li><strong>hello world</strong></li>", actual);
        }

        [TestMethod()]
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_OrderedListItem_ReturnsString()
        {
            IHtmlConverter converter = new HtmlConverter();

            var rs = new List<R>();
            var r = new R { T = "hello world" };
            rs.Add(r);

            OpenXmlParagraphWrapper? wrapper = new()
            {
                PPr = new PPr { BuAutoNum = new BuAutoNum{ Type = "arabicPeriod"}},
                R = rs
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

            Assert.AreEqual("<li>hello world</li>", actual);
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

            OpenXmlParagraphWrapper? wrapper = new()
            {
                PPr = new PPr { BuAutoNum = new BuAutoNum { Type = "arabicPeriod" } },
                R = rs
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

            Assert.AreEqual("<li>hello world test</li>", actual);
        }
    }
}