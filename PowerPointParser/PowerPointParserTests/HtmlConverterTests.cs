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
        public void ConvertOpenXmlParagraphWrapperToHtmlTest_ConsectiveParagraph_ReturnsString()
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
        /*
         * Assert.AreEqual("And this is a ", actual.R![0].T);
            Assert.AreEqual(0, actual.R![0].RPr!.B);
            Assert.AreEqual(0, actual.R![0].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![0].RPr!.Lang);

            Assert.AreEqual("second paragraph", actual.R![1].T);
            Assert.AreEqual(0, actual.R![1].RPr!.B);
            Assert.AreEqual(0, actual.R![1].RPr!.Dirty);
            Assert.AreEqual("en-US", actual.R![1].RPr!.Lang);
         */

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
                PPr = new PPr {BuNone = null},
                R = rs
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

            Assert.AreEqual("<li><strong>hello world</strong></li>", actual);
        }
    }
}