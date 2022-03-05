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

            Assert.AreEqual("<p><b>hello world</b></p>", actual);
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
                PPr = new PPr {BuNone = null},
                R = rs
            };

            var actual = converter.ConvertOpenXmlParagraphWrapperToHtml(wrapper);

            Assert.AreEqual("<li><b>hello world</b></li>", actual);
        }
    }
}