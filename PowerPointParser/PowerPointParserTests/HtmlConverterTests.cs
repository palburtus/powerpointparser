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
    }
}