using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser.Dto;

namespace PowerPointParserTests.Html
{
    public abstract class BaseHtmlTests
    {
        protected static OpenXmlParagraphWrapper BuildUnorderedListItem(string text, RPr? rPr = null, int level = 0)
        {
            var wrapper = BuildLine(text, rPr);
            wrapper.PPr = new PPr { BuNone = null, BuChar = new BuChar { Char = "•" }, Lvl = level };
            return wrapper;
        }

        protected static OpenXmlParagraphWrapper BuildOrderListItem(string text, RPr? rPr = null, int level = 0)
        {
            var wrapper = BuildLine(text, rPr);
            wrapper.PPr = new PPr { BuAutoNum = new BuAutoNum { Type = "arabicPeriod" }, Lvl = level };
            return wrapper;
        }

        protected static OpenXmlParagraphWrapper BuildParagraphLine(string text, RPr? rPr = null)
        {
            var wrapper = BuildLine(text, rPr);
            wrapper.PPr = new PPr { BuNone = new object() };
            return wrapper;
        }

        protected static OpenXmlParagraphWrapper BuildLine(string text, RPr? rPr = null)
        {
            var rs = new List<R>();
            var r = BuildR(text, rPr);
            rs.Add(r);

            OpenXmlParagraphWrapper wrapper = new()
            {
                R = rs
            };
            return wrapper;
        }

        protected static R BuildR(string text, RPr? rPr = null)
        {
            return new R { RPr = rPr, T = text };
        }
    }
}
