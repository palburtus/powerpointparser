using System.Collections.Generic;
using Aaks.PowerPointParser.Dto;

namespace PowerPointParserTests.Html
{
    public abstract class BaseHtmlTests
    {
        protected static OpenXmlTextWrapper BuildUnorderedListItem(string text, RPr? rPr = null, int level = 0, string? align = null, string? specialChar = null)
        {
            var buChar = new BuChar {Character = specialChar ?? "•"};

            var wrapper = BuildLine(text, rPr);
            wrapper.PPr = new PPr { BuNone = null, BuChar = buChar, Lvl = level, Algn = align};
            return wrapper;
        }

        protected static OpenXmlTextWrapper BuildOrderListItem(string text, RPr? rPr = null, int level = 0, string? align = null, string? specialFont = null)
        {
            var buAutoNum = new BuAutoNum {Type = specialFont?? "arabicPeriod"};

            var wrapper = BuildLine(text, rPr);
            wrapper.PPr = new PPr { BuAutoNum = buAutoNum, Lvl = level, Algn = align};
            return wrapper;
        }

        protected static OpenXmlTextWrapper BuildParagraphLine(string text, RPr? rPr = null, PPr? ppr = null)
        {
            var wrapper = BuildLine(text, rPr);
            wrapper.PPr = ppr;
            return wrapper;
        }

        protected static OpenXmlTextWrapper BuildLine(string text, RPr? rPr = null, PPr? ppr = null)
        {
            var rs = new List<R>();
            var r = BuildR(text, rPr);
            rs.Add(r);

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = ppr,
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
