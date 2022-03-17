using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser.Dto;

namespace PowerPointParser.Html
{
    public class InnerHtmlBuilder : IInnerHtmlBuilder
    {
        public string BuildInnerHtmlParagraph(OpenXmlParagraphWrapper paragraphWrapper)
        {
            StringBuilder sb = new();
            sb.Append("<p>");

            foreach (var r in paragraphWrapper.R!.Where(r => r.T != null))
            {
                if (IsBold(r)) sb.Append("<strong>");

                sb.Append(r.T);

                if (IsBold(r)) sb.Append("</strong>");
            }

            sb.Append("</p>");

            return sb.ToString();
        }

        public string BuildInnerHtmlListItem(OpenXmlParagraphWrapper paragraphWrapper)
        {
            StringBuilder sb = new();
            sb.Append("<li>");

            foreach (var r in paragraphWrapper.R!.Where(r => r.T != null))
            {
                if (IsBold(r)) sb.Append("<strong>");

                sb.Append(r.T);

                if (IsBold(r)) sb.Append("</strong>");
            }

            sb.Append("</li>");

            return sb.ToString();
        }

        private static bool IsBold(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.B == 1;
        }
    }
}
