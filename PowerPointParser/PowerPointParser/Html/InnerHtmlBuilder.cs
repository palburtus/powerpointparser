using System.Linq;
using System.Text;
using System.Web;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public class InnerHtmlBuilder : IInnerHtmlBuilder
    {
        public string BuildInnerHtmlParagraph(OpenXmlParagraphWrapper paragraphWrapper)
        {
            StringBuilder sb = new();

            sb.Append("<p>");
            sb.Append(BuildInnerHtml(paragraphWrapper));
            sb.Append("</p>");

            return sb.ToString();
        }

        public string BuildInnerHtmlListItem(OpenXmlParagraphWrapper paragraphWrapper)
        {
            StringBuilder sb = new();

            sb.Append("<li>");
            sb.Append(BuildInnerHtml(paragraphWrapper));
            sb.Append("</li>");

            return sb.ToString();
        }

        private static string BuildInnerHtml(OpenXmlParagraphWrapper paragraphWrapper)
        {
            StringBuilder sb = new();
            foreach (var r in paragraphWrapper.R!.Where(r => r.T != null))
            {

                if (IsBold(r)) sb.Append("<strong>");
                if (IsItalic(r)) sb.Append("<i>");

                sb.Append(HttpUtility.HtmlEncode(r.T));

                if (IsItalic(r)) sb.Append("</i>");
                if (IsBold(r)) sb.Append("</strong>");
            }

            return sb.ToString();
        }

        private static bool IsBold(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.B == 1;
        }

        private static bool IsItalic(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.I == 1;
        }
    }
}
