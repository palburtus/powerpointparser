using System.Linq;
using System.Text;
using System.Web;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public class InnerHtmlBuilder : IInnerHtmlBuilder
    {
        public string BuildInnerHtmlParagraph(OpenXmlTextWrapper textWrapper)
        {
            StringBuilder sb = new();

            sb.Append("<p>");
            sb.Append(BuildInnerHtml(textWrapper));
            sb.Append("</p>");

            return sb.ToString();
        }

        public string BuildInnerHtmlListItem(OpenXmlTextWrapper textWrapper)
        {
            StringBuilder sb = new();

            sb.Append("<li>");
            sb.Append(BuildInnerHtml(textWrapper));
            sb.Append("</li>");

            return sb.ToString();
        }

        private static string BuildInnerHtml(OpenXmlTextWrapper textWrapper)
        {
            StringBuilder sb = new();
            foreach (var r in textWrapper.R!.Where(r => r.T != null))
            {
                if (IsBold(r)) sb.Append(Tags.Open(Tags.Bold));
                if (IsUnderlined(r)) sb.Append(Tags.Open(Tags.Underlined));
                if (IsItalic(r)) sb.Append(Tags.Open(Tags.Italic));
                if (IsStrikeThrough(r)) sb.Append(Tags.Open(Tags.StrikeThrough));

                sb.Append(HttpUtility.HtmlEncode(r.T));

                if (IsStrikeThrough(r)) sb.Append(Tags.Close(Tags.StrikeThrough));
                if (IsItalic(r)) sb.Append(Tags.Close(Tags.Italic));
                if (IsUnderlined(r)) sb.Append(Tags.Close(Tags.Underlined));
                if (IsBold(r)) sb.Append(Tags.Close(Tags.Bold));
            }

            return sb.ToString();
        }

        private static bool IsStrikeThrough(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.Strike == "sngStrike";
        }

        private static bool IsUnderlined(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.U == "sng";
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
