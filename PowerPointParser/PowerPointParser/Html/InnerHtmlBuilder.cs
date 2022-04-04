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
            return textWrapper.R?.Count == 0 ? string.Empty : $"{HtmlTags.Open(HtmlTags.Paragraph, GetTextAlignment(textWrapper.PPr))}{BuildInnerHtml(textWrapper)}{HtmlTags.Close(HtmlTags.Paragraph)}";
        }

        public string BuildInnerHtmlListItem(OpenXmlTextWrapper textWrapper)
        {
            return $"{HtmlTags.Open(HtmlTags.ListItem, GetTextAlignment(textWrapper.PPr))}{BuildInnerHtml(textWrapper)}{HtmlTags.Close(HtmlTags.ListItem)}";
        }

        public string BuildInnerHtmlListItemBeforeNesting(OpenXmlTextWrapper textWrapper)
        {
            return $"{HtmlTags.Open(HtmlTags.ListItem, GetTextAlignment(textWrapper.PPr))}{BuildInnerHtml(textWrapper)}";
        }

        private static string BuildInnerHtml(OpenXmlTextWrapper textWrapper)
        {
            StringBuilder sb = new();
            foreach (var r in textWrapper.R!.Where(r => r.T != null))
            {
                if (IsBold(r)) sb.Append(HtmlTags.Open(HtmlTags.Bold));
                if (IsUnderlined(r)) sb.Append(HtmlTags.Open(HtmlTags.Underlined));
                if (IsItalic(r)) sb.Append(HtmlTags.Open(HtmlTags.Italic));
                if (IsStrikeThrough(r)) sb.Append(HtmlTags.Open(HtmlTags.StrikeThrough));

                sb.Append(HttpUtility.HtmlEncode(r.T));

                if (IsStrikeThrough(r)) sb.Append(HtmlTags.Close(HtmlTags.StrikeThrough));
                if (IsItalic(r)) sb.Append(HtmlTags.Close(HtmlTags.Italic));
                if (IsUnderlined(r)) sb.Append(HtmlTags.Close(HtmlTags.Underlined));
                if (IsBold(r)) sb.Append(HtmlTags.Close(HtmlTags.Bold));
            }

            return sb.ToString();
        }

        private static string GetTextAlignment(PPr? pPr)
        {
            return pPr?.Algn switch
            {
                OpenXmlTextModifiers.AlignTextCenter => TextAlignment.Center,
                OpenXmlTextModifiers.AlignTextRight => TextAlignment.Right,
                OpenXmlTextModifiers.AlignTextJustify => TextAlignment.Justify,
                _ => TextAlignment.Left
            };
        }

        private static bool IsStrikeThrough(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.Strike == OpenXmlTextModifiers.StrikeThrough;
        }

        private static bool IsUnderlined(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.U == OpenXmlTextModifiers.Underlined;
        }

        private static bool IsBold(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.B == OpenXmlTextModifiers.Bold;
        }

        private static bool IsItalic(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.I == OpenXmlTextModifiers.Italic;
        }
    }
}
