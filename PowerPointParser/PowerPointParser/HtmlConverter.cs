using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser.Dto;

namespace PowerPointParser
{
    public class HtmlConverter : IHtmlConverter
    {
        public string? ConvertOpenXmlParagraphWrapperToHtml(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper == null) return null;
            if (paragraphWrapper.R == null) return null;
            if (paragraphWrapper.R.Count == 0) return null;

            StringBuilder sb = new StringBuilder();


            bool isListItem = IsListItem(paragraphWrapper);
            
            if (!isListItem) sb.Append("<p>");

            foreach (var r in paragraphWrapper.R)
            {
                if (r.T != null)
                {
                    if (isListItem) sb.Append("<li>");
                    if (IsBold(r)) sb.Append("<strong>");

                    sb.Append(r.T);

                    if (IsBold(r)) sb.Append("</strong>");
                    if (isListItem) sb.Append("</li>");
                }
            }

            if (!isListItem) sb.Append("</p>");

            return sb.ToString();
        }

        private bool IsOrderedListItem(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper == null) return false;
            if (paragraphWrapper.PPr == null) return false;
            if (paragraphWrapper.PPr.BuAutoNum == null) return false;
            if (paragraphWrapper.PPr.BuAutoNum.Type == null) return false;
            return IsListItem(paragraphWrapper) && paragraphWrapper.PPr.BuAutoNum.Type == "arabicPeriod";
        }

        private bool IsListItem(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if(paragraphWrapper == null) return false;
            if(paragraphWrapper.PPr == null) return false;
            return paragraphWrapper.PPr.BuNone == null;
        }

        private bool IsBold(R? r)
        {
            if(r == null) return false;
            if (r.RPr == null) return false;
            return r.RPr.B == 1;
        }
    }
}
