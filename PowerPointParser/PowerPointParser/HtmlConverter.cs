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
        public string? ConvertOpenXmlParagraphWrapperToHtml(Queue<OpenXmlParagraphWrapper?>? paragraphWrappers)
        {
            return ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, null);
        }

        public string? ConvertHtmlParagraphWrapperToHtml(Queue<OpenXmlParagraphWrapper?>? paragraphWrappers, OpenXmlParagraphWrapper? previous)
        {
            if (paragraphWrappers == null) { return null; }
            
            StringBuilder sb = new StringBuilder();
            while (paragraphWrappers.Count > 0)
            {
                var paragraphWrapper = paragraphWrappers.Dequeue();

                if (paragraphWrapper?.R == null) return null;
                if (paragraphWrapper.R.Count == 0) return null;
                
                bool isListItem = IsListItem(paragraphWrapper);

                if (!isListItem)
                {
                    sb.Append(BuildInnerHtml(paragraphWrapper, isListItem));
                    sb.Append(ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, paragraphWrapper));
                }
                else
                {
                    bool isOrderListItem = IsOrderedListItem(paragraphWrapper);

                    if (IsFirstListItem(previous))
                    {
                        sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                    }

                    sb.Append(BuildInnerHtml(paragraphWrapper, isListItem));
                    sb.Append(ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, paragraphWrapper));

                    
                    if (IsLastListItem(previous))
                    {
                        sb.Append(isOrderListItem ? "</ol>" : "</ul>");
                    }
                }
            }

            
            
                
                
                /* else
                 {
                    int indent = GetParagraphIndentLevel(paragraphWrapper);

                     if (indent > currentIndentLevel)
                     {
                         sb.Append("<ul>");
                         sb.Append(ConvertHtmlParagraphWrapperToHtml(paragraphWrapper, indent));
                     }
                     else if (indent < currentIndentLevel)
                     {
                         sb.Append("</ul>");
                         sb.Append(ConvertHtmlParagraphWrapperToHtml(paragraphWrapper, indent));
                     }
                     else
                     {
                         sb.Append(BuildInnerHtml(paragraphWrapper, isListItem));
                     }
                 }*/
            

            return sb.ToString();
        }

        private bool IsLastListItem(OpenXmlParagraphWrapper? previous)
        {
            return previous == null;
        }

        private bool IsFirstListItem(OpenXmlParagraphWrapper? previous)
        {
            return previous == null || !IsListItem(previous);
        }

        private bool IsListItem(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper == null) return false;
            if (paragraphWrapper.PPr == null) return false;

            if (paragraphWrapper.PPr.BuAutoNum != null)
            {
                return paragraphWrapper.PPr.BuAutoNum.Type == "arabicPeriod";
            }

            if (paragraphWrapper.PPr.BuChar != null)
            {
                return paragraphWrapper.PPr.BuChar.Char == "•";
            }

            return false;
        }

        private string BuildInnerHtml(OpenXmlParagraphWrapper paragraphWrapper, bool isListItem)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(!isListItem ? "<p>" : "<li>");

            foreach (var r in paragraphWrapper.R!)
            {
                if (r.T != null)
                {
                    if (IsBold(r)) sb.Append("<strong>");

                    sb.Append(r.T);

                    if (IsBold(r)) sb.Append("</strong>");
                }
            }

            sb.Append(!isListItem ? "</p>" : "</li>");

            return sb.ToString();
        }

        private static int GetParagraphIndentLevel(OpenXmlParagraphWrapper paragraphWrapper)
        {
            return (paragraphWrapper.PPr?.Lvl ?? 0) + 1;
        }

        private bool IsOrderedListItem(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper == null) return false;
            if (paragraphWrapper.PPr == null) return false;
            if (paragraphWrapper.PPr.BuAutoNum == null) return false;
            if (paragraphWrapper.PPr.BuAutoNum.Type == null) return false;
            return IsListItem(paragraphWrapper) && paragraphWrapper.PPr.BuAutoNum.Type == "arabicPeriod";
        }

        private bool IsBold(R? r)
        {
            if(r == null) return false;
            if (r.RPr == null) return false;
            return r.RPr.B == 1;
        }
    }
}
