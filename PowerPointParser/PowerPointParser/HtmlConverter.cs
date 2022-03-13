using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            
            StringBuilder sb = new();
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

                    if (IsFirstListItem(previous, paragraphWrapper))
                    {
                        sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                    }

                    sb.Append(BuildInnerHtml(paragraphWrapper, isListItem));
                    
                    if (IsEndOfNestedList(previous, paragraphWrapper))
                    {
                        sb.Append(isOrderListItem ? "</ol>" : "</ul>");
                    }
                    
                    sb.Append(ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, paragraphWrapper));

                    if (IsLastListItem(previous, paragraphWrapper))
                    {
                        sb.Append(isOrderListItem ? "</ol>" : "</ul>");
                    }
                }
            }
            
            return sb.ToString();
        }

        private static bool IsEndOfNestedList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current)
        {
            return previous?.PPr?.Lvl < current?.PPr?.Lvl;
        }

        private static bool IsLastListItem(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current)
        {
            return IsListItem(current) && previous == null;
        }

        private static bool IsFirstListItem(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current)
        {
            if (current?.PPr?.Lvl > previous?.PPr?.Lvl)
            {
                return true;
            }

            return previous == null || !IsListItem(previous);
        }

        private static bool IsListItem(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper?.PPr == null) return false;

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

        private static string BuildInnerHtml(OpenXmlParagraphWrapper paragraphWrapper, bool isListItem)
        {
            StringBuilder sb = new();
            sb.Append(!isListItem ? "<p>" : "<li>");

            foreach (var r in paragraphWrapper.R!.Where(r => r.T != null))
            {
                if (IsBold(r)) sb.Append("<strong>");

                sb.Append(r.T);

                if (IsBold(r)) sb.Append("</strong>");
            }

            sb.Append(!isListItem ? "</p>" : "</li>");

            return sb.ToString();
        }

        private static bool IsOrderedListItem(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper?.PPr?.BuAutoNum?.Type == null) return false;
            return IsListItem(paragraphWrapper) && paragraphWrapper.PPr.BuAutoNum.Type == "arabicPeriod";
        }

        private static bool IsBold(R? r)
        {
            if (r?.RPr == null) return false;
            return r.RPr.B == 1;
        }
    }
}
