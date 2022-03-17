using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
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

        private string? ConvertHtmlParagraphWrapperToHtml(Queue<OpenXmlParagraphWrapper?>? paragraphWrappers, OpenXmlParagraphWrapper? previous)
        {
            if (paragraphWrappers == null) { return null; }
            
            StringBuilder sb = new();
            while (paragraphWrappers.Count > 0)
            {
                var current = paragraphWrappers.Dequeue();
                paragraphWrappers.TryPeek(out var next);

                if (current?.R == null) return null;
                if (current.R.Count == 0) return null;
                
                bool isListItem = IsListItem(current);

                if (!isListItem)
                {
                    sb.Append(BuildInnerHtml(current, isListItem));
                    sb.Append(ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, current));
                }
                else
                {
                    bool isOrderListItem = IsOrderedListItem(current);

                    if (IsListOrderTypeChanged(previous, current))
                    {
                        sb.Append(isOrderListItem ? "</ul><ol>" : "</ol><ul>");
                    }

                    
                    if (IsFirstListItem(previous, current) && !IsListOrderTypeChanged(previous, current))//TODO fix this
                    {
                        sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                    }

                    sb.Append(BuildInnerHtml(current, isListItem));

                    //TODO fix this
                    if (IsEndOfNestedList(previous, current, next))
                    {
                        sb.Append(isOrderListItem ? "</ol>" : "</ul>");
                    }

                    if (IsLastListItem(current, next))
                    {
                        sb.Append(isOrderListItem ? "</ol>" : "</ul>");
                    }

                    sb.Append(ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, current));
 
                }
            }
            
            return sb.ToString();
        }

        private static bool IsEndOfNestedList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current, OpenXmlParagraphWrapper? next)
        {
            
            if (next == null && current?.PPr?.Lvl > 0)
            {
                return true;
            }

            return current?.PPr?.Lvl > next?.PPr?.Lvl;
        }

        private static bool IsLastListItem(OpenXmlParagraphWrapper? current, OpenXmlParagraphWrapper? next)
        {
            return IsListItem(current) && next == null;
        }

        private static bool IsListOrderTypeChanged(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current)
        {
            return IsUnOrderedListItem(previous) && IsOrderedListItem(current) ||
                   IsOrderedListItem(previous) && IsUnOrderedListItem(current);
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

        private static bool IsUnOrderedListItem(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper?.PPr?.BuChar == null) return false;
            return IsListItem(paragraphWrapper) && paragraphWrapper.PPr.BuChar.Char == "•";
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
