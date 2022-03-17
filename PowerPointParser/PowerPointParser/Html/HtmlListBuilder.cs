using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser.Dto;

namespace PowerPointParser.Html
{
    public class HtmlListBuilder : IHtmlListBuilder
    {
        private readonly IInnerHtmlBuilder _innerHtmlBuilder;

        public HtmlListBuilder(IInnerHtmlBuilder innerHtmlBuilder)
        {
            _innerHtmlBuilder = innerHtmlBuilder;
        }

        public string BuildList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper current,
            OpenXmlParagraphWrapper? next)
        {
            StringBuilder sb = new();
            bool isOrderListItem = IsOrderedListItem(current);

            if (IsListOrderTypeChanged(previous, current))
            {
                sb.Append(isOrderListItem ? "</ul><ol>" : "</ol><ul>");
            }


            if (IsFirstListItem(previous, current) && !IsListOrderTypeChanged(previous, current)) //TODO fix this
            {
                sb.Append(isOrderListItem ? "<ol>" : "<ul>");
            }

            sb.Append(_innerHtmlBuilder.BuildInnerHtmlListItem(current));

            //TODO fix this
            if (IsEndOfNestedList(previous, current, next))
            {
                sb.Append(isOrderListItem ? "</ol>" : "</ul>");
            }

            if (IsLastListItem(current, next))
            {
                sb.Append(isOrderListItem ? "</ol>" : "</ul>");
            }

            return sb.ToString();
        }

        public bool IsListItem(OpenXmlParagraphWrapper? paragraphWrapper)
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

        private bool IsEndOfNestedList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current, OpenXmlParagraphWrapper? next)
        {

            if (next == null && current?.PPr?.Lvl > 0)
            {
                return true;
            }

            return current?.PPr?.Lvl > next?.PPr?.Lvl;
        }

        private  bool IsLastListItem(OpenXmlParagraphWrapper? current, OpenXmlParagraphWrapper? next)
        {
            return IsListItem(current) && next == null;
        }

        private bool IsListOrderTypeChanged(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current)
        {
            return IsUnOrderedListItem(previous) && IsOrderedListItem(current) ||
                   IsOrderedListItem(previous) && IsUnOrderedListItem(current);
        }

        private  bool IsFirstListItem(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current)
        {
            if (current?.PPr?.Lvl > previous?.PPr?.Lvl)
            {
                return true;
            }

            return previous == null || !IsListItem(previous);
        }

        private bool IsUnOrderedListItem(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper?.PPr?.BuChar == null) return false;
            return IsListItem(paragraphWrapper) && paragraphWrapper.PPr.BuChar.Char == "•";
        }

        private bool IsOrderedListItem(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper?.PPr?.BuAutoNum?.Type == null) return false;
            return IsListItem(paragraphWrapper) && paragraphWrapper.PPr.BuAutoNum.Type == "arabicPeriod";
        }
    }
}
