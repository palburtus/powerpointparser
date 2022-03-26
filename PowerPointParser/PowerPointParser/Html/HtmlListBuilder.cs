using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public class HtmlListBuilder : IHtmlListBuilder
    {
        private readonly IInnerHtmlBuilder _innerHtmlBuilder;
       
        public HtmlListBuilder(IInnerHtmlBuilder innerHtmlBuilder)
        {
            _innerHtmlBuilder = innerHtmlBuilder;
        }

        Stack<string> _closingListBracketsStack = new Stack<string>();
        private string _closingBracket;

        public string BuildList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper current,
            OpenXmlParagraphWrapper? next)
        {
            StringBuilder sb = new();
            bool isOrderListItem = IsOrderedListItem(current);
            bool isLastOrderTypeChange = IsListOrderTypeChanged(previous, current);
            bool isNextOrderTypeChange = IsListOrderTypeChanged(current, next);

            if (isLastOrderTypeChange)
            {
                if (IsNotNested(next))
                {
                    if (_closingListBracketsStack.Count > 0)
                    {
                        sb.Append(_closingListBracketsStack.Pop());
                        sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                        _closingBracket = isOrderListItem ? "</ol>" : "</ul>";
                    }
                    
                    //sb.Append(isOrderListItem ? "</ul><ol>" : "</ol><ul>");
                    //_closingBracket = isOrderListItem ? "</ol>" : "</ul>";
                } 
            }

            if (IsFirstListItem(current, previous)) 
            {
                sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                _closingListBracketsStack.Push(isOrderListItem ? "</ol>" : "</ul>");
                _closingBracket = isOrderListItem ? "</ol>" : "</ul>";
            }

            if (IsStartOfNestedList(previous, current, next))
            {
                sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                _closingListBracketsStack.Push(isOrderListItem ? "</ol>" : "</ul>");
            }

            sb.Append(_innerHtmlBuilder.BuildInnerHtmlListItem(current));

            if (IsEndOfNestedList(previous, current, next))
            {
                int nextLevel = next !=null ? next.PPr!.Lvl : 0;

                for (int i = nextLevel; i < current.PPr?.Lvl; i++)
                {
                    if (_closingListBracketsStack.Count > 0)
                    {
                        sb.Append(_closingListBracketsStack.Pop());
                    }
                }
            }

            if (IsLastListItem(current, next))
            {
                sb.Append(_closingBracket);
            }

            return sb.ToString();
        }

        private static bool IsNotNested(OpenXmlParagraphWrapper? next)
        {
            return next?.PPr?.Lvl == 0;
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

        private bool IsStartOfNestedList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current, OpenXmlParagraphWrapper? next)
        {
            if (previous == null && current?.PPr?.Lvl > 0)
            {
                return true;
            }

            return current?.PPr?.Lvl > previous?.PPr?.Lvl;
        }

        private bool IsEndOfNestedList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current, OpenXmlParagraphWrapper? next)
        {
            if ((next == null || !IsListItem(next)) && current?.PPr?.Lvl > 0)
            {
                return true;
            }

            return current?.PPr?.Lvl > next?.PPr?.Lvl;
        }

        private  bool IsLastListItem(OpenXmlParagraphWrapper? current, OpenXmlParagraphWrapper? next)
        {
            return IsListItem(current) && (next == null || !IsListItem(next));
        }

        private bool IsFirstListItem(OpenXmlParagraphWrapper? current, OpenXmlParagraphWrapper? previous)
        {
            return IsListItem(current) && (previous == null || !IsListItem(previous));
        }

        private bool IsListOrderTypeChanged(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current)
        {
            return IsUnOrderedListItem(previous) && IsOrderedListItem(current) ||
                   IsOrderedListItem(previous) && IsUnOrderedListItem(current);
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
