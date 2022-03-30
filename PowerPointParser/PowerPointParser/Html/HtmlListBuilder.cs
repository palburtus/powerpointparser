using System.Collections.Generic;
using System.Text;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public class HtmlListBuilder : IHtmlListBuilder
    {
        private readonly Stack<string> _closingListBracketsStack;

        private readonly IInnerHtmlBuilder _innerHtmlBuilder;
       
        public HtmlListBuilder(IInnerHtmlBuilder innerHtmlBuilder)
        {
            _closingListBracketsStack = new Stack<string>();
            _innerHtmlBuilder = innerHtmlBuilder;
        }

        public string BuildList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper current,
            OpenXmlParagraphWrapper? next)
        {
            StringBuilder sb = new();
            bool isOrderListItem = IsOrderedListItem(current);
            bool isLastOrderTypeChange = IsListOrderTypeChanged(previous, current);
           
            if (isLastOrderTypeChange)
            {

                if (IsNotNested(next) || IsLastListItemForLevel(previous, current, next))
                {
                    sb.Append(_closingListBracketsStack.Pop());
                    sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                    _closingListBracketsStack.Push(isOrderListItem ? "</ol>" : "</ul>");
                }
                else if(previous?.PPr?.Lvl == current.PPr?.Lvl && current.PPr?.Lvl == next?.PPr?.Lvl)
                {
                    sb.Append(_closingListBracketsStack.Pop());
                    sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                    _closingListBracketsStack.Push(isOrderListItem ? "</ol>" : "</ul>");
                }
            }

            if (IsFirstListItem(current, previous))
            {
                sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                _closingListBracketsStack.Push(isOrderListItem ? "</ol>" : "</ul>");
            }

            if (IsStartOfNestedList(previous, current))
            {
                int previousLevel = GetPreviousNestingLevel(previous);
                int currentLevel = GetCurrentNestingLevel(current);
                if (previousLevel < currentLevel)
                {
                    for (int i = previousLevel; i < current.PPr?.Lvl; i++)
                    {
                        sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                        _closingListBracketsStack.Push(isOrderListItem ? "</ol>" : "</ul>");
                    }
                }
                else
                {
                    sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                    _closingListBracketsStack.Push(isOrderListItem ? "</ol>" : "</ul>");
                }
                //sb.Append(isOrderListItem ? "<ol>" : "<ul>");
                //_closingListBracketsStack.Push(isOrderListItem ? "</ol>" : "</ul>");
            }

            sb.Append(_innerHtmlBuilder.BuildInnerHtmlListItem(current));

            if (IsEndOfNestedList(current, next))
            {
                int nextLevel = GetNextNestingLevel(next);

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
                for (int i = 0; i < _closingListBracketsStack.Count; i++)
                {
                    if (_closingListBracketsStack.Count > 0)
                    {
                        sb.Append(_closingListBracketsStack.Pop());
                    }
                }
            }

            return sb.ToString();
        }

        private int GetCurrentNestingLevel(OpenXmlParagraphWrapper current)
        {
            return current is { PPr: { } } ? current.PPr!.Lvl : 0; ;
        }

        private int GetNextNestingLevel(OpenXmlParagraphWrapper? next)
        {
            return next is {PPr: { }} ? next.PPr!.Lvl : 0;
        }

        private int GetPreviousNestingLevel(OpenXmlParagraphWrapper? previous)
        {
            return previous is {PPr: { }} ? previous.PPr!.Lvl : 0;
        }

        private bool IsLastListItemForLevel(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper current, OpenXmlParagraphWrapper? next)
        {
            if (next == null && current.PPr?.Lvl == previous?.PPr?.Lvl) return true;

            if (next != null && IsListOrderTypeChanged(previous, current) &&
                IsListOrderTypeChanged(current, next) &&
                current.PPr?.Lvl == previous?.PPr?.Lvl) return true;


            if (next != null && 
                IsListOrderTypeChanged(previous, current) && 
                IsListOrderTypeChanged(current, next) &&
                IsListOrderTypeChanged(previous, next) &&
                previous?.PPr?.Lvl < current.PPr?.Lvl) return true;

            
            return false;
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

        private bool IsStartOfNestedList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper? current)
        {
            if (previous == null && current?.PPr?.Lvl > 0)
            {
                return true;
            }

            return current?.PPr?.Lvl > previous?.PPr?.Lvl;
        }

        private bool IsEndOfNestedList(OpenXmlParagraphWrapper? current, OpenXmlParagraphWrapper? next)
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
