using System.Collections.Generic;
using System.Text;
using Aaks.PowerPointParser.Dto;
using Aaks.PowerPointParser.Extensions;

namespace Aaks.PowerPointParser.Html
{
    public class HtmlListBuilder : IHtmlListBuilder
    {
        private readonly Stack<string> _closingListBracketsStack;
        private readonly Stack<string> _closingListItemBracketsStack;

        private readonly IInnerHtmlBuilder _innerHtmlBuilder;
        private readonly INestedHtmlListBuilder _nestedHtmlListBuilder;
       
        public HtmlListBuilder(IInnerHtmlBuilder innerHtmlBuilder)
        {
            _closingListBracketsStack = new Stack<string>();
            _closingListItemBracketsStack = new Stack<string>();

            _innerHtmlBuilder = innerHtmlBuilder;
            _nestedHtmlListBuilder = new NestedHtmlListBuilder();
        }

        public string BuildList(OpenXmlLineItem? previous, OpenXmlLineItem current, OpenXmlLineItem? next)
        {
            StringBuilder sb = new();
            bool isOrderListItem = current.IsOrderedListItem();
           
            if (_closingListBracketsStack.TryPeek(out var top))
            {
                if (_nestedHtmlListBuilder.ShouldChangeListTypes(previous, current, next, top))
                {
                    sb.Append(_closingListBracketsStack.Pop());
                    sb.Append(isOrderListItem ? HtmlTags.Open(HtmlTags.OrderedList) : HtmlTags.Open(HtmlTags.UnorderedList));
                    _closingListBracketsStack.Push(isOrderListItem ? HtmlTags.Close(HtmlTags.OrderedList) : HtmlTags.Close(HtmlTags.UnorderedList));
                }
            }

            if (IsFirstListItem(current, previous))
            {
                sb.Append(isOrderListItem ? HtmlTags.Open(HtmlTags.OrderedList) : HtmlTags.Open(HtmlTags.UnorderedList));
                _closingListBracketsStack.Push(isOrderListItem ? HtmlTags.Close(HtmlTags.OrderedList) : HtmlTags.Close(HtmlTags.UnorderedList));
            }

            if (IsStartOfNestedList(previous, current))
            {
                int previousLevel = GetPreviousNestingLevel(previous);
                int currentLevel = GetCurrentNestingLevel(current);
                if (previousLevel < currentLevel)
                {
                    for (int i = previousLevel; i < current.PPr?.Lvl; i++)
                    {
                        sb.Append(isOrderListItem ? HtmlTags.Open(HtmlTags.OrderedList) : HtmlTags.Open(HtmlTags.UnorderedList));
                        _closingListBracketsStack.Push(isOrderListItem ? HtmlTags.Close(HtmlTags.OrderedList) : HtmlTags.Close(HtmlTags.UnorderedList));
                    }
                }
                else
                {
                    sb.Append(isOrderListItem ? HtmlTags.Open(HtmlTags.OrderedList) : HtmlTags.Open(HtmlTags.UnorderedList));
                    _closingListBracketsStack.Push(isOrderListItem ? HtmlTags.Close(HtmlTags.OrderedList) : HtmlTags.Close(HtmlTags.UnorderedList));
                }
            }

            bool isNextStartOfNesting = _nestedHtmlListBuilder.DoNotCloseListItemDueToNesting(current, next);
            bool isNestedWithoutParent = next?.PPr?.Lvl - current.PPr?.Lvl > 1;

            if (isNextStartOfNesting && !isNestedWithoutParent)
            {
                sb.Append(_innerHtmlBuilder.BuildInnerHtmlListItemBeforeNesting(current));
                _closingListItemBracketsStack.Push(HtmlTags.Close(HtmlTags.ListItem));
            }
            else
            {
                sb.Append(_innerHtmlBuilder.BuildInnerHtmlListItem(current));
            }

            if (IsEndOfNestedList(current, next))
            {
                int nextLevel = GetNextNestingLevel(next);

                for (int i = nextLevel; i < current.PPr?.Lvl; i++)
                {
                    if (_closingListBracketsStack.Count > 0)
                    {
                        sb.Append(_closingListBracketsStack.Pop());
                    }

                    if (_closingListItemBracketsStack.Count > 0)
                    {
                        sb.Append(_closingListItemBracketsStack.Pop());
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

        private static int GetCurrentNestingLevel(OpenXmlLineItem current)
        {
            return current is { PPr: { } } ? current.PPr!.Lvl : 0; 
        }

        private static int GetNextNestingLevel(OpenXmlLineItem? next)
        {
            return next is {PPr: { }} ? next.PPr!.Lvl : 0;
        }

        private static int GetPreviousNestingLevel(OpenXmlLineItem? previous)
        {
            return previous is {PPr: { }} ? previous.PPr!.Lvl : 0;
        }

        public bool IsListItem(OpenXmlLineItem? paragraphWrapper)
        {
            return paragraphWrapper.IsUnOrderedListItem() || paragraphWrapper.IsOrderedListItem();
        }

        private static bool IsStartOfNestedList(OpenXmlLineItem? previous, OpenXmlLineItem? current)
        {
            if (previous == null && current?.PPr?.Lvl > 0)
            {
                return true;
            }

            return current?.PPr?.Lvl > previous?.PPr?.Lvl;
        }

        private bool IsEndOfNestedList(OpenXmlLineItem? current, OpenXmlLineItem? next)
        {
            if ((next == null || !IsListItem(next)) && current?.PPr?.Lvl > 0)
            {
                return true;
            }

            return current?.PPr?.Lvl > next?.PPr?.Lvl;
        }

        private  bool IsLastListItem(OpenXmlLineItem? current, OpenXmlLineItem? next)
        {
            return IsListItem(current) && (next == null || !IsListItem(next));
        }

        private bool IsFirstListItem(OpenXmlLineItem? current, OpenXmlLineItem? previous)
        {
            return IsListItem(current) && (previous == null || !IsListItem(previous));
        }
    }
}
