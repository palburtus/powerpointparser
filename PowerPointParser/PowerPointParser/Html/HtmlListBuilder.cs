﻿using System.Collections.Generic;
using System.Text;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public class HtmlListBuilder : IHtmlListBuilder
    {
        private readonly Stack<string> _closingListBracketsStack;
        private Stack<string> _closingListItemBracketsStack;

        private readonly IInnerHtmlBuilder _innerHtmlBuilder;
        private readonly INestedHtmlListBuilder _nestedHtmlListBuilder;
       
        public HtmlListBuilder(IInnerHtmlBuilder innerHtmlBuilder)
        {
            _closingListBracketsStack = new Stack<string>();
            _closingListItemBracketsStack = new Stack<string>();

            _innerHtmlBuilder = innerHtmlBuilder;
            _nestedHtmlListBuilder = new NestedHtmlListBuilder();
        }

        public string BuildList(OpenXmlTextWrapper? previous, OpenXmlTextWrapper current,
            OpenXmlTextWrapper? next)
        {
            StringBuilder sb = new();
            bool isOrderListItem = IsOrderedListItem(current);
            //bool isLastOrderTypeChange = IsListOrderTypeChanged(previous, current);
           
            if (_closingListBracketsStack.Count > 0 &&
                _nestedHtmlListBuilder.ShouldChangeListTypes(previous, current, next, _closingListBracketsStack.Peek()))
            {

                sb.Append(_closingListBracketsStack.Pop());
                sb.Append(isOrderListItem ? HtmlTags.Open(HtmlTags.OrderedList) : HtmlTags.Open(HtmlTags.UnorderedList));
                _closingListBracketsStack.Push(isOrderListItem ? HtmlTags.Close(HtmlTags.OrderedList) : HtmlTags.Close(HtmlTags.UnorderedList));

                //IsNotNested(current))

                //if (_nestedHtmlListBuilder.IsOnSameNestingLevel(current, next) || (IsLastListItemForLevel(previous, current, next) || (previous?.PPr?.Lvl == current.PPr?.Lvl && current.PPr?.Lvl == next?.PPr?.Lvl) /*&& !IsOnlyLevelForListItem(previous, current, next)*/))
                //{

                //}

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

            //sb.Append(_innerHtmlBuilder.BuildInnerHtmlListItem(current));

            bool isNextStartOfNesting = _nestedHtmlListBuilder.DoNotCloseListItemDueToNesting(current, next);
            bool isNestedWithoutParent = next?.PPr?.Lvl - current.PPr?.Lvl > 1;

            if (isNextStartOfNesting && !isNestedWithoutParent)
            {
                sb.Append(_innerHtmlBuilder.BuildInnerHtmlListItemBeforeNesting(current));
                _closingListItemBracketsStack.Push("</li>");
            }
            else
            {
                sb.Append(_innerHtmlBuilder.BuildInnerHtmlListItem(current));
            }

            //sb.Append(isNextStartOfNesting && !IsListOrderTypeChanged(current, next) ? _innerHtmlBuilder.BuildInnerHtmlListItemBeforeNesting(current) :
              //

            if (IsEndOfNestedList(current, next))
            {
                int nextLevel = GetNextNestingLevel(next);

                for (int i = nextLevel; i < current.PPr?.Lvl; i++)
                {
                    if (_closingListBracketsStack.Count > 0)
                    {
                        sb.Append(_closingListBracketsStack.Pop());
                            //sb.Append("</li>");
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

        private int GetCurrentNestingLevel(OpenXmlTextWrapper current)
        {
            return current is { PPr: { } } ? current.PPr!.Lvl : 0; 
        }

        private int GetNextNestingLevel(OpenXmlTextWrapper? next)
        {
            return next is {PPr: { }} ? next.PPr!.Lvl : 0;
        }

        private int GetPreviousNestingLevel(OpenXmlTextWrapper? previous)
        {
            return previous is {PPr: { }} ? previous.PPr!.Lvl : 0;
        }

        /*private bool IsLastListItemForLevel(OpenXmlTextWrapper? previous, OpenXmlTextWrapper current, OpenXmlTextWrapper? next)
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
        }*/

        

        public bool IsListItem(OpenXmlTextWrapper? paragraphWrapper)
        {
            return IsUnOrderedListItem(paragraphWrapper) || IsOrderedListItem(paragraphWrapper);
        }

        private bool IsStartOfNestedList(OpenXmlTextWrapper? previous, OpenXmlTextWrapper? current)
        {
            if (previous == null && current?.PPr?.Lvl > 0)
            {
                return true;
            }

            return current?.PPr?.Lvl > previous?.PPr?.Lvl;
        }

        private bool IsEndOfNestedList(OpenXmlTextWrapper? current, OpenXmlTextWrapper? next)
        {
            if ((next == null || !IsListItem(next)) && current?.PPr?.Lvl > 0)
            {
                return true;
            }

            return current?.PPr?.Lvl > next?.PPr?.Lvl;
        }

        private  bool IsLastListItem(OpenXmlTextWrapper? current, OpenXmlTextWrapper? next)
        {
            return IsListItem(current) && (next == null || !IsListItem(next));
        }

        private bool IsFirstListItem(OpenXmlTextWrapper? current, OpenXmlTextWrapper? previous)
        {
            return IsListItem(current) && (previous == null || !IsListItem(previous));
        }

        

        private bool IsUnOrderedListItem(OpenXmlTextWrapper? paragraphWrapper)
        {
            return paragraphWrapper?.PPr?.BuChar?.Char is 
                OpenXmlTextModifiers.UlFilledRoundBullet or 
                OpenXmlTextModifiers.UlHollowRoundBullet or 
                OpenXmlTextModifiers.UlFilledSquareBullet or
                OpenXmlTextModifiers.UlHollowSquareBullet or
                OpenXmlTextModifiers.UlStarBullet or
                OpenXmlTextModifiers.UlArrowBullet or 
                OpenXmlTextModifiers.UlCheckmarkBullet;
        }

        private bool IsOrderedListItem(OpenXmlTextWrapper? paragraphWrapper)
        {
            return paragraphWrapper?.PPr?.BuAutoNum?.Type is 
                OpenXmlTextModifiers.OlArabicPeriod or
                OpenXmlTextModifiers.OlArabicParenRight or
                OpenXmlTextModifiers.OlCapitalRomanNumeralsPeriod or 
                OpenXmlTextModifiers.OlCapitalAlphaPeriod or
                OpenXmlTextModifiers.OlLowercaseAlphaRightParen or
                OpenXmlTextModifiers.OlLowerCaseAlphaPeriod or 
                OpenXmlTextModifiers.OlLowercaseRomanNumeralsPeriod;
        }
    }
}
