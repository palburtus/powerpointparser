using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public class NestedHtmlListBuilder : INestedHtmlListBuilder
    {
        public bool DoNotCloseListItemDueToNesting(OpenXmlTextWrapper? current, OpenXmlTextWrapper? next)
        {

            if (next?.PPr?.Lvl > current?.PPr?.Lvl)
            {
                return true;
            }

            return false;
        }

        public bool IsOnSameNestingLevel(OpenXmlTextWrapper? previous, OpenXmlTextWrapper? current)
        {
            if (current == null) return true;
            if (previous?.PPr == null && current?.PPr == null) return true;
            return previous?.PPr?.Lvl == current?.PPr?.Lvl;
        }

        public bool ShouldChangeListTypes(OpenXmlTextWrapper? previous, OpenXmlTextWrapper? current, OpenXmlTextWrapper? next, string closingBracketTop)
        {
            if (IsListOrderTypeChanged(previous, current))
            {
                if (IsOnSameNestingLevel(previous, current)) return true;
                if (current?.PPr?.Lvl > previous?.PPr?.Lvl) return false;
                if (current?.PPr?.Lvl < previous?.PPr?.Lvl)
                {
                    if (IsClosingUnorderedList(closingBracketTop) && !IsUnOrderedListItem(current))
                    {
                        return true;
                    }
                    
                    if (IsClosingOrderedList(closingBracketTop) && !IsOrderedListItem(current))
                    {
                        return true;
                    }
                    
                }

            }
            

            return false;
        }

        private bool IsListOrderTypeChanged(OpenXmlTextWrapper? previous, OpenXmlTextWrapper? current)
        {
            return IsUnOrderedListItem(previous) && IsOrderedListItem(current) ||
                   IsOrderedListItem(previous) && IsUnOrderedListItem(current);
        }

        private bool IsClosingUnorderedList(string bracket)
        {
            return bracket.Contains("ul");
        }

        private bool IsClosingOrderedList(string bracket)
        {
            return bracket.Contains("0l");
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
