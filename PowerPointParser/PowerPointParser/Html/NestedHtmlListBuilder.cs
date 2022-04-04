using Aaks.PowerPointParser.Dto;
using Aaks.PowerPointParser.Extensions;

namespace Aaks.PowerPointParser.Html
{
    public class NestedHtmlListBuilder : INestedHtmlListBuilder
    {
        public bool DoNotCloseListItemDueToNesting(OpenXmlTextWrapper? current, OpenXmlTextWrapper? next)
        {
            return next?.PPr?.Lvl > current?.PPr?.Lvl;
        }

        public bool IsOnSameNestingLevel(OpenXmlTextWrapper? previous, OpenXmlTextWrapper? current)
        {
            if (current == null) return true;
            if (previous?.PPr == null && current.PPr == null) return true;
            return previous?.PPr?.Lvl == current.PPr?.Lvl;
        }

        public bool ShouldChangeListTypes(OpenXmlTextWrapper? previous, OpenXmlTextWrapper? current, OpenXmlTextWrapper? next, string closingBracketTop)
        {
            if (!IsListOrderTypeChanged(previous, current))
            {
                if (IsClosingOrderedList(closingBracketTop) != current.IsOrderedListItem()) return true;
                if (IsClosingUnorderedList(closingBracketTop) != current.IsUnOrderedListItem()) return true;
            }
            else
            {
                if (IsOnSameNestingLevel(previous, current)) return true;
                if (current?.PPr?.Lvl > previous?.PPr?.Lvl) return false;
                if (!(current?.PPr?.Lvl < previous?.PPr?.Lvl)) return false;
                if (IsClosingUnorderedList(closingBracketTop) && !current.IsUnOrderedListItem()) return true;
                if (IsClosingOrderedList(closingBracketTop) && !current.IsOrderedListItem()) return true;
            }

            return false;
        }

        private bool IsListOrderTypeChanged(OpenXmlTextWrapper? previous, OpenXmlTextWrapper? current)
        {
            return previous.IsUnOrderedListItem() && current.IsOrderedListItem() ||
                   previous.IsOrderedListItem() && current.IsUnOrderedListItem();
        }

        private bool IsClosingUnorderedList(string bracket)
        {
            return bracket.Contains(HtmlTags.UnorderedList);
        }

        private bool IsClosingOrderedList(string bracket)
        {
            return bracket.Contains(HtmlTags.OrderedList);
        }

    }
}
