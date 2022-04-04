using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface INestedHtmlListBuilder
    {
        bool DoNotCloseListItemDueToNesting(OpenXmlTextWrapper? current, OpenXmlTextWrapper? next);
        bool IsOnSameNestingLevel(OpenXmlTextWrapper? previous, OpenXmlTextWrapper? current);
        bool ShouldChangeListTypes(OpenXmlTextWrapper? previous, OpenXmlTextWrapper? current, OpenXmlTextWrapper? next, string closingBracketTop);

    }
}
