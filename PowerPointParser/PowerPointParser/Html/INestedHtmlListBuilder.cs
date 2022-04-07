using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface INestedHtmlListBuilder
    {
        bool DoNotCloseListItemDueToNesting(OpenXmlLineItem? current, OpenXmlLineItem? next);
        bool IsOnSameNestingLevel(OpenXmlLineItem? previous, OpenXmlLineItem? current);
        bool ShouldChangeListTypes(OpenXmlLineItem? previous, OpenXmlLineItem? current, OpenXmlLineItem? next, string closingBracketTop);

    }
}
