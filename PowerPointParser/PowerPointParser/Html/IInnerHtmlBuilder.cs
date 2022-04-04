using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface IInnerHtmlBuilder
    {
        string BuildInnerHtmlParagraph(OpenXmlTextWrapper textWrapper);
        string BuildInnerHtmlListItem(OpenXmlTextWrapper textWrapper);
        string BuildInnerHtmlListItemBeforeNesting(OpenXmlTextWrapper textWrapper);
    }
}
