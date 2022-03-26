using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface IInnerHtmlBuilder
    {
        string BuildInnerHtmlParagraph(OpenXmlParagraphWrapper paragraphWrapper);
        string BuildInnerHtmlListItem(OpenXmlParagraphWrapper paragraphWrapper);
    }
}
