using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface IHtmlListBuilder
    {
        string BuildList(OpenXmlLineItem? previous, OpenXmlLineItem current, OpenXmlLineItem? next);
        bool IsListItem(OpenXmlLineItem? paragraphWrapper);
    }
}
