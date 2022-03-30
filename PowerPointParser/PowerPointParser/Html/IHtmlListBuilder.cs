using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface IHtmlListBuilder
    {
        string BuildList(OpenXmlTextWrapper? previous, OpenXmlTextWrapper current, OpenXmlTextWrapper? next);
        bool IsListItem(OpenXmlTextWrapper? paragraphWrapper);
    }
}
