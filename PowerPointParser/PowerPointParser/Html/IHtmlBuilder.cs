using System.Collections.Generic;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface IHtmlBuilder
    {
        string? ConvertOpenXmlParagraphWrapperToHtml(Queue<OpenXmlTextWrapper?>? paragraphWrapper);

        public Dictionary<int, string> ConvertOpenXmlParagraphWrapperToHtml(
            IDictionary<int, IList<OpenXmlTextWrapper?>> paragraphWrappers);
    }
}
