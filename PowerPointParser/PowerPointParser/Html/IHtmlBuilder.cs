using System.Collections.Generic;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface IHtmlBuilder
    {
        string? ConvertOpenXmlParagraphWrapperToHtml(Queue<OpenXmlLineItem?>? paragraphWrapper);

        public Dictionary<int, string> ConvertOpenXmlParagraphWrapperToHtml(
            IDictionary<int, IList<OpenXmlLineItem?>> paragraphWrappers);
    }
}
