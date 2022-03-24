using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface IHtmlBuilder
    {
        string? ConvertOpenXmlParagraphWrapperToHtml(Queue<OpenXmlParagraphWrapper?>? paragraphWrapper);

        public Dictionary<int, string> ConvertOpenXmlParagraphWrapperToHtml(
            IDictionary<int, IList<OpenXmlParagraphWrapper?>> paragraphWrappers);
    }
}
