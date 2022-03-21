using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface IInnerHtmlBuilder
    {
        string BuildInnerHtmlParagraph(OpenXmlParagraphWrapper paragraphWrapper);
        string BuildInnerHtmlListItem(OpenXmlParagraphWrapper paragraphWrapper);
    }
}
