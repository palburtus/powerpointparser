using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser.Dto;

namespace PowerPointParser.Html
{
    public interface IInnerHtmlBuilder
    {
        string BuildInnerHtmlParagraph(OpenXmlParagraphWrapper paragraphWrapper);
        string BuildInnerHtmlListItem(OpenXmlParagraphWrapper paragraphWrapper);
    }
}
