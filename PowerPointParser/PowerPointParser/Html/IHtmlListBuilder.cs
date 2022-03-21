using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Html
{
    public interface IHtmlListBuilder
    {
        string BuildList(OpenXmlParagraphWrapper? previous, OpenXmlParagraphWrapper current, OpenXmlParagraphWrapper? next);
        bool IsListItem(OpenXmlParagraphWrapper? paragraphWrapper);
    }
}
