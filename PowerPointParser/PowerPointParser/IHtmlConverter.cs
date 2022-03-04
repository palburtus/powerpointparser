using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser.Dto;

namespace PowerPointParser
{
    public interface IHtmlConverter
    {
        string? ConvertOpenXmlParagraphWrapperToHtml(OpenXmlParagraphWrapper? paragraphWrapper);
    }
}
