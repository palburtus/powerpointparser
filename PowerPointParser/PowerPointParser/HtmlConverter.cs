using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser.Dto;

namespace PowerPointParser
{
    public class HtmlConverter : IHtmlConverter
    {
        public string? ConvertOpenXmlParagraphWrapperToHtml(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            //if (paragraphWrapper?.R == null) return null;
            //return paragraphWrapper.R.T ?? null;

            if (paragraphWrapper == null) return null;
            if (paragraphWrapper.R == null) return null;
            if (paragraphWrapper.R.T == null) return null;
            return paragraphWrapper.R.T;
        }
    }
}
