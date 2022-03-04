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
            if (paragraphWrapper == null) return null;
            if (paragraphWrapper.R == null) return null;
            if (paragraphWrapper.R.Count == 0) return null;

            StringBuilder sb = new StringBuilder();

            foreach (var r in paragraphWrapper.R)
            {
                sb.Append(r.T);
            }

            return sb.ToString();
        }
    }
}
