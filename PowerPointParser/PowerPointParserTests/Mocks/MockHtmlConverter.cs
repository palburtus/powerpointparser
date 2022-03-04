using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointParser;
using PowerPointParser.Dto;

namespace PowerPointParserTests.Mocks
{
    public class MockHtmlConverter : IHtmlConverter
    {
        public string? ConvertOpenXmlParagraphWrapperToHtml(OpenXmlParagraphWrapper? paragraphWrapper)
        {
            if (paragraphWrapper == null) return null;
            if(paragraphWrapper.R == null) return null;
            if(paragraphWrapper.R.T == null) return null;
            return paragraphWrapper.R.T;
        }
    }
}
