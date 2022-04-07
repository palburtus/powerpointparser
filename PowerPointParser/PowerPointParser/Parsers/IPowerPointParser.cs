using System.Collections.Generic;
using System.IO;
using Aaks.PowerPointParser.Dto;
using DocumentFormat.OpenXml.Drawing;

namespace Aaks.PowerPointParser.Parsers
{
    public interface IPowerPointParser
    {
        IDictionary<int, IList<OpenXmlLineItem?>> ParseSlide(string path);
        IDictionary<int, IList<OpenXmlLineItem?>> ParseSpeakerNotes(string path);
        IDictionary<int, IList<OpenXmlLineItem?>> ParseSpeakerNotes(MemoryStream memoryStream);
    }
}
