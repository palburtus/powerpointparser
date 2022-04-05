using System.Collections.Generic;
using System.IO;
using Aaks.PowerPointParser.Dto;
using DocumentFormat.OpenXml.Drawing;

namespace Aaks.PowerPointParser.Parsers
{
    public interface IPowerPointParser
    {
        IDictionary<int, IList<Paragraph?>> ParseSlide(string path);
        IDictionary<int, IList<OpenXmlTextWrapper?>> ParseSpeakerNotes(string path);
        IDictionary<int, IList<OpenXmlTextWrapper?>> ParseSpeakerNotes(MemoryStream memoryStream);
    }
}
