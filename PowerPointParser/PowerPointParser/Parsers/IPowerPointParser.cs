using System.Collections.Generic;
using System.IO;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Parsers
{
    public interface IPowerPointParser
    {
        IDictionary<int, IList<OpenXmlTextWrapper?>> ParseSpeakerNotes(string path);
        IDictionary<int, IList<OpenXmlTextWrapper?>> ParseSpeakerNotes(MemoryStream memoryStream);
    }
}
