using System.Collections.Generic;
using System.IO;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser
{
    public interface IParser
    {
        IDictionary<int, IList<OpenXmlParagraphWrapper?>> ParseSpeakerNotes(string path);
        IDictionary<int, IList<OpenXmlParagraphWrapper?>> ParseSpeakerNotes(MemoryStream memoryStream);
    }
}
