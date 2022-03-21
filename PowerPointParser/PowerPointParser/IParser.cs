using System;

using System.Collections.Generic;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser
{
    public interface IParser
    {
        IDictionary<int, IList<OpenXmlParagraphWrapper?>> ParseSpeakerNotes(string path);
    }
}
