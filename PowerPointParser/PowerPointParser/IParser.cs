using System;
using PowerPointParser.Dto;

namespace PowerPointParser
{
    public interface IParser
    {
        IDictionary<int, IList<OpenXmlParagraphWrapper>> ParseSpeakerNotes(string path);
    }
}
