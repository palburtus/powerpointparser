using System;
using PowerPointParser.Model;

namespace PowerPointParser
{
    public interface IParser
    {
        IList<Slide> ParseSpeakerNotes(string path);
    }
}
