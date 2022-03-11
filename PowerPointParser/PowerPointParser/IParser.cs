using System;
using System.Collections.Generic;
using PowerPointParser.Model;

namespace PowerPointParser
{
    public interface IParser
    {
        IList<Slide> ParseSpeakerNotes(string path);
    }
}
