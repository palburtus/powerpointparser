using System.Collections.Generic;
using System.Linq;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser
{

    public static class PowerPointParserExtensions
    {
        public static Queue<OpenXmlTextWrapper> ToQueue(this IDictionary<int, IList<OpenXmlTextWrapper>> items)
        {
            Queue<OpenXmlTextWrapper> openXmlParagraphWrappers = new();
            var xmlParagraphWrappers = items.Select(x => x.Value).SelectMany(y => y).ToList();
            foreach (var openXmlParagraphWrapper in xmlParagraphWrappers)
            {
                openXmlParagraphWrappers.Enqueue(openXmlParagraphWrapper);
            }
            return openXmlParagraphWrappers;
        }
        public static Queue<OpenXmlTextWrapper> ToQueue(this IList<OpenXmlTextWrapper> items)
        {
            Queue<OpenXmlTextWrapper> openXmlParagraphWrappers = new();
            foreach (var current in items)
            {
                openXmlParagraphWrappers.Enqueue(current);
            }
            return openXmlParagraphWrappers;
        }
    }
}