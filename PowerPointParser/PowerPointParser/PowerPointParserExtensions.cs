using System.Collections.Generic;
using System.Linq;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser
{

    public static class PowerPointParserExtensions
    {
        public static Queue<OpenXmlParagraphWrapper> ToQueue(this IDictionary<int, IList<OpenXmlParagraphWrapper>> items)
        {
            Queue<OpenXmlParagraphWrapper> openXmlParagraphWrappers = new();
            var xmlParagraphWrappers = items.Select(x => x.Value).SelectMany(y => y).ToList();
            foreach (var openXmlParagraphWrapper in xmlParagraphWrappers)
            {
                openXmlParagraphWrappers.Enqueue(openXmlParagraphWrapper);
            }
            return openXmlParagraphWrappers;
        }
        public static Queue<OpenXmlParagraphWrapper> ToQueue(this IList<OpenXmlParagraphWrapper> items)
        {
            Queue<OpenXmlParagraphWrapper> openXmlParagraphWrappers = new();
            foreach (var current in items)
            {
                openXmlParagraphWrappers.Enqueue(current);
            }
            return openXmlParagraphWrappers;
        }
    }
}