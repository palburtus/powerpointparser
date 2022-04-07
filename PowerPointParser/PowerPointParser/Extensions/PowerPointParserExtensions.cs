using System.Collections.Generic;
using System.Linq;
using Aaks.PowerPointParser.Dto;

namespace Aaks.PowerPointParser.Extensions
{

    public static class PowerPointParserExtensions
    {
        public static Queue<OpenXmlLineItem> ToQueue(this IDictionary<int, IList<OpenXmlLineItem>> items)
        {
            Queue<OpenXmlLineItem> openXmlParagraphWrappers = new();
            var xmlParagraphWrappers = items.Select(x => x.Value).SelectMany(y => y).ToList();
            foreach (var openXmlParagraphWrapper in xmlParagraphWrappers)
            {
                openXmlParagraphWrappers.Enqueue(openXmlParagraphWrapper);
            }
            return openXmlParagraphWrappers;
        }
        public static Queue<OpenXmlLineItem> ToQueue(this IList<OpenXmlLineItem> items)
        {
            Queue<OpenXmlLineItem> openXmlParagraphWrappers = new();
            foreach (var current in items)
            {
                openXmlParagraphWrappers.Enqueue(current);
            }
            return openXmlParagraphWrappers;
        }
    }
}