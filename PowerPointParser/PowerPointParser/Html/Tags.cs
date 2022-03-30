using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aaks.PowerPointParser.Html
{
    public static class Tags
    {
        public static string Open(string type) => $"<{type}>";
        public static string Close(string type) => $"</{type}>";
        public static string Paragraph = "p";
        public static string ListItem = "li";
        public static string UnorderedList = "ul";
        public static string OrderedList = "ol";
        public static string Bold = "strong";
        public static string Italic = "i";
        public static string Underlined = "u";
        public static string StrikeThrough = "del";
    }
}
