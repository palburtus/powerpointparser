
namespace Aaks.PowerPointParser.Html
{
    public static class HtmlTags
    {
        public static string Open(string type, string alignment = "left") => 
            $"<{type}{(alignment == TextAlignment.Left ? string.Empty : TextAlignment.Align(alignment))}>";
        public static string Close(string type, string alignment = "left") => 
            $"</{type}{(alignment == TextAlignment.Left ? string.Empty : TextAlignment.Align(alignment))}>";
        public static readonly string Paragraph = "p";
        public static readonly string ListItem = "li";
        public static readonly string UnorderedList = "ul";
        public static readonly string OrderedList = "ol";
        public static readonly string Bold = "strong";
        public static readonly string Italic = "i";
        public static readonly string Underlined = "u";
        public static readonly string StrikeThrough = "del";
    }
}
