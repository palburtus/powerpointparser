
namespace Aaks.PowerPointParser.Html
{
    public static class HtmlTags
    {
        public static string Open(string type, string alignment = "left", string? style = null)
        {
            if (style == null)
            {
                return $"<{type}{(alignment == TextAlignment.Left ? string.Empty : TextAlignment.Align(alignment))}>";
            }
            else
            {
                return $"<{type}{(alignment == TextAlignment.Left ? string.Empty : TextAlignment.Align(alignment))} style=\"{style}\">";
            }
        } 
            

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
        public const string DoubleStrikeThrough = "del";
        public const string DoubleStrikeThroughTextDecoration = "text-decoration-style: double;";
        public const string LineBreak = "<br/>";
    }
}
