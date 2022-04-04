
namespace Aaks.PowerPointParser.Html
{
    public static class TextAlignment
    {
        public static string Align(string direction) => $" style=\"text-align: {direction};\"";
        public static string Left = "left";
        public static string Right = "right";
        public static string Center = "center";
        public static string Justify = "justify";
    }
}
