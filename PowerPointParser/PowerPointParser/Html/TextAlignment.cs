
namespace Aaks.PowerPointParser.Html
{
    public static class TextAlignment
    {
        public static string Align(string direction) => $" style=\"text-align: {direction};\"";
        public const string Left = "left";
        public const string Right = "right";
        public const string Center = "center";
        public const string Justify = "justify";
    }
}
