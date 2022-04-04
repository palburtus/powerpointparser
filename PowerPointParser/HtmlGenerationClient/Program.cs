using System.Text;
using Aaks.PowerPointParser.Extensions;
using Aaks.PowerPointParser.Html;
using Aaks.PowerPointParser.Parsers;

Console.WriteLine("Enter Absolute Powerpoint File Path");
var powerpointFilePath = Console.ReadLine();

IPowerPointParser powerPointParser = new PowerPointParser();
var speakerNotesMap = powerPointParser.ParseSpeakerNotes(powerpointFilePath!);

IInnerHtmlBuilder innerHtmlBuilder = new InnerHtmlBuilder();
IHtmlBuilder builder = new HtmlBuilder(new HtmlListBuilder(innerHtmlBuilder), innerHtmlBuilder);


StringBuilder sb = new ();

long ticks = DateTime.UtcNow.Ticks;

sb.Append(GetOpeningHtml(ticks));

foreach (int key in speakerNotesMap.Keys)
{
    sb.Append("<div class=\"slide\">");

    sb.Append($"<h3>Slide - {key}</h3>");
    sb.Append("<hr>");

    sb.Append("<div class=\"speaker_note\">");

    sb.Append(builder.ConvertOpenXmlParagraphWrapperToHtml(speakerNotesMap[key]!.ToQueue()!));

    sb.Append("</div>");

    sb.Append("</div>");
}

sb.Append(GetClosingHtml());

string outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
string outputPath = Path.Combine(outputFolder, $"{ticks}.html");

using StreamWriter outputFile = new(outputPath);
outputFile.Write(sb.ToString());

static string GetOpeningHtml(long ticks)
{
    StringBuilder html = new();

    html.Append("<!DOCTYPE html>");
    html.Append("<html lang=\"en\">");
    html.Append("<head>");
    html.Append($"<title>{ticks}</title>");
    html.Append("<style>");
    html.Append(GetStyle());
    html.Append("</style>");
    html.Append("</head>");
    html.Append("<body>");

    return html.ToString();
}

static string GetStyle()
{
    StringBuilder css = new();

    css.Append(".slide { border: 1px solid #cacaca; margin: 5px 5px 15px 5px; width: 33%;}");
    css.Append(".speaker_note { margin: 5px 15px 15px 15px;}");
    css.Append(".speaker_note p {font-family: Arial, sans-serif; line-height: 2em;}");
    css.Append("h3 { color: #444444; margin: 15px 15px 5px 15px; font-size: 22px; font-family: Arial, sans-serif; }");
    css.Append("hr { border: 1px solid #cacaca; margin: 0px 15px 0px 15px;}");

    return css.ToString();
}

static string GetClosingHtml()
{
    StringBuilder html = new();

    html.Append("</body>");
    html.Append("</html>");

    return html.ToString();
}