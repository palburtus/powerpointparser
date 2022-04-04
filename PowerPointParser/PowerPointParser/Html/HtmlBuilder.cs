using System.Collections.Generic;
using System.Text;
using Aaks.PowerPointParser.Dto;
using Aaks.PowerPointParser.Extensions;

namespace Aaks.PowerPointParser.Html
{
    public class HtmlBuilder : IHtmlBuilder
    {
        private readonly IHtmlListBuilder _htmlListBuilder;
        private readonly IInnerHtmlBuilder _innerHtmlBuilder;

        public HtmlBuilder(IHtmlListBuilder htmlListBuilder, IInnerHtmlBuilder innerHtmlBuilder)
        {
            _htmlListBuilder = htmlListBuilder;
            _innerHtmlBuilder = innerHtmlBuilder;
        }

        public string ConvertOpenXmlParagraphWrapperToHtml(Queue<OpenXmlTextWrapper?>? paragraphWrappers)
        {
            return ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, null);
        }
        public Dictionary<int, string> ConvertOpenXmlParagraphWrapperToHtml(IDictionary<int, IList<OpenXmlTextWrapper?>> paragraphWrappers)
        {
            
            var speakerNotes = new Dictionary<int, string>();
            foreach (var (key, list) in paragraphWrappers)
            {
                var openXmlParagraphWrappers = list!.ToQueue();
                var htmlParagraph = ConvertHtmlParagraphWrapperToHtml(openXmlParagraphWrappers!, null);
                speakerNotes[key] = htmlParagraph;
            }
            return speakerNotes;
        }
        private string ConvertHtmlParagraphWrapperToHtml(Queue<OpenXmlTextWrapper?>? paragraphWrappers, OpenXmlTextWrapper? previous)
        {
            StringBuilder sb = new();
            if (paragraphWrappers == null) { return sb.ToString(); }
            
            
            while (paragraphWrappers.Count > 0)
            {
                var current = paragraphWrappers.Dequeue();
                paragraphWrappers.TryPeek(out var next);

                if (current?.R == null) { return sb.ToString(); }
                
                bool isListItem = _htmlListBuilder.IsListItem(current);

                if (!isListItem)
                {
                    sb.Append(_innerHtmlBuilder.BuildInnerHtmlParagraph(current));
                    sb.Append(ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, current));
                }
                else
                {
                    sb.Append(_htmlListBuilder.BuildList(previous, current, next));
                    sb.Append(ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, current));
                }
            }
            
            return sb.ToString();
        }


    }
}
