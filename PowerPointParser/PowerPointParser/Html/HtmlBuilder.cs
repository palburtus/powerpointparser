using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aaks.PowerPointParser.Dto;

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

        public string ConvertOpenXmlParagraphWrapperToHtml(Queue<OpenXmlParagraphWrapper?>? paragraphWrappers)
        {
            return ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, null);
        }
        public Dictionary<int, string> ConvertOpenXmlParagraphWrapperToHtml(IDictionary<int, IList<OpenXmlParagraphWrapper?>> paragraphWrappers)
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
        private string ConvertHtmlParagraphWrapperToHtml(Queue<OpenXmlParagraphWrapper?>? paragraphWrappers, OpenXmlParagraphWrapper? previous)
        {
            StringBuilder sb = new();
            if (paragraphWrappers == null) { return sb.ToString(); }
            
            
            while (paragraphWrappers.Count > 0)
            {
                var current = paragraphWrappers.Dequeue();
                paragraphWrappers.TryPeek(out var next);

                if (current?.R == null) { return sb.ToString(); };
                if (current.R.Count == 0) { return sb.ToString(); };
                
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
