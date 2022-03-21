using System.Collections.Generic;
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

        public string? ConvertOpenXmlParagraphWrapperToHtml(Queue<OpenXmlParagraphWrapper?>? paragraphWrappers)
        {
            return ConvertHtmlParagraphWrapperToHtml(paragraphWrappers, null);
        }

        private string? ConvertHtmlParagraphWrapperToHtml(Queue<OpenXmlParagraphWrapper?>? paragraphWrappers, OpenXmlParagraphWrapper? previous)
        {
            if (paragraphWrappers == null) { return null; }
            
            StringBuilder sb = new();
            while (paragraphWrappers.Count > 0)
            {
                var current = paragraphWrappers.Dequeue();
                paragraphWrappers.TryPeek(out var next);

                if (current?.R == null) return null;
                if (current.R.Count == 0) return null;
                
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
