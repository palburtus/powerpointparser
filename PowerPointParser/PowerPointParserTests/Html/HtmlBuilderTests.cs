using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Aaks.PowerPointParser.Dto;
using FluentAssertions;
using PowerPointParserTests.Html;

// ReSharper disable once CheckNamespace - Test namespaces should match production
namespace Aaks.PowerPointParser.Html.Tests
{
    [TestClass]
    public class HtmlBuilderTests : BaseHtmlTests
    {
        private static IHtmlBuilder _htmlConverter = null!;

        [ClassInitialize]
        public static void ClassSetup(TestContext context)
        {
            var innerHtmlBuilder = new InnerHtmlBuilder();
            _htmlConverter = new HtmlBuilder(new HtmlListBuilder(innerHtmlBuilder), innerHtmlBuilder);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_WrapperListNull_ReturnsNull()
        {
            Queue<OpenXmlTextWrapper?>? wrapperList = null;

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(wrapperList);

            actual.Should().BeEmpty();
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_WrapperNull_ReturnsNull()
        {
            OpenXmlTextWrapper? wrapper = null;
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            actual.Should().BeEmpty();
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_RNull_ReturnsNull()
        {
            OpenXmlTextWrapper wrapper = new();
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            actual.Should().BeEmpty();
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_REmptyNull_ReturnsNull()
        {
            OpenXmlTextWrapper wrapper = new()
            {
                R = new List<R>()
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            actual.Should().BeEmpty();
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ParagraphTag_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p>hello world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_BoldTag_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", new RPr {B = 1}));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong>hello world</strong></p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ItalicTag_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", new RPr { I = 1 }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><i>hello world</i></p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_UnderlineTag_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", new RPr { U = "sng" }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><u>hello world</u></p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_StrikeThroughTag_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", new RPr { Strike = "sngStrike" }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><del>hello world</del></p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_StrongItalicTag_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", new RPr { I = 1 , B = 1}));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong><i>hello world</i></strong></p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_StrongUnderlinedItalicTag_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", new RPr { I = 1, B = 1, U = "sng" }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong><u><i>hello world</i></u></strong></p>", actual);
        }


        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_BoldAndNonBoldMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {B = 1}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr {BuNone = new object()},
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong>hello:</strong> world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ItalicAndNonItalicMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {I = 1}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><i>hello:</i> world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_UnderlinedAndNonUnderlinedMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {U = "sng"}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><u>hello:</u> world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_StrikethroughAndNonStrikethrough_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {Strike = "sngStrike"}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><del>hello:</del> world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ItalicAndBoldMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {I = 1}),
                BuildR(" world", new RPr {B = 1})
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><i>hello:</i><strong> world</strong></p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_BoldItalicMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {I = 1, B = 1}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong><i>hello:</i></strong> world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_BoldItalicUnderLineMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {I = 1, B = 1, U = "sng"}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong><u><i>hello:</i></u></strong> world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_BoldItalicUnderLineStrikeThroughMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {I = 1, B = 1, U = "sng", Strike = "sngStrike"}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><strong><u><i><del>hello:</del></i></u></strong> world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlignParagraphCenter_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", ppr: new PPr{Algn = "ctr"}));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p style=\"text-align: center;\">hello world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlignParagraphRight_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", ppr: new PPr { Algn = "r" }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p style=\"text-align: right;\">hello world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlignParagraphJustify_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello world", ppr: new PPr { Algn = "just" }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p style=\"text-align: justify;\">hello world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlignOrderedListItemCenter_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world", align: "ctr"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li style=\"text-align: center;\">hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlignOrderedListItemRight_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world", align: "r"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li style=\"text-align: right;\">hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlignOrderedListItemJustify_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world", align: "just"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li style=\"text-align: justify;\">hello world</li></ol>", actual);
        }


        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlignUnOrderedListItemCenter_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", align: "ctr"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li style=\"text-align: center;\">hello world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlignUnOrderedListItemRight_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", align: "r"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li style=\"text-align: right;\">hello world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlignUnOrderedListItemJustify_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", align: "just"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li style=\"text-align: justify;\">hello world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_BoldItalicUnderLineStrikeAlignRightMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {I = 1, B = 1, U = "sng", Strike = "sngStrike"}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object(), Algn = "r"},
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p style=\"text-align: right;\"><strong><u><i><del>hello:</del></i></u></strong> world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_UlBoldItalicUnderLineStrikeAlignCenterMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {I = 1, B = 1, U = "sng", Strike = "sngStrike"}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object(), Algn = "ctr", BuChar = new BuChar { Char = "•" } },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li style=\"text-align: center;\"><strong><u><i><del>hello:</del></i></u></strong> world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OlBoldItalicUnderLineStrikeAlignJustifyMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello:", new RPr {I = 1, B = 1, U = "sng", Strike = "sngStrike"}),
                BuildR(" world")
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object(), Algn = "just", BuAutoNum = new BuAutoNum { Type = "arabicPeriod" } },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li style=\"text-align: justify;\"><strong><u><i><del>hello:</del></i></u></strong> world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ItalicUnderlinedAndBoldNormalMix_ReturnsString()
        {
            var rs = new List<R>
            {
                BuildR("hello", new RPr {I = 1}),
                BuildR(" and"),
                BuildR(" world", new RPr {B = 1}),
                BuildR(" or"),
                BuildR(" one", new RPr {U = "sng"}),
                BuildR(" not", new RPr {Strike = "sngStrike"})
            };

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr { BuNone = new object() },
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p><i>hello</i> and<strong> world</strong> or<u> one</u><del> not</del></p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ConsecutiveParagraph_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello"));

            Queue<OpenXmlTextWrapper?> queue2 = new();
            queue2.Enqueue(BuildParagraphLine("world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);
            actual += _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue2);

            Assert.AreEqual("<p>hello</p><p>world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TwoConsecutiveEmptyParagraphLines_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello"));
            queue.Enqueue(new()
            {
                R = new List<R>()
            });
            queue.Enqueue(new()
            {
                R = new List<R>()
            });
            queue.Enqueue(BuildParagraphLine("world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);
            

            Assert.AreEqual("<p>hello</p><p>world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ThreeConsecutiveEmptyParagraphLines_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("hello"));
            queue.Enqueue(new()
            {
                R = new List<R>()
            });
            queue.Enqueue(new()
            {
                R = new List<R>()
            });
            queue.Enqueue(new()
            {
                R = new List<R>()
            });
            queue.Enqueue(BuildParagraphLine("world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);


            Assert.AreEqual("<p>hello</p><p>world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OneLineInTwoRRecords_ReturnsString()
        {
            var rs = new List<R>();
            var rOne = new R {T = "hello "};
            var rTwo = new R {T = "world"};
            rs.Add(rOne);
            rs.Add(rTwo);

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr {BuNone = new object()},
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p>hello world</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ParagraphBeforeUnorderedList_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("Paragraph Before"));
            queue.Enqueue(BuildUnorderedListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("two"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p>Paragraph Before</p><ul><li>one</li><li>two</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ParagraphAfterUnorderedList_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("two"));
            queue.Enqueue(BuildParagraphLine("Paragraph After"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>one</li><li>two</li></ul><p>Paragraph After</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ParagraphBeforeAndAfterUnorderedList_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildParagraphLine("Paragraph Before"));
            queue.Enqueue(BuildUnorderedListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("two"));
            queue.Enqueue(BuildParagraphLine("Paragraph After"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<p>Paragraph Before</p><ul><li>one</li><li>two</li></ul><p>Paragraph After</p>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_StrongUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", new RPr {B = 1}));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li><strong>hello world</strong></li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_HollowRoundBulletsUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", specialChar: "o"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_FilledSquareBulletsUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", specialChar: "§"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_HollowSquareBulletsUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", specialChar: "q"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_StarBulletsUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", specialChar: "v"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ArrowBulletsUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", specialChar: "Ø"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_CheckmarkBulletsUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", specialChar: "ü"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ArabicParenRightOrderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world", specialFont: "arabicParenR"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_CapitalRomanNumeralsPeriodOrderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world", specialFont: "romanUcPeriod"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_CapitalLettersPeriodOrderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world", specialFont: "alphaUcPeriod"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_LowercaseAlphaRightParenOrderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world", specialFont: "alphaLcParenR"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_LowercaseAlphaPeriodOrderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world", specialFont: "alphaLcPeriod"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_LowercaseRomanNumeralsPeriodOrderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world", specialFont: "romanLcPeriod"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_ItalicUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", new RPr { I = 1 }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li><i>hello world</i></li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_UnderlinedUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", new RPr { U = "sng" }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li><u>hello world</u></li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_StrikeThroughUnorderedListItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world", new RPr { Strike = "sngStrike" }));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li><del>hello world</del></li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_UnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><li>goodbye world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedListItem_ReturnsString()
        {
            var wrapper = BuildOrderListItem("hello world");

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_EmbeddedOrderedListItem_ReturnsString()
        {
            var rs = new List<R>();
            var rOne = BuildR("hello");
            var rTwo = BuildR(" world");
            var rThree = BuildR(" ");
            var rFour = BuildR("test");

            rs.Add(rOne);
            rs.Add(rTwo);
            rs.Add(rThree);
            rs.Add(rFour);

            OpenXmlTextWrapper wrapper = new()
            {
                PPr = new PPr {BuAutoNum = new BuAutoNum {Type = "arabicPeriod"}},
                R = rs
            };

            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(wrapper);

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world test</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("hello world"));
            queue.Enqueue(BuildOrderListItem("goodbye world"));
            queue.Enqueue(BuildOrderListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>hello world</li><li>goodbye world</li><li>test world</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><ul><li>nested item</li></ul><li>test world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TwiceNestedUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual(
                "<ul><li>hello world</li><ul><li>nested item</li><ul><li>nested double</li></ul></ul><li>test world</li></ul>",
                actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TwiceTwoNestedUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("nested double two", level: 2));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual(
                "<ul><li>hello world</li><ul><li>nested item</li><ul><li>nested double</li><li>nested double two</li></ul></ul><li>test world</li></ul>",
                actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedListTypeChangeAfterFirstItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("It supports"));
            queue.Enqueue(BuildUnorderedListItem("Un-ordered lists"));
            queue.Enqueue(BuildUnorderedListItem("Nested Lists", level: 1));
            queue.Enqueue(BuildOrderListItem("And", level: 1));
            queue.Enqueue(BuildOrderListItem("Ordered Lists", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>It supports</li><li>Un-ordered lists</li><ul><li>Nested Lists</li></ul><ol><li>And</li><li>Ordered Lists</li></ol></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedTwiceAfterNoNesting_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("It supports"));
            queue.Enqueue(BuildUnorderedListItem("Un-ordered lists"));
            queue.Enqueue(BuildUnorderedListItem("Nested Lists", level: 2));
            queue.Enqueue(BuildOrderListItem("And", level: 2));
            queue.Enqueue(BuildOrderListItem("Ordered Lists", level: 2 ));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>It supports</li><li>Un-ordered lists</li><ul><ul><li>Nested Lists</li></ul><ol><li>And</li><li>Ordered Lists</li></ol></ul></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedFiveTimesAfterNoNesting_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("It supports"));
            queue.Enqueue(BuildUnorderedListItem("Un-ordered lists"));
            queue.Enqueue(BuildUnorderedListItem("Nested Lists", level: 5));
            queue.Enqueue(BuildOrderListItem("And", level: 5));
            queue.Enqueue(BuildOrderListItem("Ordered Lists", level: 2));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>It supports</li><li>Un-ordered lists</li><ul><ul><ul><ul><ul><li>Nested Lists</li><li>And</li></ul></ul></ul><li>Ordered Lists</li></ul></ul></ul>", actual);
        }


        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TwiceNestedFollowedBySingleNested_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("nested double two", level: 2));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual(
                "<ul><li>hello world</li><ul><li>nested item</li><ul><li>nested double</li><li>nested double two</li></ul><li>test world</li></ul></ul>",
                actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlternateEveryItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li></ol><ul><li>one one</li></ul><ol><li>one one one</li></ol><ul><li>one one one one</li></ul>", actual);

        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlternateEveryNestFourItem_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li></ol><ul><li>one one</li></ul><ol><li>one one one</li></ol><ul><li>one one one one</li><ol><li>two</li></ol><ul><li>two two</li></ul><ol><li>two two two</li></ol><ul><li>two two two two</li></ul></ul>", actual);

        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlternateEveryFourNestItemFinalNestZero_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("three"));
            queue.Enqueue(BuildUnorderedListItem("three three"));
            queue.Enqueue(BuildOrderListItem("three three three"));
            queue.Enqueue(BuildUnorderedListItem("three three three three"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li></ol><ul><li>one one</li></ul><ol><li>one one one</li></ol><ul><li>one one one one</li><ol><li>two</li></ol><ul><li>two two</li></ul><ol><li>two two two</li></ol><ul><li>two two two two</li></ul></ul><ol><li>three</li></ol><ul><li>three three</li></ul><ol><li>three three three</li></ol><ul><li>three three three three</li></ul>", actual);

        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AlternateEveryThreeNestItemFinalNestZero_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("three"));
            queue.Enqueue(BuildUnorderedListItem("three three"));
            queue.Enqueue(BuildOrderListItem("three three three"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li></ol><ul><li>one one</li><li>one one one</li><ol><li>two</li></ol><ul><li>two two</li><li>two two two two</li></ul></ul><ol><li>three</li></ol><ul><li>three three</li></ul><ol><li>three three three</li></ol>", actual);

        }



        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TripleNested_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildOrderListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("nested double two", level: 2));
            queue.Enqueue(BuildOrderListItem("three nested one", level: 3));
            queue.Enqueue(BuildOrderListItem("three nested two", level: 3));
            queue.Enqueue(BuildOrderListItem("test world", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><ol><li>nested item</li><ul><li>nested double</li><li>nested double two</li><ol><li>three nested one</li><li>three nested two</li></ol></ul><li>test world</li></ol></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedNestUnorderedNestOrdered_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item", level: 1));
            queue.Enqueue(BuildUnorderedListItem("nested double", level: 2));
            queue.Enqueue(BuildUnorderedListItem("nested double two", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three nested one", level: 3));
            queue.Enqueue(BuildUnorderedListItem("three nested two", level: 3));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><ul><li>nested item</li><ul><li>nested double</li><li>nested double two</li><ul><li>three nested one</li><li>three nested two</li></ul></ul><li>test world</li></ul></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedLastItemUnOrderListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("nested item"));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><li>nested item</li><ul><li>test world</li></ul></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedUnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("two"));
            queue.Enqueue(BuildUnorderedListItem("hello world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("goodbye world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three"));


            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>one</li><li>two</li><ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul><li>three</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_UnorderedFollowedByOrderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));
            queue.Enqueue(BuildUnorderedListItem("test world"));
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildOrderListItem("three"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul><ol><li>one</li><li>two</li><li>three</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedFollowedByUnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildOrderListItem("three"));
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));
            queue.Enqueue(BuildUnorderedListItem("test world"));
            
            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><li>three</li></ol><ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedOrderedFollowedByUnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildOrderListItem("three", level: 1));
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><ol><li>three</li></ol></ol><ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_OrderedFollowedByNestedUnorderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildOrderListItem("three"));
            queue.Enqueue(BuildUnorderedListItem("hello world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("goodbye world"));
            queue.Enqueue(BuildUnorderedListItem("test world"));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><li>three</li></ol><ul><ul><li>hello world</li></ul><li>goodbye world</li><li>test world</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedTwoUnorderedInsideOrderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildUnorderedListItem("hello world"));
            queue.Enqueue(BuildUnorderedListItem("goodbye world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));
            queue.Enqueue(BuildOrderListItem("three"));


            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><li>hello world</li><ul><li>goodbye world</li><li>test world</li></ul><li>three</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_AllNestedTwoUnorderedInsideOrderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one", level: 1));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("hello world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("goodbye world", level: 2));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 2));
            queue.Enqueue(BuildOrderListItem("three", level: 1));


            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><ol><li>one</li><li>two</li><li>hello world</li><ul><li>goodbye world</li><li>test world</li></ul><li>three</li></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_NestedUnorderedInsideOrderedListItems_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two"));
            queue.Enqueue(BuildUnorderedListItem("hello world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("goodbye world", level: 1));
            queue.Enqueue(BuildUnorderedListItem("test world", level: 1));
            queue.Enqueue(BuildOrderListItem("three"));
            

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>two</li><ul><li>hello world</li><li>goodbye world</li><li>test world</li></ul><li>three</li></ol>", actual);
        }


        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingAlternatingOrderEveryTwo_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><ol><li>two</li><ul><li>three</li><ul><li>four</li><ol><li>five</li><ol><li>six</li><ul><li>seven</li><ul><li>eight</li><ol><li>nine</li><ol><li>ten</li></ol></ol></ul></ul></ol></ol></ul></ul></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingAlternatingTwoOrderEveryTwo_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>one one</li><ol><li>two</li><li>two two</li><ul><li>three</li><li>three three</li>" +
                            "<ul><li>four</li><li>four four</li><ol><li>five</li><li>five five</li>" +
                            "<ol><li>six</li><li>six six</li><ul><li>seven</li><li>seven seven</li>" +
                            "<ul><li>eight</li><li>eight eight</li><ol><li>nine</li><li>nine nine</li>" +
                            "<ol><li>ten</li><li>ten ten</li></ol></ol></ul></ul></ol></ol></ul></ul></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingAlternatingThreeOrderEveryTwo_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>one one</li><li>one one one</li><ol><li>two</li><li>two two</li><li>two two two</li><ul><li>three</li><li>three three</li><li>three three three</li><ul><li>four</li><li>four four</li><li>four four four</li><ol><li>five</li><li>five five</li><li>five five five</li><ol><li>six</li><li>six six</li><li>six six six six</li><ul><li>seven</li><li>seven seven</li><li>seven seven seven</li><ul><li>eight</li><li>eight eight</li><li>eight eight eight</li><ol><li>nine</li><li>nine nine</li><li>nine nine nine</li><ol><li>ten</li><li>ten ten</li><li>ten ten ten</li></ol></ol></ul></ul></ol></ol></ul></ul></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingAlternatingFourOrderEveryTwo_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildOrderListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li><ol><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li><ul><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li><ul><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li><ol><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li><ol><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li><ul><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li><ul><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li><ol><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li><ol><li>ten</li><li>ten ten</li><li>ten ten ten</li><li>ten ten ten ten</li></ol></ol></ul></ul></ol></ol></ul></ul></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingFourUnordered_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));
            queue.Enqueue(BuildUnorderedListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li><ul><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li><ul><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li><ul><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li><ul><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li><ul><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li><ul><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li><ul><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li><ul><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li><ul><li>ten</li><li>ten ten</li><li>ten ten ten</li><li>ten ten ten ten</li></ul></ul></ul></ul></ul></ul></ul></ul></ul></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingFourUnorderedFinalNestZero_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildUnorderedListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));
            queue.Enqueue(BuildUnorderedListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten ten ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));


            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ul><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li><ul><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li><ul><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li><ul><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li><ul><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li><ul><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li><ul><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li><ul><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li><ul><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li><ul><li>ten</li><li>ten ten</li><li>ten ten ten</li><li>ten ten ten ten</li></ul><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li></ul><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li></ul><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li></ul><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li></ul><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li></ul><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li></ul><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li></ul><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li></ul><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li></ul>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingFourOrdered_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildOrderListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten ten", level: 9));

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li><ol><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li><ol><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li><ol><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li><ol><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li><ol><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li><ol><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li><ol><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li><ol><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li><ol><li>ten</li><li>ten ten</li><li>ten ten ten</li><li>ten ten ten ten</li></ol></ol></ol></ol></ol></ol></ol></ol></ol></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingFourOrderedFinalNestZero_ReturnsString()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildOrderListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildOrderListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildOrderListItem("one one one one"));


            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li><ol><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li><ol><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li><ol><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li><ol><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li><ol><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li><ol><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li><ol><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li><ol><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li><ol><li>ten</li><li>ten ten</li><li>ten ten ten</li><li>ten ten ten ten</li></ol><li>nine</li><li>nine nine</li><li>nine nine nine</li><li>nine nine nine nine</li></ol><li>eight</li><li>eight eight</li><li>eight eight eight</li><li>eight eight eight eight</li></ol><li>seven</li><li>seven seven</li><li>seven seven seven</li><li>seven seven seven seven</li></ol><li>six</li><li>six six</li><li>six six six six</li><li>six six six six six</li></ol><li>five</li><li>five five</li><li>five five five</li><li>five five five five</li></ol><li>four</li><li>four four</li><li>four four four</li><li>four four four four</li></ol><li>three</li><li>three three</li><li>three three three</li><li>three three three three</li></ol><li>two</li><li>two two</li><li>two two two</li><li>two two two two</li></ol><li>one</li><li>one one</li><li>one one one</li><li>one one one one</li></ol>", actual);
        }

        [TestMethod]
        public void ConvertOpenXmlParagraphWrapperToHtml_TenDeepNestingFourAlternateEachOrdered_ReturnsString()
        {
            var queue = GetTenDeepNestingFourAlternateEachOrdered();

            var actual = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(queue);

            Assert.AreEqual("<ol><li>one</li></ol><ul><li>one one</li></ul><ol><li>one one one</li></ol><ul><li>one one one one</li><ol><li>two</li></ol><ul><li>two two</li></ul><ol><li>two two two</li></ol><ul><li>two two two two</li><ol><li>three</li></ol><ul><li>three three</li></ul><ol><li>three three three</li></ol><ul><li>three three three three</li><ol><li>four</li></ol><ul><li>four four</li></ul><ol><li>four four four</li></ol><ul><li>four four four four</li><ol><li>five</li></ol><ul><li>five five</li></ul><ol><li>five five five</li></ol><ul><li>five five five five</li><ol><li>six</li></ol><ul><li>six six</li></ul><ol><li>six six six six</li></ol><ul><li>six six six six six</li><ol><li>seven</li></ol><ul><li>seven seven</li></ul><ol><li>seven seven seven</li></ol><ul><li>seven seven seven seven</li><ol><li>eight</li></ol><ul><li>eight eight</li></ul><ol><li>eight eight eight</li></ol><ul><li>eight eight eight eight</li><ol><li>nine</li></ol><ul><li>nine nine</li></ul><ol><li>nine nine nine</li></ol><ul><li>nine nine nine nine</li><ol><li>ten</li></ol><ul><li>ten ten</li></ul><ol><li>ten ten ten</li></ol><ul><li>ten ten ten ten</li></ul></ul></ul></ul></ul></ul></ul></ul></ul></ul>", actual);
        }

        public static Queue<OpenXmlTextWrapper?> GetTenDeepNestingFourAlternateEachOrdered()
        {
            Queue<OpenXmlTextWrapper?> queue = new();
            queue.Enqueue(BuildOrderListItem("one"));
            queue.Enqueue(BuildUnorderedListItem("one one"));
            queue.Enqueue(BuildOrderListItem("one one one"));
            queue.Enqueue(BuildUnorderedListItem("one one one one"));
            queue.Enqueue(BuildOrderListItem("two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two", level: 1));
            queue.Enqueue(BuildOrderListItem("two two two", level: 1));
            queue.Enqueue(BuildUnorderedListItem("two two two two", level: 1));
            queue.Enqueue(BuildOrderListItem("three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three", level: 2));
            queue.Enqueue(BuildOrderListItem("three three three", level: 2));
            queue.Enqueue(BuildUnorderedListItem("three three three three", level: 2));
            queue.Enqueue(BuildOrderListItem("four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four", level: 3));
            queue.Enqueue(BuildOrderListItem("four four four", level: 3));
            queue.Enqueue(BuildUnorderedListItem("four four four four", level: 3));
            queue.Enqueue(BuildOrderListItem("five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five", level: 4));
            queue.Enqueue(BuildOrderListItem("five five five", level: 4));
            queue.Enqueue(BuildUnorderedListItem("five five five five", level: 4));
            queue.Enqueue(BuildOrderListItem("six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six", level: 5));
            queue.Enqueue(BuildOrderListItem("six six six six", level: 5));
            queue.Enqueue(BuildUnorderedListItem("six six six six six", level: 5));
            queue.Enqueue(BuildOrderListItem("seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("seven seven seven", level: 6));
            queue.Enqueue(BuildUnorderedListItem("seven seven seven seven", level: 6));
            queue.Enqueue(BuildOrderListItem("eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("eight eight eight", level: 7));
            queue.Enqueue(BuildUnorderedListItem("eight eight eight eight", level: 7));
            queue.Enqueue(BuildOrderListItem("nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("nine nine nine", level: 8));
            queue.Enqueue(BuildUnorderedListItem("nine nine nine nine", level: 8));
            queue.Enqueue(BuildOrderListItem("ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten", level: 9));
            queue.Enqueue(BuildOrderListItem("ten ten ten", level: 9));
            queue.Enqueue(BuildUnorderedListItem("ten ten ten ten", level: 9));
            return queue;
        }
    }
}