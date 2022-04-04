using System.Collections.Generic;
using Aaks.PowerPointParser.Dto;

namespace PowerPointParserTests.Html
{
    public  abstract class BaseOpenXmlParserTests
    {
        protected static OpenXmlTextWrapper BuildOlTextWrapper(string text, int nestingLevel = 0)
        {

            int marL = nestingLevel switch
            {
                2 => 1143000,
                1 => 685800,
                _ => 228600
            };

            return new OpenXmlTextWrapper
            {
                PPr = new PPr
                {
                    Lvl = nestingLevel,
                    BuAutoNum = new BuAutoNum { Type = "arabicPeriod" },
                    MarL = marL,
                    Indent = -228600,
                    BuFont = new BuFont { Typeface = "+mj-lt" }
                },
                R = new List<R> { BuildRItem(text) }
            };
        }

        protected static PPr BuildSpecialUlCharacterDefaultPPr(string charType)
        {
            return new PPr
            {
                BuChar = new BuChar { Char = charType },
                SpcAft = BuildDefaultSpcAft(),
                SpcBef = BuildDefaultSpcBef(),
                MarL = 171450,
                Indent = -171450,
                Algn = "l",
                BuFont = new BuFont { Typeface = "Wingdings" }
            };
        }

        protected static PPr BuildSpecialOlDefaultPPr(string type)
        {
            return new PPr
            {
                MarL = 228600,
                Indent = -228600,
                Algn = "l",
                BuFont = new BuFont { Typeface = "+mj-lt" },
                BuAutoNum = new BuAutoNum { Type = type },
                SpcBef = new SpcBef
                {
                    SpcPts = new SpcPts { Val = 0 }
                },
                SpcAft = new SpcAft
                {
                    SpcPts = new SpcPts { Val = 0 }
                }

            };
        }

        protected static R BuildRItem(string text)
        {
            return new()
            {
                T = text,
                RPr = new RPr
                {
                    B = 0,
                    Dirty = 0,
                    Lang = "en-US"
                }
            };
        }

        protected static OpenXmlTextWrapper BuildParagraphTextWrapper(string text)
        {
            return new OpenXmlTextWrapper
            {
                R = new List<R> {new()
                    {
                        T = text,
                        RPr = new RPr
                        {
                            B = 0,
                            Dirty = 0,
                            Lang = "en-US"
                        }
                    }
                }
            };
        }

        protected static PPr BuildDefaultParagraphPpr()
        {
            return new PPr
            {
                Algn = "l",
                Indent = 0,
                Lvl = 0,
                MarL = 0,
                Rtl = 0
            };
        }

        protected static SpcAft BuildDefaultSpcAft()
        {
            return new SpcAft
            {
                SpcPts = new SpcPts
                {
                    Val = 0
                }
            };
        }

        protected static SpcBef BuildDefaultSpcBef()
        {
            return new SpcBef
            {
                SpcPts = new SpcPts
                {
                    Val = 0
                }
            };
        }

        protected static OpenXmlTextWrapper BuildUlTextWrapper(string text, int nestingLevel = 0)
        {

            int marL = nestingLevel switch
            {
                2 => 1085850,
                1 => 628650,
                _ => 171450
            };

            return new OpenXmlTextWrapper
            {
                PPr = new PPr
                {
                    Lvl = nestingLevel,
                    BuChar = new BuChar { Char = "•" },
                    BuFont = new BuFont { Typeface = "Arial" },
                    MarL = marL,
                    Indent = -171450
                },
                R = new List<R> { new()
                {
                    T = text,
                    RPr = new RPr
                    {
                        B = 0,
                        Dirty = 0,
                        Lang = "en-US"
                    }
                }}
            };
        }
    }
}
