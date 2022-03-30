using System.Collections.Generic;
using System.Xml.Serialization;

namespace Aaks.PowerPointParser.Dto
{
	[XmlRoot(ElementName = "spcPts")]
    public class SpcPts
    {

        [XmlAttribute(AttributeName = "val")]
        public int Val { get; set; }
    }

    [XmlRoot(ElementName = "spcBef")]
    public class SpcBef
    {

        [XmlElement(ElementName = "spcPts")]
        public SpcPts? SpcPts { get; set; }
    }

    [XmlRoot(ElementName = "spcAft")]
    public class SpcAft
    {

        [XmlElement(ElementName = "spcPts")]
        public SpcPts? SpcPts { get; set; }
    }

    [XmlRoot(ElementName = "pPr")]
    public class PPr
    {

        [XmlElement(ElementName = "spcBef")]
        public SpcBef? SpcBef { get; set; }

        [XmlElement(ElementName = "spcAft")]
        public SpcAft? SpcAft { get; set; }

        [XmlElement(ElementName = "buNone")]
        public object? BuNone { get; set; }

        [XmlElement(ElementName = "buFont")]
        public BuFont? BuFont { get; set; }

        [XmlElement(ElementName = "buChar")]
        public BuChar? BuChar { get; set; }

        [XmlElement(ElementName = "buAutoNum")]
        public BuAutoNum? BuAutoNum { get; set; }

        [XmlAttribute(AttributeName = "marL")]
        public int MarL { get; set; }

        [XmlAttribute(AttributeName = "lvl")]
        public int Lvl { get; set; }

        [XmlAttribute(AttributeName = "indent")]
        public int Indent { get; set; }

        [XmlAttribute(AttributeName = "algn")]
        public string? Algn { get; set; }

        [XmlAttribute(AttributeName = "rtl")]
        public int Rtl { get; set; }
    }

    [XmlRoot(ElementName = "buFont")]
    public class BuFont
    {

        [XmlAttribute(AttributeName = "typeface")]
        public string? Typeface { get; set; }
    }

    [XmlRoot(ElementName = "buChar")]
    public class BuChar
    {

        [XmlAttribute(AttributeName = "char")]
        public string? Char { get; set; }
    }

    [XmlRoot(ElementName = "buAutoNum")]
    public class BuAutoNum
    {

        [XmlAttribute(AttributeName = "type")]
        public string? Type { get; set; }
    }

    [XmlRoot(ElementName = "rPr")]
    public class RPr
    {

        [XmlAttribute(AttributeName = "lang")]
        public string? Lang { get; set; }

        [XmlAttribute(AttributeName = "b")]
        public int B { get; set; }

        [XmlAttribute(AttributeName = "i")]
        public int I { get; set; }

        [XmlAttribute(AttributeName = "u")]
        public string? U { get; set; }

        [XmlAttribute(AttributeName = "strike")]
        public string? Strike { get; set; }

        [XmlAttribute(AttributeName = "dirty")]
        public int Dirty { get; set; }
    }

    [XmlRoot(ElementName = "r")]
    public class R
    {

        [XmlElement(ElementName = "rPr")]
        public RPr? RPr { get; set; }

        [XmlElement(ElementName = "t")]
        public string? T { get; set; }
    }

    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", ElementName = "p")]
    public class OpenXmlTextWrapper
    {

        [XmlElement(ElementName = "pPr")]
        public PPr? PPr { get; set; }

        [XmlElement(ElementName = "r")]
        public List<R>? R { get; set; }

        [XmlAttribute(AttributeName = "a")]
        public string? A { get; set; }

        [XmlText]
        public string? Text { get; set; }
    }
}
