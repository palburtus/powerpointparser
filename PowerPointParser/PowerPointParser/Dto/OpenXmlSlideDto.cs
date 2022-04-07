using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Aaks.PowerPointParser.Dto
{


	[XmlRoot(ElementName = "schemeClr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class SchemeClr
	{
		[XmlAttribute(AttributeName = "val")]
		public string Val { get; set; }
		[XmlElement(ElementName = "shade", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public Shade Shade { get; set; }
	}

	[XmlRoot(ElementName = "solidFill", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class SolidFill
	{
		[XmlElement(ElementName = "schemeClr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public SchemeClr SchemeClr { get; set; }
		[XmlElement(ElementName = "srgbClr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public SrgbClr SrgbClr { get; set; }
	}

	[XmlRoot(ElementName = "bgPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class BgPr
	{
		[XmlElement(ElementName = "solidFill", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public SolidFill SolidFill { get; set; }
		[XmlElement(ElementName = "effectLst", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public string EffectLst { get; set; }
	}

	[XmlRoot(ElementName = "bg", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class Bg
	{
		[XmlElement(ElementName = "bgPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public BgPr BgPr { get; set; }
	}

	[XmlRoot(ElementName = "cNvPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class CNvPr
	{
		[XmlAttribute(AttributeName = "id")]
		public string Id { get; set; }
		[XmlAttribute(AttributeName = "name")]
		public string Name { get; set; }
		[XmlElement(ElementName = "extLst", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public ExtLst ExtLst { get; set; }
		[XmlAttribute(AttributeName = "descr")]
		public string Descr { get; set; }
	}

	[XmlRoot(ElementName = "nvGrpSpPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class NvGrpSpPr
	{
		[XmlElement(ElementName = "cNvPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public CNvPr CNvPr { get; set; }
		[XmlElement(ElementName = "cNvGrpSpPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public string CNvGrpSpPr { get; set; }
		[XmlElement(ElementName = "nvPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public string NvPr { get; set; }
	}

	[XmlRoot(ElementName = "off", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class Off
	{
		[XmlAttribute(AttributeName = "x")]
		public string X { get; set; }
		[XmlAttribute(AttributeName = "y")]
		public string Y { get; set; }
	}

	[XmlRoot(ElementName = "ext", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class Ext
	{
		[XmlAttribute(AttributeName = "cx")]
		public string Cx { get; set; }
		[XmlAttribute(AttributeName = "cy")]
		public string Cy { get; set; }
		[XmlElement(ElementName = "creationId", Namespace = "http://schemas.microsoft.com/office/drawing/2014/main")]
		public CreationId CreationId { get; set; }
		[XmlAttribute(AttributeName = "uri")]
		public string Uri { get; set; }
		[XmlElement(ElementName = "decorative", Namespace = "http://schemas.microsoft.com/office/drawing/2017/decorative")]
		public Decorative Decorative { get; set; }
		[XmlElement(ElementName = "useLocalDpi", Namespace = "http://schemas.microsoft.com/office/drawing/2010/main")]
		public UseLocalDpi UseLocalDpi { get; set; }
	}

	[XmlRoot(ElementName = "chOff", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class ChOff
	{
		[XmlAttribute(AttributeName = "x")]
		public string X { get; set; }
		[XmlAttribute(AttributeName = "y")]
		public string Y { get; set; }
	}

	[XmlRoot(ElementName = "chExt", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class ChExt
	{
		[XmlAttribute(AttributeName = "cx")]
		public string Cx { get; set; }
		[XmlAttribute(AttributeName = "cy")]
		public string Cy { get; set; }
	}

	[XmlRoot(ElementName = "xfrm", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class Xfrm
	{
		[XmlElement(ElementName = "off", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public Off Off { get; set; }
		[XmlElement(ElementName = "ext", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public Ext Ext { get; set; }
		[XmlElement(ElementName = "chOff", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public ChOff ChOff { get; set; }
		[XmlElement(ElementName = "chExt", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public ChExt ChExt { get; set; }
	}

	[XmlRoot(ElementName = "grpSpPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class GrpSpPr
	{
		[XmlElement(ElementName = "xfrm", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public Xfrm Xfrm { get; set; }
	}

	[XmlRoot(ElementName = "creationId", Namespace = "http://schemas.microsoft.com/office/drawing/2014/main")]
	public class CreationId
	{
		[XmlAttribute(AttributeName = "a16", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string A16 { get; set; }
		[XmlAttribute(AttributeName = "id")]
		public string Id { get; set; }
	}

	[XmlRoot(ElementName = "decorative", Namespace = "http://schemas.microsoft.com/office/drawing/2017/decorative")]
	public class Decorative
	{
		[XmlAttribute(AttributeName = "adec", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string Adec { get; set; }
		[XmlAttribute(AttributeName = "val")]
		public string Val { get; set; }
	}

	[XmlRoot(ElementName = "extLst", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class ExtLst
	{
		[XmlElement(ElementName = "ext", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public List<Ext> Ext { get; set; }
	}

	[XmlRoot(ElementName = "spLocks", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class SpLocks
	{
		[XmlAttribute(AttributeName = "noGrp")]
		public string NoGrp { get; set; }
		[XmlAttribute(AttributeName = "noRot")]
		public string NoRot { get; set; }
		[XmlAttribute(AttributeName = "noChangeAspect")]
		public string NoChangeAspect { get; set; }
		[XmlAttribute(AttributeName = "noMove")]
		public string NoMove { get; set; }
		[XmlAttribute(AttributeName = "noResize")]
		public string NoResize { get; set; }
		[XmlAttribute(AttributeName = "noEditPoints")]
		public string NoEditPoints { get; set; }
		[XmlAttribute(AttributeName = "noAdjustHandles")]
		public string NoAdjustHandles { get; set; }
		[XmlAttribute(AttributeName = "noChangeArrowheads")]
		public string NoChangeArrowheads { get; set; }
		[XmlAttribute(AttributeName = "noChangeShapeType")]
		public string NoChangeShapeType { get; set; }
		[XmlAttribute(AttributeName = "noTextEdit")]
		public string NoTextEdit { get; set; }
	}

	[XmlRoot(ElementName = "cNvSpPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class CNvSpPr
	{
		[XmlElement(ElementName = "spLocks", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public SpLocks SpLocks { get; set; }
	}

	[XmlRoot(ElementName = "designElem", Namespace = "http://schemas.microsoft.com/office/powerpoint/2015/main")]
	public class DesignElem
	{
		[XmlAttribute(AttributeName = "p16", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string P16 { get; set; }
		[XmlAttribute(AttributeName = "val")]
		public string Val { get; set; }
	}

	[XmlRoot(ElementName = "ext", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class Ext2
	{
		[XmlElement(ElementName = "designElem", Namespace = "http://schemas.microsoft.com/office/powerpoint/2015/main")]
		public DesignElem DesignElem { get; set; }
		[XmlAttribute(AttributeName = "uri")]
		public string Uri { get; set; }
		[XmlElement(ElementName = "creationId", Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main")]
		public CreationId2 CreationId2 { get; set; }
	}

	[XmlRoot(ElementName = "extLst", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class ExtLst2
	{
		[XmlElement(ElementName = "ext", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public Ext2 Ext2 { get; set; }
	}

	[XmlRoot(ElementName = "nvPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class NvPr
	{
		[XmlElement(ElementName = "extLst", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public ExtLst2 ExtLst2 { get; set; }
		[XmlElement(ElementName = "ph", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public Ph Ph { get; set; }
	}

	[XmlRoot(ElementName = "nvSpPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class NvSpPr
	{
		[XmlElement(ElementName = "cNvPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public CNvPr CNvPr { get; set; }
		[XmlElement(ElementName = "cNvSpPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public CNvSpPr CNvSpPr { get; set; }
		[XmlElement(ElementName = "nvPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public NvPr NvPr { get; set; }
	}

	[XmlRoot(ElementName = "prstGeom", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class PrstGeom
	{
		[XmlElement(ElementName = "avLst", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public string AvLst { get; set; }
		[XmlAttribute(AttributeName = "prst")]
		public string Prst { get; set; }
	}

	[XmlRoot(ElementName = "ln", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class Ln
	{
		[XmlElement(ElementName = "noFill", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public string NoFill { get; set; }
	}

	[XmlRoot(ElementName = "spPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class SpPr
	{
		[XmlElement(ElementName = "xfrm", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public Xfrm Xfrm { get; set; }
		[XmlElement(ElementName = "prstGeom", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public PrstGeom PrstGeom { get; set; }
		[XmlElement(ElementName = "ln", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public Ln Ln { get; set; }
		[XmlElement(ElementName = "solidFill", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public SolidFill SolidFill { get; set; }
		[XmlElement(ElementName = "effectLst", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public string EffectLst { get; set; }
	}

	[XmlRoot(ElementName = "shade", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class Shade
	{
		[XmlAttribute(AttributeName = "val")]
		public string Val { get; set; }
	}

	[XmlRoot(ElementName = "lnRef", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class LnRef
	{
		[XmlElement(ElementName = "schemeClr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public SchemeClr SchemeClr { get; set; }
		[XmlAttribute(AttributeName = "idx")]
		public string Idx { get; set; }
	}

	[XmlRoot(ElementName = "fillRef", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class FillRef
	{
		[XmlElement(ElementName = "schemeClr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public SchemeClr SchemeClr { get; set; }
		[XmlAttribute(AttributeName = "idx")]
		public string Idx { get; set; }
	}

	[XmlRoot(ElementName = "effectRef", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class EffectRef
	{
		[XmlElement(ElementName = "schemeClr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public SchemeClr SchemeClr { get; set; }
		[XmlAttribute(AttributeName = "idx")]
		public string Idx { get; set; }
	}

	[XmlRoot(ElementName = "fontRef", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class FontRef
	{
		[XmlElement(ElementName = "schemeClr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public SchemeClr SchemeClr { get; set; }
		[XmlAttribute(AttributeName = "idx")]
		public string Idx { get; set; }
	}

	[XmlRoot(ElementName = "style", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class Style
	{
		[XmlElement(ElementName = "lnRef", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public LnRef LnRef { get; set; }
		[XmlElement(ElementName = "fillRef", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public FillRef FillRef { get; set; }
		[XmlElement(ElementName = "effectRef", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public EffectRef EffectRef { get; set; }
		[XmlElement(ElementName = "fontRef", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public FontRef FontRef { get; set; }
	}

	[XmlRoot(ElementName = "bodyPr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class BodyPr
	{
		[XmlAttribute(AttributeName = "rtlCol")]
		public string RtlCol { get; set; }
		[XmlAttribute(AttributeName = "anchor")]
		public string Anchor { get; set; }
		[XmlElement(ElementName = "normAutofit", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public string NormAutofit { get; set; }
	}


	[XmlRoot(ElementName = "endParaRPr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class EndParaRPr
	{
		[XmlAttribute(AttributeName = "lang")]
		public string Lang { get; set; }
		[XmlAttribute(AttributeName = "dirty")]
		public string Dirty { get; set; }
	}

	[XmlRoot(ElementName = "p", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class P
	{
		[XmlElement(ElementName = "pPr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public PPr PPr { get; set; }
		[XmlElement(ElementName = "endParaRPr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public EndParaRPr EndParaRPr { get; set; }
		[XmlElement(ElementName = "r", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public R R { get; set; }
	}

	[XmlRoot(ElementName = "txBody", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class TxBody
	{
		[XmlElement(ElementName = "bodyPr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public BodyPr BodyPr { get; set; }
		[XmlElement(ElementName = "lstStyle", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public string LstStyle { get; set; }
		[XmlElement(ElementName = "p", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public P P { get; set; }
	}

	[XmlRoot(ElementName = "sp", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class Sp
	{
		[XmlElement(ElementName = "nvSpPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public NvSpPr NvSpPr { get; set; }
		[XmlElement(ElementName = "spPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public SpPr SpPr { get; set; }
		[XmlElement(ElementName = "style", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public Style Style { get; set; }
		[XmlElement(ElementName = "txBody", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public TxBody TxBody { get; set; }
		[XmlAttribute(AttributeName = "useBgFill")]
		public string UseBgFill { get; set; }
	}

	[XmlRoot(ElementName = "ph", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class Ph
	{
		[XmlAttribute(AttributeName = "type")]
		public string Type { get; set; }
		[XmlAttribute(AttributeName = "idx")]
		public string Idx { get; set; }
	}

	[XmlRoot(ElementName = "srgbClr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class SrgbClr
	{
		[XmlAttribute(AttributeName = "val")]
		public string Val { get; set; }
	}

	[XmlRoot(ElementName = "picLocks", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class PicLocks
	{
		[XmlAttribute(AttributeName = "noChangeAspect")]
		public string NoChangeAspect { get; set; }
	}

	[XmlRoot(ElementName = "cNvPicPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class CNvPicPr
	{
		[XmlElement(ElementName = "picLocks", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public PicLocks PicLocks { get; set; }
	}

	[XmlRoot(ElementName = "nvPicPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class NvPicPr
	{
		[XmlElement(ElementName = "cNvPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public CNvPr CNvPr { get; set; }
		[XmlElement(ElementName = "cNvPicPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public CNvPicPr CNvPicPr { get; set; }
		[XmlElement(ElementName = "nvPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public string NvPr { get; set; }
	}

	[XmlRoot(ElementName = "useLocalDpi", Namespace = "http://schemas.microsoft.com/office/drawing/2010/main")]
	public class UseLocalDpi
	{
		[XmlAttribute(AttributeName = "a14", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string A14 { get; set; }
		[XmlAttribute(AttributeName = "val")]
		public string Val { get; set; }
	}

	[XmlRoot(ElementName = "blip", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
	public class Blip
	{
		[XmlElement(ElementName = "extLst", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public ExtLst ExtLst { get; set; }
		[XmlAttribute(AttributeName = "embed", Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
		public string Embed { get; set; }
	}

	[XmlRoot(ElementName = "blipFill", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class BlipFill
	{
		[XmlElement(ElementName = "blip", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public Blip Blip { get; set; }
		[XmlElement(ElementName = "srcRect", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public string SrcRect { get; set; }
		[XmlElement(ElementName = "stretch", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public string Stretch { get; set; }
		[XmlAttribute(AttributeName = "rotWithShape")]
		public string RotWithShape { get; set; }
	}

	[XmlRoot(ElementName = "pic", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class Pic
	{
		[XmlElement(ElementName = "nvPicPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public NvPicPr NvPicPr { get; set; }
		[XmlElement(ElementName = "blipFill", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public BlipFill BlipFill { get; set; }
		[XmlElement(ElementName = "spPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public SpPr SpPr { get; set; }
	}

	[XmlRoot(ElementName = "spTree", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class SpTree
	{
		[XmlElement(ElementName = "nvGrpSpPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public NvGrpSpPr NvGrpSpPr { get; set; }
		[XmlElement(ElementName = "grpSpPr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public GrpSpPr GrpSpPr { get; set; }
		[XmlElement(ElementName = "sp", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public List<Sp> Sp { get; set; }
		[XmlElement(ElementName = "pic", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public Pic Pic { get; set; }
	}

	[XmlRoot(ElementName = "creationId", Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main")]
	public class CreationId2
	{
		[XmlAttribute(AttributeName = "p14", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string P14 { get; set; }
		[XmlAttribute(AttributeName = "val")]
		public string Val { get; set; }
	}

	[XmlRoot(ElementName = "cSld", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class CSld
	{
		[XmlElement(ElementName = "bg", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public Bg Bg { get; set; }
		[XmlElement(ElementName = "spTree", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public SpTree SpTree { get; set; }
		[XmlElement(ElementName = "extLst", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public ExtLst2 ExtLst2 { get; set; }
	}

	[XmlRoot(ElementName = "clrMapOvr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class ClrMapOvr
	{
		[XmlElement(ElementName = "masterClrMapping", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
		public string MasterClrMapping { get; set; }
	}

	[XmlRoot(ElementName = "sld", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
	public class OpenXmlSlide
	{
		[XmlElement(ElementName = "cSld", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public CSld CSld { get; set; }
		[XmlElement(ElementName = "clrMapOvr", Namespace = "http://schemas.openxmlformats.org/presentationml/2006/main")]
		public ClrMapOvr ClrMapOvr { get; set; }
		[XmlAttribute(AttributeName = "a16", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string A16 { get; set; }
		[XmlAttribute(AttributeName = "adec", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string Adec { get; set; }
		[XmlAttribute(AttributeName = "p16", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string P16 { get; set; }
		[XmlAttribute(AttributeName = "a14", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string A14 { get; set; }
		[XmlAttribute(AttributeName = "p14", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string P14 { get; set; }
		[XmlAttribute(AttributeName = "a", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string A { get; set; }
		[XmlAttribute(AttributeName = "r", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string R { get; set; }
		[XmlAttribute(AttributeName = "p", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string P { get; set; }
	}

}
