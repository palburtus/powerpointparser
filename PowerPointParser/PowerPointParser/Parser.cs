﻿using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Logging;
using PowerPointParser.Dto;
using Slide = PowerPointParser.Model.Slide;

namespace PowerPointParser
{
    public class Parser : IParser
    {
        private const string XpathNotesToSp = @"/*[local-name() = 'notes']/*[local-name() = 'cSld']/*[local-name() = 'spTree']/*[local-name() = 'sp']";
        private readonly IHtmlConverter _htmlConverter;
        private ILogger _logger;

        public Parser(IHtmlConverter htmlConverter, ILogger logger)
        {
            _htmlConverter = htmlConverter;
            _logger = logger;
        }

        public IList<Slide> ParseSpeakerNotes(string path)
        {
            IList<Slide> slides = new List<Slide>();

            using PresentationDocument presentationDocument = PresentationDocument.Open(path, false);
            var presentationPart = presentationDocument.PresentationPart;

            if (presentationPart == null) return slides;

            var presentation = presentationPart.Presentation;

            if (presentation.SlideIdList == null) return slides;

            foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
            {
                Slide slide = new Slide();

                var note = GetnotesSlidePart(presentationPart, slideId);

                if (DoesSlideHaveSpeakerNotes(note))
                {
                    
                    var pNodesList = ParsePNodesList(note!);

                    if (pNodesList != null)
                    {
                        XmlSerializer xmlSerializer = new XmlSerializer(typeof(OpenXmlParagraphWrapper));
                        foreach (XmlNode node in pNodesList)
                        {

                            try
                            {
                                using StringReader stringReader = new StringReader(node.OuterXml);
                                OpenXmlParagraphWrapper? paragraphNode = (OpenXmlParagraphWrapper)xmlSerializer.Deserialize(stringReader)!;
                                
                                slide.SpeakerNotes = _htmlConverter.ConvertOpenXmlParagraphWrapperToHtml(paragraphNode);
                                
                            }
                            catch (InvalidOperationException ex)
                            {
                                _logger.Log(LogLevel.Error, ex, $"Slide Note Deserialization Failed");
                            }
                            catch (Exception ex)
                            {
                                if (ex != null) _logger.Log(LogLevel.Critical, ex, ex.Message);
                                _logger.Log(LogLevel.Critical, ex, "Unknown Exception Occurred");
                            }    
                        }
                    }
                }

                slides.Add(slide);
                
            }

            return slides;
        }

        private static NotesSlidePart? GetnotesSlidePart(PresentationPart presentationPart, SlideId slideId)
        {
            SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

            var note = slidePart.NotesSlidePart;
            return note;
        }

        private static bool DoesSlideHaveSpeakerNotes(NotesSlidePart? note)
        {
            if(note == null) return false;

            return !string.IsNullOrEmpty(note.NotesSlide.InnerText);
        }

        private static XmlNodeList? ParsePNodesList(NotesSlidePart note)
        {
            var xml = note.NotesSlide.OuterXml;
            XmlDocument xmlDocument = new();
            xmlDocument.LoadXml(xml);

            var xmlNamespaceManager = new XmlNamespaceManager(xmlDocument.NameTable);
            xmlNamespaceManager.AddNamespace(@"p", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var spNodesList = xmlDocument.SelectNodes(XpathNotesToSp, xmlNamespaceManager);
            
            if(spNodesList == null) return null;
            
            var bodyNode = spNodesList[1];

            if(bodyNode == null) return null;

            XmlDocument bodyNodeXmlDocument = new();
            
            byte[] bytes = Encoding.UTF8.GetBytes(bodyNode.OuterXml);
            MemoryStream stream = new MemoryStream(bytes);
            bodyNodeXmlDocument.Load(stream);

            var pNodesList =
                bodyNodeXmlDocument.SelectNodes("/*[local-name() = 'sp']/*[local-name() = 'txBody']/*[local-name() = 'p']",
                    xmlNamespaceManager);
            return pNodesList;
        }
    }
}