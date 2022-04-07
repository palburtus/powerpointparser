using System.Text;
using System.Xml;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Aaks.PowerPointParser.Dto;
using System.IO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.Extensions.Logging;

namespace Aaks.PowerPointParser.Parsers
{
    public class PowerPointParser : IPowerPointParser
    {
        private const string XpathNotesToSp = @"/*[local-name() = 'notes']/*[local-name() = 'cSld']/*[local-name() = 'spTree']/*[local-name() = 'sp']";
        private const string PNodesListXPath = @"/*[local-name() = 'sp']/*[local-name() = 'txBody']/*[local-name() = 'p']";
        
        private readonly ILogger<PowerPointParser>? _logger;

        public IDictionary<int, IList<OpenXmlSlide?>> ParseSlide(string path)
        {
            using var presentationDocument = PresentationDocument.Open(path, false);
            var slidesContentMap = ParseSlides(presentationDocument);
            return slidesContentMap!;
        }

        public PowerPointParser(ILogger<PowerPointParser>? logger = null)
        {
            _logger = logger;
        }

        public IDictionary<int, IList<OpenXmlLineItem?>> ParseSpeakerNotes(MemoryStream memoryStream)
        {
            using var presentationDocument = PresentationDocument.Open(memoryStream, false);
            var slidesContentMap = ParseSpeakerNotes(presentationDocument);
            return slidesContentMap;

        }

        public IDictionary<int, IList<OpenXmlLineItem?>> ParseSpeakerNotes(string path)
        {
            using var presentationDocument = PresentationDocument.Open(path, false);
            var slidesContentMap = ParseSpeakerNotes(presentationDocument);
            return slidesContentMap;
        }

        private IDictionary<int, IList<OpenXmlSlide?>> ParseSlides(PresentationDocument presentationDocument)
        {
            IDictionary<int, IList<OpenXmlSlide?>> slidesContentMap = new Dictionary<int, IList<OpenXmlSlide?>>();
            var presentationPart = presentationDocument.PresentationPart;
            
            var slideIds = GetSlideIds(presentationPart);

            int slideIndex = 1;

            foreach (var slideId in slideIds)
            {
                if (slideId.RelationshipId == null) return slidesContentMap;

                var openXmlPart = presentationPart!.GetPartById(slideId.RelationshipId!);
                SlidePart? slidePart = openXmlPart as SlidePart;
                Slide? slide = slidePart?.Slide;
                //var lineItems = slidePart?.Slide.Descendants<Paragraph>();

                var slides = new List<OpenXmlSlide>();

                if (slide != null)
                {
                    slides.Add(Deserialize<OpenXmlSlide>(slide.OuterXml, typeof(OpenXmlSlide))?? new OpenXmlSlide());
                }

                slideIndex++;
                slidesContentMap.Add(slideIndex, slides!);
            }

            return slidesContentMap;
        }

        private IDictionary<int, IList<OpenXmlLineItem?>> ParseSpeakerNotes(PresentationDocument presentationDocument)
        {
            var slidesContentMap = new Dictionary<int, IList<OpenXmlLineItem>>();
            var presentationPart = presentationDocument.PresentationPart;
            
            var slideIds = GetSlideIds(presentationPart);

            int slideIndex = 1;

            foreach (var slideId in slideIds)
            {
                var notesSlidePart = GetNotesSlidePart(presentationPart!, slideId);
                var openXmlSpeakerNotes = new List<OpenXmlLineItem>();

                if (DoesSlideHaveSpeakerNotes(notesSlidePart))
                {
                    var pNodesList = ParsePNodesList(notesSlidePart!.NotesSlide.OuterXml);

                    if (pNodesList != null)
                    {
                        foreach (XmlNode node in pNodesList)
                        {
                            openXmlSpeakerNotes.AddRange(DeserializeList<OpenXmlLineItem>(node.OuterXml, typeof(OpenXmlLineItem)));
                        }
                    }
                }

                slidesContentMap.Add(slideIndex, openXmlSpeakerNotes);
                slideIndex++;
            }

            return slidesContentMap!;
        }

        private T? Deserialize<T>(string xml, Type type)
        {
            T? t = default(T);

            try
            {
                var xmlSerializer = new XmlSerializer(type);
                using StringReader stringReader = new(xml);
                using XmlTextReader xmlReader = new(stringReader);
                t = (T)xmlSerializer.Deserialize(xmlReader)!;
            }
            catch (InvalidOperationException ex)
            {
                string message = $"{ex.Message} Slide Note Deserialization Failed";
                Console.WriteLine(message);
                _logger?.LogError(message);
            }
            catch (Exception ex)
            {
                string message = $"{ex.Message} Unknown Exception Occurred";
                Console.WriteLine(message);
                _logger?.LogError(ex, message);
            }

            return t;
        }

        private List<T> DeserializeList<T>(string xml, Type type)
        {
            List<T> openXmlObjects = new List<T>();
            
            try
            {
                var xmlSerializer = new XmlSerializer(type);
                using StringReader stringReader = new(xml);
                using XmlTextReader xmlReader = new(stringReader);
                var item = (T) xmlSerializer.Deserialize(xmlReader)!;
                openXmlObjects.Add(item);
            }
            catch (InvalidOperationException ex)
            {
                string message = $"{ex.Message} Slide Note Deserialization Failed";
                Console.WriteLine(message);
                _logger?.LogError(message);
            }
            catch (Exception ex)
            {
                string message = $"{ex.Message} Unknown Exception Occurred";
                Console.WriteLine(message);
                _logger?.LogError(ex, message);
            }

            return openXmlObjects;
        }

        private static NotesSlidePart? GetNotesSlidePart(OpenXmlPartContainer presentationPart, SlideId? slideId)
        {
            if (slideId == null) return null;
            if (slideId.RelationshipId == null) return null;
            
            var openXmlPart = presentationPart.GetPartById(slideId.RelationshipId!);
            
            SlidePart? slidePart = openXmlPart as SlidePart;

            return slidePart?.NotesSlidePart;
        }

        private static IEnumerable<SlideId> GetSlideIds(PresentationPart? presentationPart)
        {
            if (presentationPart == null) return new List<SlideId>();

            var presentation = presentationPart.Presentation;

            if (presentation.SlideIdList == null) return new List<SlideId>();

            var slideIds = presentation.SlideIdList.Elements<SlideId>();
            return slideIds;
        }

        private static bool DoesSlideHaveSpeakerNotes(NotesSlidePart? note)
        {
            if(note == null) return false;

            return !string.IsNullOrEmpty(note.NotesSlide.InnerText);
        }

        private static XmlNodeList? ParsePNodesList(string xml)
        {
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
            using var stream = new MemoryStream(bytes);
            bodyNodeXmlDocument.Load(stream);

            return bodyNodeXmlDocument.SelectNodes(PNodesListXPath, xmlNamespaceManager);
        }
    }
}