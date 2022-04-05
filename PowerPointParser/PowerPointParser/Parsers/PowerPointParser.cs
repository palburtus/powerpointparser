﻿using System.Text;
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

        public IDictionary<int, IList<Paragraph?>> ParseSlide(string path)
        {
            using var presentationDocument = PresentationDocument.Open(path, false);
            var slidesContentMap = ParseSlides(presentationDocument);
            return slidesContentMap!;
        }

        public PowerPointParser(ILogger<PowerPointParser>? logger = null)
        {
            _logger = logger;
        }

        public IDictionary<int, IList<OpenXmlTextWrapper?>> ParseSpeakerNotes(MemoryStream memoryStream)
        {
            using var presentationDocument = PresentationDocument.Open(memoryStream, false);
            var slidesContentMap = ParseSpeakerNotes(presentationDocument);
            return slidesContentMap;

        }

        public IDictionary<int, IList<OpenXmlTextWrapper?>> ParseSpeakerNotes(string path)
        {
            using var presentationDocument = PresentationDocument.Open(path, false);
            var slidesContentMap = ParseSpeakerNotes(presentationDocument);
            return slidesContentMap;
        }

        private IDictionary<int, IList<Paragraph>?> ParseSlides(PresentationDocument presentationDocument)
        {
            IDictionary<int, IList<Paragraph>?> slidesContentMap = new Dictionary<int, IList<Paragraph?>>()!;
            var presentationPart = presentationDocument.PresentationPart;
            if (presentationPart == null) return slidesContentMap!;

            var presentation = presentationPart.Presentation;

            if (presentation.SlideIdList == null) return slidesContentMap;

            var slideIds = presentation.SlideIdList.Elements<SlideId>();

            int slideIndex = 1;

            foreach (var slideId in slideIds)
            {
                var openXmlItems = new List<Paragraph>();

                if (slideId == null) return null;
                if (slideId.RelationshipId == null) return null;

                var openXmlPart = presentationPart.GetPartById(slideId.RelationshipId!);
                SlidePart? slidePart = openXmlPart as SlidePart;

                
                foreach (var item in slidePart?.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>()!)
                {
                    openXmlItems.Add(item);
                }

                slideIndex++;

                slidesContentMap.Add(slideIndex, openXmlItems);
            }


            return slidesContentMap;
        }

        private IDictionary<int, IList<OpenXmlTextWrapper?>> ParseSpeakerNotes(PresentationDocument presentationDocument)
        {
            var slidesContentMap = new Dictionary<int, IList<OpenXmlTextWrapper>>();
            var presentationPart = presentationDocument.PresentationPart;
            if (presentationPart == null) return slidesContentMap!;

            var presentation = presentationPart.Presentation;

            if (presentation.SlideIdList == null) return slidesContentMap!;

            var slideIds = presentation.SlideIdList.Elements<SlideId>();

            int slideIndex = 1;

            foreach (var slideId in slideIds)
            {
                var note = GetNotesSlidePart(presentationPart, slideId);
                var openXmlParagraphWrappers = new List<OpenXmlTextWrapper>();

                if (DoesSlideHaveSpeakerNotes(note))
                {
                    //var pNodesList = note.NotesSlide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>();
                    var pNodesList = ParsePNodesList(note!);

                    if (pNodesList != null)
                    {
                        var xmlSerializer = new XmlSerializer(typeof(OpenXmlTextWrapper));
                        foreach (/*var*/XmlNode node in pNodesList)
                        {
                            try
                            {
                                using StringReader stringReader = new(node.OuterXml);
                                using XmlTextReader xmlReader = new (stringReader);
                                var wrapper = (OpenXmlTextWrapper)xmlSerializer.Deserialize(xmlReader)!;
                                openXmlParagraphWrappers.Add(wrapper);
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
                        }
                    }
                }

                slidesContentMap.Add(slideIndex, openXmlParagraphWrappers);
                slideIndex++;
            }

            return slidesContentMap!;
        }
        private static NotesSlidePart? GetNotesSlidePart(OpenXmlPartContainer presentationPart, SlideId? slideId)
        {
            if (slideId == null) return null;
            if (slideId.RelationshipId == null) return null;
            
            var openXmlPart = presentationPart.GetPartById(slideId.RelationshipId!);
            
            SlidePart? slidePart = openXmlPart as SlidePart;

            return slidePart?.NotesSlidePart;
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
            using var stream = new MemoryStream(bytes);
            bodyNodeXmlDocument.Load(stream);

            return bodyNodeXmlDocument.SelectNodes(PNodesListXPath, xmlNamespaceManager);
        }
    }
}