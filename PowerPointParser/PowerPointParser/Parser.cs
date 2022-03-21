using System.Text;
using System.Xml;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Aaks.PowerPointParser.Dto;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aaks.PowerPointParser
{
    public class Parser : IParser
    {
        private const string XpathNotesToSp = @"/*[local-name() = 'notes']/*[local-name() = 'cSld']/*[local-name() = 'spTree']/*[local-name() = 'sp']";
        private const string PNodesListXPath = @"/*[local-name() = 'sp']/*[local-name() = 'txBody']/*[local-name() = 'p']";
        public IDictionary<int, IList<OpenXmlParagraphWrapper?>> ParseSpeakerNotes(MemoryStream memoryStream)
        {

            using var presentationDocument = PresentationDocument.Open(memoryStream, false);
            var slidesContentMap = ParseSpeakerNotes(presentationDocument);
            return slidesContentMap;

        }
        public IDictionary<int, IList<OpenXmlParagraphWrapper?>> ParseSpeakerNotes(string path)
        {
            
            using var presentationDocument = PresentationDocument.Open(path, false);
            var slidesContentMap = ParseSpeakerNotes(presentationDocument);
            return slidesContentMap;

        }

        private IDictionary<int, IList<OpenXmlParagraphWrapper?>> ParseSpeakerNotes(PresentationDocument presentationDocument)
        {
            var slidesContentMap = new Dictionary<int, IList<OpenXmlParagraphWrapper>>();
            var presentationPart = presentationDocument.PresentationPart;
            if (presentationPart == null) return slidesContentMap!;

            var presentation = presentationPart.Presentation;

            if (presentation.SlideIdList == null) return slidesContentMap!;

            var slideIds = presentation.SlideIdList.Elements<SlideId>();

            int slideIndex = 1;

            foreach (var slideId in slideIds)
            {
                var note = GetNotesSlidePart(presentationPart, slideId);
                var openXmlParagraphWrappers = new List<OpenXmlParagraphWrapper>();

                if (DoesSlideHaveSpeakerNotes(note))
                {
                    var pNodesList = ParsePNodesList(note!);

                    if (pNodesList != null)
                    {
                        var xmlSerializer = new XmlSerializer(typeof(OpenXmlParagraphWrapper));
                        foreach (XmlNode node in pNodesList)
                        {
                            try
                            {
                                using StringReader stringReader = new(node.OuterXml);
                                var wrapper = (OpenXmlParagraphWrapper)xmlSerializer.Deserialize(stringReader)!;
                                openXmlParagraphWrappers.Add(wrapper);
                            }
                            catch (InvalidOperationException ex)
                            {
                                Console.WriteLine($"{ex.Message} Slide Note Deserialization Failed");

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"{ex.Message} Unknown Exception Occurred");
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
            MemoryStream stream = new MemoryStream(bytes);
            bodyNodeXmlDocument.Load(stream);

            return bodyNodeXmlDocument.SelectNodes(PNodesListXPath, xmlNamespaceManager);
        }
    }
}