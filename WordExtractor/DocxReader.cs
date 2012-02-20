using System;
using System.Collections.Generic;
using System.Linq;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.IO;

namespace WordExtractor
{
    public class DocxReader
    {
        const string DocumentXmlPath = "/word/document.xml";
        const string DocumentRelsXmlPath = "/word/_rels/document.xml.rels";
        const string FootnotesXmlPath = "/word/footnotes.xml";
        const string FootnotesRelsXmlPath = "/word/_rels/footnotes.xml.rels";
        const string NumberingXmlPath = "/word/numbering.xml";

        static readonly Uri DocumentXmlUri = new Uri(DocumentXmlPath, UriKind.Relative);
        static readonly Uri DocumentRelsXmlUri = new Uri(DocumentRelsXmlPath, UriKind.Relative);
        static readonly Uri FootnotesXmlUri = new Uri(FootnotesXmlPath, UriKind.Relative);
        static readonly Uri FootnotesRelsXmlUri = new Uri(FootnotesRelsXmlPath, UriKind.Relative);
        static readonly Uri NumberingXmlUri = new Uri(NumberingXmlPath, UriKind.Relative);

        const string WordMLSchema = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string RelationsSchema = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string MathSchema = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string CustomXmlSchema = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        const string HyperlinkSchema = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        public DocxReader(string sourceFile) {
            using (var SourceFile = Package.Open(sourceFile, System.IO.FileMode.Open, FileAccess.Read, FileShare.Read)) {
                Document = Read(SourceFile.GetPart(DocumentXmlUri));
                DocumentRels = Read(SourceFile.GetPart(DocumentRelsXmlUri));
                try {
                    Footnotes = Read(SourceFile.GetPart(FootnotesXmlUri));
                    FootnotesRels = Read(SourceFile.GetPart(FootnotesRelsXmlUri));
                } catch (InvalidOperationException) {
                    if (Footnotes == null) {
                        Footnotes = new List<Token>();
                    }
                    FootnotesRels = new List<Token>();
                }
                try {
                    Numbering = Read(SourceFile.GetPart(NumberingXmlUri));
                } catch (InvalidOperationException) {
                    Numbering = new List<Token>();
                }
            }
        }

        public IList<Token> Document { get; private set; }
        public IList<Token> DocumentRels { get; private set; }
        public IList<Token> Footnotes { get; private set; }
        public IList<Token> FootnotesRels { get; private set; }
        public IList<Token> Numbering { get; private set; }

        private List<Token> Read(PackagePart part) {
            // Read part
            if (part == null) throw new EndOfStreamException();

            XElement xml;
            using (var stream = part.GetStream(FileMode.Open, FileAccess.Read)) {
                xml = XElement.Load(stream, LoadOptions.None);
            }

            // Read contents
            var result = new List<Token>();
            RecurseFillBuffer(xml, result);
            return result;
        }

        private void RecurseFillBuffer(XElement xml, List<Token> buffer) {
            if (xml.Name.LocalName == "oMath") {
                var reader = xml.CreateReader();
                reader.Read();
                buffer.Add(new Token("Smath", reader.ReadOuterXml()));
                return;
            }

            buffer.Add(new Token("S<", xml.Name.LocalName));
            foreach (var a in xml.Attributes()) {
                buffer.Add(new Token("Sa", a.Name.LocalName));
                buffer.Add(new Token("S=", a.Value));
            }
            foreach (var e in xml.Elements()) {
                RecurseFillBuffer(e, buffer);
            }
            if (!xml.HasElements && !string.IsNullOrEmpty(xml.Value)) {
                if (xml.Value == "*")
                    buffer.Add(new Token("\\*"));
                else
                    buffer.Add(new Token(xml.Value));
            }
            buffer.Add(new Token("S>", xml.Name.LocalName));
        }


    }
}
