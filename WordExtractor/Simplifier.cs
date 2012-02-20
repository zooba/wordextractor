using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Xml.Xsl;
using System.Xml.XPath;
using System.Xml;

namespace WordExtractor
{
    public class Simplifier
    {
        private DocxReader Reader;
        private TextWriter Warnings;

        private int UpTo;
        private LinkedList<Token> DocumentTokens;
        private LinkedList<Token> DocumentRelsTokens;
        private LinkedList<Token> FootnoteTokens;
        private LinkedList<Token> FootnoteRelsTokens;
        private LinkedList<Token> NumberingTokens;

        private Dictionary<string, string> NiceReferenceNames;
        private HashSet<string> IgnoredBookmarks;

        public IEnumerable<Token> Document { get { return DocumentTokens.AsEnumerable(); } }

        public Simplifier(DocxReader reader, TextWriter warnings) {
            Reader = reader;
            DocumentTokens = new LinkedList<Token>(reader.Document);
            DocumentRelsTokens = new LinkedList<Token>(reader.DocumentRels);
            FootnoteTokens = new LinkedList<Token>(reader.Footnotes);
            FootnoteRelsTokens = new LinkedList<Token>(reader.FootnotesRels);
            NumberingTokens = new LinkedList<Token>(reader.Numbering);
            NiceReferenceNames = new Dictionary<string, string>();
            IgnoredBookmarks = new HashSet<string>();
            UpTo = 0;
            Warnings = warnings;
        }

        public void RunAll(int limit = 9) {
            if (UpTo >= limit) return;
            Pass1();
            if (UpTo >= limit) return;
            Pass2();
            if (UpTo >= limit) return;
            Pass3();
            if (UpTo >= limit) return;
            Pass4();
            if (UpTo >= limit) return;
            Pass5();
            if (UpTo >= limit) return;
        }

        internal void Pass1() {
            if (UpTo > 0) return;

            ConvertFootnoteCharacters();
            ConvertFootnoteRuns();

            ConvertSoftHyphens();

            RenameInstrTextRuns();
            RemoveDeletedTextRuns();

            RemoveInvisibleText();
            RemoveBibliography();

            ConvertFieldCharacters();
            ConvertBookmarks();
            ConvertImages();

            RemoveRSIDAttributes();
            RemoveEmptyElements("noProof");
            RemoveEmptyElements("lastRenderedPageBreak");
            RemoveEmptyElements("rPr");
            RemoveProofingErrors();

            CondenseTables();

            ConvertBreaks();

            ReplaceEmptyElements("tab", "hspace");

            UpTo = 1;
        }

        internal void Pass2() {
            if (UpTo < 1) Pass1();
            if (UpTo > 1) return;

            RemoveTextElementBrackets();

            ExtractSdtContent();

            ConvertMath();
            ConvertHyperlinks();

            ConvertNumbering();
            ConvertStyles();

            IgnoreProperties();
            RemoveProperties();
            SimplifyRuns();

            UpTo = 2;
        }

        internal void Pass3() {
            if (UpTo < 2) Pass2();
            if (UpTo > 2) return;


            SimplifyParagraphs();
            CombineRunStyles();
            ConvertStylesToEnvironments();
            ConvertBookmarksToLabels();

            TidyFootnotes();

            CombineNumberedLists();

            TidyDocumentStartEnd();

            ConvertFields();
            ResolveURLs();

            UpTo = 3;
        }

        internal void Pass4() {
            if (UpTo < 3) Pass3();
            if (UpTo > 3) return;

            ConvertTables();
            ConvertCaptions();
            ConvertFloats();

            ConvertOtherBookmarks();

            UpTo = 4;
        }

        internal void Pass5() {
            if (UpTo < 4) Pass4();
            if (UpTo > 4) return;

            DetectCodeListings();
            //UseNiceReferenceNames();
            CombineReferences();
            CombineCitations();

            //WrapDottedNames();
            MarkTableFirstRow();

            UpTo = 5;
        }

        private struct FindResult
        {
            public bool IsMatch { get { return Start != null; } }
            public LinkedList<Token> List { get { return Start.List ?? Mark.List ?? End.List; } }

            public LinkedListNode<Token> Start;
            public LinkedListNode<Token> Mark;
            public LinkedListNode<Token> End;

            public static FindResult Empty = new FindResult { Start = null, Mark = null, End = null };
        }

        private FindResult FindSequence(LinkedListNode<Token> start, string code) {
            int offset = 0, index = 0;
            var codeParts = code.Split(' ').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();

            for (int i = 0; i < codeParts.Count; i += 1) {
                if (codeParts[i] == "|") {
                    offset = i;
                    codeParts.RemoveAt(i);
                    i -= 1;
                } else if (codeParts[i] == "!") {
                    codeParts[i] = "!\0:\0";
                }
            }

            var codeTokens = codeParts.Select(Token.Parse).ToList();
            var result = new FindResult();

            for (var c = start; c != null; c = c.Next) {
                if (index >= codeTokens.Count) {
                    result.End = c.Previous;
                    return result;
                }

                if (codeTokens[index].Metadata != null && codeTokens[index].Metadata.StartsWith("!")) {
                    var t = codeTokens[index];
                    var abortToken = new Token(t.Metadata.Substring(1), t.Value);
                    index += 1;
                    var nextToken = (index < codeTokens.Count) ? codeTokens[index] : null;
                    while (c != null && !c.Value.WildcardEquals(abortToken) && !c.Value.WildcardEquals(nextToken)) c = c.Next;
                    if (c == null) return FindResult.Empty;
                    if (c.Value.WildcardEquals(abortToken)) index = 0;
                }

                if (index >= codeTokens.Count) continue;
                else if (c.Value.WildcardEquals(codeTokens[index])) {
                    if (index == 0) result.Start = c;
                    if (index == offset) result.Mark = c;
                    index += 1;
                } else {
                    index = 0;
                    c = result.Start ?? c;
                    result.Start = null;
                }
            }

            if (index >= codeTokens.Count) {
                result.End = start.List.Last;
                return result;
            }
            return FindResult.Empty;
        }

        private IEnumerable<FindResult> Find(string code) {
            LinkedListNode<Token> next = null;
            for (var c = FindSequence(DocumentTokens.First, code); c.IsMatch; c = FindSequence(next, code)) {
                next = c.Start.Previous;
                yield return c;
                if (next == null) {
                    next = DocumentTokens.First.Next;
                } else {
                    next = next.Next;
                    if (next != null) next = next.Next;
                    else break;
                }
            }
        }

        private int ReplaceSequence(LinkedListNode<Token> start, string code, string replace, int count = int.MaxValue) {
            var replaceParts = replace.Split(' ').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
            var replaceTokens = replaceParts.Select(Token.Parse).ToList();

            LinkedListNode<Token> next = start;
            int index = 0;
            int replacements = 0;
            for (var c = FindSequence(next, code); index < count && c.IsMatch; c = FindSequence(next, code), index += 1) {
                var list = c.List;
                var tokens = new List<Token>();
                for (var c2 = c.Start; c2 != null && c2 != c.End.Next; c2 = c2.Next) {
                    tokens.Add(c2.Value);
                }

                var pre = c.Start.Previous;
                var post = c.End.Next;
                next = post;

                if (pre == null) {
                    while (list.First != post) list.RemoveFirst();

                    if (replaceTokens.Any()) list.AddFirst(new Token(replaceTokens.First()));
                    pre = list.First;
                } else {
                    while (pre.Next != post) list.Remove(pre.Next);

                    if (replaceTokens.Any()) list.AddAfter(pre, new Token(replaceTokens.First()));
                    pre = pre.Next;
                }

                foreach (var t in replaceTokens.Skip(1)) {
                    list.AddAfter(pre, new Token(t));
                    pre = pre.Next;
                }

                replacements += 1;
            }

            return replacements;
        }

        private void ConvertSoftHyphens() {
            ReplaceSequence(DocumentTokens.First, "S<:softHyphen S>:softHyphen", "\0:\x00AD");
        }

        private void RenameInstrTextRuns() {
            ReplaceSequence(DocumentTokens.First, "S<:instrText", "S<:t");
            ReplaceSequence(DocumentTokens.First, "S>:instrText", "S>:t");
        }

        private void RemoveDeletedTextRuns() {
            ReplaceSequence(DocumentTokens.First, "S<:p !S:<r S<:rPr !S>:rPr S<:del ! S>:p", "");
        }

        private void RemoveInvisibleText() {
            ReplaceSequence(DocumentTokens.First, "S<:r S<:rPr !S>:rPr S<:vanish S>:vanish ! S>:r", "");
            ReplaceSequence(DocumentTokens.First, "S<:p !S<:r S<:rPr !S>:rPr S<:vanish S>:vanish ! S>:p", "");
        }

        private void RemoveBibliography() {
            ReplaceSequence(DocumentTokens.First, "S<:sdt !S>:sdt S<:sdt !S>:sdt S<:bibliography S>:bibliography ! S>:sdt ! S>:sdt", "bibliography:");
        }

        private void RemoveProofingErrors() {
            ReplaceSequence(DocumentTokens.First, "S<:proofErr ! S>:proofErr", "");
        }

        private void ConvertFieldCharacters() {
            var code = "S<:fldChar Sa:fldCharType | S=:* S>:fldChar";

            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = "field_" + c.Mark.Value.Value;
                c.Start.Value.Value = null;
                c.Start.Next.RemoveTo(c.End);
            }
        }

        private void ConvertBookmarks() {
            var code = "S<:bookmarkStart !S>:bookmarkStart Sa:id S=:* !S>:bookmarkStart Sa:name | S=:* ! S>:bookmarkStart";

            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = "bookmark_start";
                var name = c.Start.Value.Value = c.Mark.Value.Value;

                var id = FindSequence(c.Start, "Sa:id | S=:*").Mark.Value.Value;
                c.Start.Next.RemoveTo(c.End);

                var c2 = FindSequence(c.Start, "| S<:bookmarkEnd Sa:id S=:" + id + " ! S>:bookmarkEnd");
                if (!c2.IsMatch) continue;
                c2.Mark.Value.Metadata = "bookmark_end";
                c2.Mark.Value.Value = name;
                c2.Start.Next.RemoveTo(c2.End);

                if (!name.StartsWith("_Ref")) {
                    IgnoredBookmarks.Add(name);
                    c.Start.Remove();
                    c2.Mark.Remove();
                }
            }

            ReplaceSequence(DocumentTokens.First, "bookmark_start:_GoBack bookmark_end:_GoBack", "");
        }

        private void ConvertBookmarksToLabels() {
            var code = "para_style:* bookmark_start:* ! bookmark_end:* eop:*";
            foreach (var c in Find(code)) {
                if (c.Start.Value.Value.StartsWith("Heading", StringComparison.CurrentCultureIgnoreCase) ||
                    c.Start.Value.Value.StartsWith("Appendix", StringComparison.CurrentCultureIgnoreCase)) {
                    var label = c.Start.Next.Value.Value.TrimStart('_');
                    c.List.Remove(c.Start.Next);
                    c.List.Remove(c.End.Previous);
                    c.List.AddAfter(c.End, new Token("label", label));
                }
            }
        }

        private void ConvertOtherBookmarks() {
            var code = "bookmark_start:*";
            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = "label";
                ReplaceSequence(c.Start, "bookmark_end:" + c.Start.Value.Value, "", 1);
            }
        }

        private void ConvertImages() {
            foreach (var c in Find("S<:p !S>:p S<:drawing !S>:drawing S<:docPr !S>:docPr Sa:title | S=:* ! S>:p")) {
                c.Start.Value.Metadata = "image";
                c.Start.Value.Value = c.Mark.Value.Value;
                c.Start.Next.RemoveTo(c.End);
            }
            ReplaceSequence(DocumentTokens.First, "S<:drawing ! S>:drawing", "image:\0");
        }

        private void ConvertFootnoteCharacters() {
            var code = "S<:r !S>:r S<:footnoteReference Sa:id | S=:* S>:footnoteReference ! S>:r";

            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = "footnote_ref";
                c.Start.Value.Value = c.Mark.Value.Value;
                c.Start.Next.RemoveTo(c.End);
            }
        }

        private void ConvertFootnoteRuns() {
            var code = "footnote_ref:*";

            foreach (var c in Find(code)) {
                var id = c.Mark.Value.Value;
                var match = FindSequence(FootnoteTokens.First, "S<:footnote !S>:footnote Sa:id | S=:" + id + " ! S>:footnote");

                if (match.IsMatch) {
                    var dest = c.List.AddAfter(c.Mark, match.Mark.Next, match.End.Previous);
                    c.Mark.Value.Metadata = "footnote";
                    c.List.AddAfter(dest, new Token("end_footnote", c.Mark.Value.Value));
                }
            }
        }

        private void TidyFootnotes() {
            var code = "footnote:* ! \0:*";
            foreach (var c in Find(code)) {
                c.Mark.Next.RemoveTo(c.End.Previous);
                c.End.Value.Value = c.End.Value.Value.TrimStart(' ');
            }

            code = "eop:* end_footnote:*";
            foreach (var c in Find(code)) {
                c.Start.Remove();
            }
        }

        private void RemoveEmptyElements(string elementName) {
            ReplaceSequence(DocumentTokens.First, string.Format("S<:{0} S>:{0}", elementName), "");
        }

        private void ReplaceEmptyElements(string elementName, string newMetadata) {
            ReplaceSequence(DocumentTokens.First, string.Format("S<:{0} S>:{0}", elementName), newMetadata + ":\0");
        }

        private void ConvertBreaks() {
            var code = "S<:br S>:br";

            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = "linebreak";
                c.Start.Value.Value = null;
                c.End.Remove();
            }

            code = "S<:br Sa:type | S=:* S>:br";

            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = c.Mark.Value.Value + "break";
                c.Start.Value.Value = null;
                c.Start.Next.RemoveTo(c.End);
            }

            code = "S<:sectPr !S>:sectPr S<:type Sa:val S=:oddPage ! S>:sectPr";

            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = "oddsectionbreak";
                c.Start.Value.Value = null;
                c.Start.Next.RemoveTo(c.End);
            }            
            
            code = "S<:sectPr ! S>:sectPr";

            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = "sectionbreak";
                c.Start.Value.Value = null;
                c.Start.Next.RemoveTo(c.End);
            }
        }

        private void RemoveRSIDAttributes() {
            var codes = new[]
            {
                "Sa:rsidR S=:*",
                "Sa:rsidRPr S=:*",
                "Sa:rsidRDefault S=:*",
                "Sa:rsidP S=:*",
                "Sa:rsidPPr S=:*",
                "Sa:rsidTr S=:*",
            };

            foreach (var code in codes) {
                ReplaceSequence(DocumentTokens.First, code, "");
            }
        }

        private void RemoveTextElementBrackets() {
            ReplaceSequence(DocumentTokens.First, "S<:t Sa:space S=:preserve", "");
            ReplaceSequence(DocumentTokens.First, "S<:t", "");
            ReplaceSequence(DocumentTokens.First, "S>:t", "");
        }

        private void ExtractSdtContent() {
            ReplaceSequence(DocumentTokens.First, "S<:sdt ! S<:sdtContent", "");
            ReplaceSequence(DocumentTokens.First, "S>:sdtContent S>:sdt", "");
        }

        private void ConvertNumbering() {
            var code = "S<:pPr !S>:pPr | S<:numPr !S>:pPr S>:numPr";
            var levelCode = "S<:ilvl Sa:val | S=:*";
            var idCode = "S<:numId Sa:val | S=:*";

            var formatCode = "S<:abstractNum Sa:abstractNumId S=:{0} !S>:abstractNum S<:lvl Sa:ilvl S=:{1} !S>:lvl S<:numFmt Sa:val | S=:*";

            foreach (var c in Find(code)) {
                var c2 = FindSequence(c.Start, levelCode);
                var c3 = FindSequence(c.Start, idCode);

                var level = c2.IsMatch ? c2.Mark.Value.Value : "0";
                string id = null;
                string type = "unknown";

                if (c3.IsMatch) {
                    id = c3.Mark.Value.Value;
                    var c4 = FindSequence(NumberingTokens.First, "S<:num Sa:numId S=:" + id + " S<:abstractNumId Sa:val | S=:*");
                    if (c4.IsMatch) {
                        id = c4.Mark.Value.Value;
                        var c5 = FindSequence(NumberingTokens.First, string.Format(formatCode, id, level));
                        if (c5.IsMatch) {
                            type = c5.Mark.Value.Value;
                        }
                    }
                }

                c.Mark.Value.Metadata = "listitem";
                c.Mark.Value.Value = type + "," + level;
                c.Mark.Next.RemoveTo(c.End);
            }
        }

        private void CombineNumberedLists() {
            var code = "para_style:* listitem:*";

            foreach (var c in Find(code)) {
                c.Start.Remove();
            }

            code = "listitem:*";

            for (var c = FindSequence(DocumentTokens.First, code); c.IsMatch; c = FindSequence(c.End, code)) {
                var type = c.Start.Value.Value;
                if (type.Contains(',')) type = type.Substring(0, type.IndexOf(','));
                c.List.AddBefore(c.Start, new Token("list", type));

                var nextCode = string.Format("listitem:{0} ! eop:* !eop:* | listitem:{0}", c.Start.Value.Value);
                var expected = c.Start;
                for (var next = FindSequence(c.Start, nextCode);
                    next.IsMatch && next.Start == expected;
                    next = FindSequence(next.Mark, nextCode)) {
                    next.Start.Value.Metadata = "item";
                    next.Start.Value.Value = null;
                    expected = next.Mark;
                }

                expected.Value.Metadata = "item";
                expected.Value.Value = null;

                var lastCode = string.Format("item:\0 ! eop:*", expected);
                var last = FindSequence(expected, lastCode);
                if (last.IsMatch && last.Start == expected) {
                    last.List.AddAfter(last.End, new Token("end_list", type));
                } else {
                    Warnings.WriteLine("Something broke...");
                }
                c.End = expected;
            }
        }

        private void ConvertStyles() {
            var code = "S<:r !S>:r S<:rPr !S>:rPr S<:rStyle Sa:val | S=:* S>:rStyle";
            foreach (var c in Find(code)) {
                var first = c.Start.Next;
                first.Value.Metadata = "run_style";
                first.Value.Value = c.Mark.Value.Value;
                first.Next.RemoveTo(c.End);
            }
            ReplaceSequence(DocumentTokens.First, "S<:rStyle ! S>:rStyle", "");


            code = "S<:pPr !S>:pPr S<:pStyle Sa:val | S=:* S>:pStyle";
            foreach (var c in Find(code)) {
                var first = c.Start.Next;
                first.Value.Metadata = "para_style";
                first.Value.Value = c.Mark.Value.Value;
                first.Next.RemoveTo(c.End);
            }
        }

        private void ConvertFields() {
            var code = "S<:fldSimple Sa:instr | S=:* ! S>:fldSimple";
            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = "field_begin";
                c.List.Remove(c.Start.Next);
                c.Mark.Value.Metadata = null;
                c.End.Value.Metadata = "field_end";
            }

            code = "field_begin:* | \0:* ! field_end:*";
            foreach (var c in Find(code)) {
                var command = c.Mark.Value.Value.Trim();
                var options = command.Substring(command.IndexOf(' ')).TrimStart();
                command = command.Substring(0, command.IndexOf(' ')).ToUpperInvariant();
                var firstOption = (options.Contains(' ')) ? (options.Substring(0, options.IndexOf(' '))) : options;

                if (command == "CITATION" || command == "SEQ") {
                    c.Start.Value.Metadata = command.ToLowerInvariant();
                    c.Start.Value.Value = firstOption;
                    c.Start.Next.RemoveTo(c.End);
                } else if (command == "HYPERLINK") {
                    c.Start.Value.Metadata = command.ToLowerInvariant();
                    c.Start.Value.Value = options.Trim(' ', '"', '\'');
                    c.Start.Next.RemoveTo(c.End);
                } else if (command == "REF" || command == "PAGEREF") {
                    if (IgnoredBookmarks.Contains(firstOption)) {
                        while (c.Start.Next != null && c.Start.Next.Value.Metadata != "field_separate") {
                            c.Start.Next.Remove();
                        }
                        c.Start.Next.Remove();
                        c.Start.Remove();
                        c.End.Remove();
                        command = null;
                    } else {
                        c.Start.Value.Metadata = command.ToLowerInvariant();
                        c.Start.Value.Value = firstOption;
                        c.Start.Next.RemoveTo(c.End);
                    }
                } else if (command == "XE") {
                    c.Start.Value.Metadata = "index";
                    int i1 = options.IndexOf('"') + 1;
                    int i2 = options.IndexOf('"', i1);
                    var term = options.Substring(i1, i2 - i1).Replace(':', '!');
                    if (options.IndexOf("\\t", i2) > 0) {
                        var seeTermMatch = FindSequence(c.Mark.Next, "\0:* !\0:* field_end:*");
                        if (seeTermMatch.IsMatch) {
                            var seeTerm = seeTermMatch.Mark.Value.Value;
                            if (seeTerm.StartsWith("See")) {
                                seeTerm = seeTerm.Substring(4);
                            }
                            seeTerm = seeTerm.Trim('"', ' ').Replace(':', '!');
                            term += "|see{" + seeTerm + "}";
                        }
                    }
                    c.Start.Value.Value = term;
                    c.Start.Next.RemoveTo(c.End);
                } else if (command == "INDEX") {
                    c.Start.Value.Metadata = "printindex";
                    c.Start.Value.Value = null;
                    c.Start.Next.RemoveTo(c.End);
                } else {
                    Debug.WriteLine("Unhandled field: " + command);
                }

                if (command == "REF") {
                    c.Start.Value.Value = c.Start.Value.Value.TrimStart('_');
                    if (options.Contains("\\# 0")) {
                        c.Start.Value.Metadata = "ref_number";
                    } else if (options.Contains("\\#")) {
                        c.Start.Remove();
                    }
                }
            }
        }

        private void ResolveURLs() {
            var code = "url_ref:*";
            foreach (var c in Find(code)) {
                var t = c.Start.Value;
                var code2 = "Sa:Id S=:" + c.Start.Value.Value + " !S>:Relationship Sa:Target | S=:*";
                var c2 = FindSequence(DocumentRelsTokens.First, code2);
                if (!c2.IsMatch) {
                    t.Metadata = "error";
                    t.Value = "Unresolved url reference " + t.Value;
                } else {
                    t.Metadata = "hyperlink";
                    t.Value = c2.Mark.Value.Value;
                }
            }

            code = "url_ref_footnote:*";
            foreach (var c in Find(code)) {
                var t = c.Start.Value;
                var code2 = "Sa:Id S=:" + c.Start.Value.Value + " !S>:Relationship Sa:Target | S=:*";
                var c2 = FindSequence(FootnoteRelsTokens.First, code2);
                if (!c2.IsMatch) {
                    t.Metadata = "error";
                    t.Value = "Unresolved url reference " + t.Value;
                } else {
                    t.Metadata = "hyperlink";
                    t.Value = c2.Mark.Value.Value;
                }
            }
        }

        private void CombineReferences() {
            var code = "ref:* | \0:* ref:*";
            foreach (var c in Find(code)) {
                if (c.Mark.Value.Value.Trim().Equals(",")) c.Mark.Remove();
                if (c.Mark.Value.Value.Trim().Equals("and", StringComparison.CurrentCultureIgnoreCase)) c.Mark.Remove();
            }

            code = "ref:* ref:*";
            for (var c = FindSequence(DocumentTokens.First, code); c.IsMatch; c = FindSequence(c.Start, code)) {
                c.Start.Value.Value += "," + c.End.Value.Value;
                c.End.Remove();
            }

            code = "\0:* ref:*";
            foreach (var c in Find(code)) {
                if (c.Start.Value.Value.TrimEnd().EndsWith(".")) c.End.Value.Metadata = "Ref";
            }
            code = "eop:* ref:*";
            foreach (var c in Find(code)) {
                c.End.Value.Metadata = "Ref";
            }
        }

        private void IgnoreProperties() {
            var codes = new[] {
                "S<:rPr !S>:rPr | S<:kern ! S>:kern",
                "S<:rPr !S>:rPr | S<:rFonts ! S>:rFonts",
                "S<:rPr !S>:rPr | S<:b ! S>:b",
                "S<:rPr !S>:rPr | S<:bCs ! S>:bCs",
                "S<:rPr !S>:rPr | S<:sz ! S>:sz",
                "S<:rPr !S>:rPr | S<:szCs ! S>:szCs",
                "S<:pPr !S>:pPr | S<:jc ! S>:jc",
                "S<:pPr !S>:pPr | S<:ind ! S>:ind",
                "S<:lang ! S>:lang",
                "S<:proofErr ! S>:proofErr",
                "S<:numForm ! S>:numForm",
                "S<:highlight ! S>:highlight",
                "S<:cnfStyle ! S>:cnfStyle",
                "S<:spacing ! S>:spacing",
            };

            foreach (var code in codes) {
                foreach (var c in Find(code)) {
                    c.Mark.RemoveTo(c.End);
                }
            }
        }

        private void RemoveProperties() {
            ReplaceSequence(DocumentTokens.First, "S<:rPr", "");
            ReplaceSequence(DocumentTokens.First, "S>:rPr", "");
            ReplaceSequence(DocumentTokens.First, "S<:pPr", "");
            ReplaceSequence(DocumentTokens.First, "S>:pPr", "");
        }

        private void SimplifyRuns() {
            var code = "S<:r | run_style:* ! S>:r";
            foreach (var c in Find(code)) {
                c.List.Remove(c.Start);
                c.End.Value.Metadata = "end_run_style";
                c.End.Value.Value = c.Mark.Value.Value;
            }

            code = "S<:r ! S>:r";
            foreach (var c in Find(code)) {
                c.List.Remove(c.Start);
                c.List.Remove(c.End);
            }

            code = "run_style:* field_begin:* end_run_style:*";
            foreach (var c in Find(code)) {
                c.Start.Remove();
                c.End.Remove();
            }
            code = "run_style:* field_separate:* end_run_style:*";
            foreach (var c in Find(code)) {
                c.Start.Remove();
                c.End.Remove();
            }
            code = "run_style:* field_end:* end_run_style:*";
            foreach (var c in Find(code)) {
                c.Start.Remove();
                c.End.Remove();
            }
        }

        private void SimplifyParagraphs() {
            ReplaceSequence(DocumentTokens.First, "S<:p", "");
            ReplaceSequence(DocumentTokens.First, "S>:p", "eop:");
            ReplaceSequence(DocumentTokens.First, "para_style:BodyText", "");

            var code = "\0:* \0:*";
            for (var c = FindSequence(DocumentTokens.First, code); c.IsMatch; c = FindSequence(c.Start, code)) {
                c.Start.Value.Value += c.End.Value.Value;
                c.List.Remove(c.End);
            }
        }

        private void CombineRunStyles() {
            var code = "end_run_style:* run_style:*";

            foreach (var c in Find(code)) {
                if (c.Start.Value.Value == c.End.Value.Value) {
                    c.Start.Remove();
                    c.End.Remove();
                }
            }
        }

        private void CombineCitations() {
            var code = "citation:* citation:*";
            for (var c = FindSequence(DocumentTokens.First, code); c.IsMatch; c = FindSequence(c.Start, code)) {
                c.Start.Value.Value += "," + c.End.Value.Value;
                c.End.Remove();
            }
        }

        private void ConvertStylesToEnvironments() {
            var styles = new[] { "Abstract" };
            foreach (var style in styles) {
                var code = "para_style:" + style + " ! | eop:";
                foreach (var c in Find(code)) {
                    c.Start.Value.Metadata = "wxbegin";
                    c.Start.Value.Value = style.ToLowerInvariant();
                    c.Mark.Value.Metadata = "wxend";
                    c.Mark.Value.Value = style.ToLowerInvariant();
                }
            }

            {
                var code = "para_style:* !eop:* | \0:*";
                foreach (var c in Find(code)) {
                    if (c.Start.Value.Value.StartsWith("heading", StringComparison.InvariantCultureIgnoreCase)) {
                        if (c.Mark.Value.Value.StartsWith("appendix", StringComparison.CurrentCultureIgnoreCase)) {
                            c.Start.Value.Value = "Appendix";
                            var match = Regex.Match(c.Mark.Value.Value, @"appendix.*?(\:|\.) *(.+)", RegexOptions.IgnoreCase);
                            if (match != null && match.Groups.Count == 3) {
                                c.Mark.Value.Value = match.Groups[2].Value;
                            }
                        }
                    }
                }
            }
        }

        private void ConvertMath() {
            var code = "Smath:*";

            foreach (var c in Find(code)) {
                var xslt1 = new XslCompiledTransform();
                var xslt2 = new XslCompiledTransform();
                xslt1.Load(typeof(OMMLMML));
                xslt2.Load(typeof(MMLTeX));

                var xpathdoc = new XPathDocument(new StringReader(c.Mark.Value.Value));
                var result = new StringWriter();
                var resultXml = XmlWriter.Create(result);
                xslt1.Transform(xpathdoc.CreateNavigator(), null, resultXml);

                var xmlSource = XmlReader.Create(new StringReader(result.ToString()));
                var result2 = new StringWriter();
                var resultXml2 = XmlWriter.Create(result2, new XmlWriterSettings { ConformanceLevel = ConformanceLevel.Fragment });
                xslt2.Transform(xmlSource, null, resultXml2);

                var text = result2.ToString().Trim('$', ' ');
                if (text.EndsWith(@"\right\")) text = text.Substring(0, text.Length - 1) + ".";

                c.Mark.Value.Metadata = "math";
                c.Mark.Value.Value = text;
            }

            code = "S<:tc !S>:tc S<:oMathPara | math:* S>:oMathPara";
            foreach (var c in Find(code)) {
                c.Mark.Previous.Remove();
                c.Mark.Next.Remove();
            }

            code = "S<:oMathPara | math:* S>:oMathPara";
            foreach (var c in Find(code)) {
                c.Mark.Value.Metadata = "math_para";
                c.Mark.Previous.Remove();
                c.Mark.Next.Remove();
            }
        }

        private void ConvertHyperlinks() {
            var code = "S<:hyperlink !S<:r Sa:id | S=:* ! S>:hyperlink";
            foreach (var c in Find(code)) {
                c.Start.Value.Metadata = "url_ref";
                c.Start.Value.Value = c.Mark.Value.Value;
                c.Start.Next.RemoveTo(c.End);
            }

            ReplaceSequence(DocumentTokens.First, "S<:hyperlink ! S>:hyperlink", "");

            code = "footnote:* !end_footnote:* | url_ref:*";
            foreach (var c in Find(code)) {
                c.Mark.Value.Metadata = "url_ref_footnote";
            }
        }

        private void CondenseTables() {
            ReplaceSequence(DocumentTokens.First, "S<:tblPr ! S>:tblPr", "");
            ReplaceSequence(DocumentTokens.First, "S<:gridCol ! S>:gridCol", "table_def_col:\0");
            ReplaceSequence(DocumentTokens.First, "S<:tcPr ! S>:tcPr", "");
            ReplaceSequence(DocumentTokens.First, "S<:trPr ! S>:trPr", "");
        }

        private void ConvertTables() {
            ReplaceSequence(DocumentTokens.First, "S<:tbl", "table:");
            ReplaceSequence(DocumentTokens.First, "S<:tblGrid", "");
            ReplaceSequence(DocumentTokens.First, "S>:tblGrid", "");
            ReplaceSequence(DocumentTokens.First, "S<:tr para_style:*", "");
            ReplaceSequence(DocumentTokens.First, "S<:tr", "");
            ReplaceSequence(DocumentTokens.First, "S<:tc para_style:*", "");
            ReplaceSequence(DocumentTokens.First, "S<:tc", "");
            ReplaceSequence(DocumentTokens.First, "eop:\0 S>:tc", "end_table_col:");
            ReplaceSequence(DocumentTokens.First, "end_table_col:* S>:tr", "end_table_row:");
            ReplaceSequence(DocumentTokens.First, "S>:tbl", "end_table:");

            var code = "table:* table_def_col:*";
            for (var c = FindSequence(DocumentTokens.First, code); c.IsMatch; c = FindSequence(c.Start, code)) {
                c.Start.Value.Value = (c.Start.Value.Value ?? "|").ToLowerInvariant() + "L|";
                c.End.Remove();
            }
        }

        private void MarkTableFirstRow() {
            var code = "table:* !end_table_row_first:* end_table_row:*";
            foreach (var c in Find(code)) {
                c.End.Value.Metadata = "end_table_row_first";
            }
        }

        private HashSet<string> UsedUnreferencedNames = new HashSet<string>();

        private void ConvertCaptions() {
            // Named (referenced) floats
            var code = "para_style:Caption | bookmark_start:* ! seq:* bookmark_end:* ! eop:*";
            foreach (var c in Find(code)) {
                var label = c.Mark.Value.Value;
                var c2 = FindSequence(c.Start, "seq:* bookmark_end:" + label);
                if (!c2.IsMatch) continue;
                var type = c2.Start.Value.Value.ToLower();

                c.Start.Value.Metadata = "float";
                c.Start.Value.Value = type;
                c.Start.Next.RemoveTo(c2.End);

                if (c.Start.Next.Value.Metadata == null) {
                    c.Start.Next.Value.Value = c.Start.Next.Value.Value.TrimStart(' ', '-', '.', ':', '‒', '–', '—');
                }

                c.End.Value.Metadata = "end_caption";
                c.List.AddAfter(c.End, new Token("label", label));
            }

            // Anonymous (unreferenced) floats
            code = "para_style:Caption ! | seq:* ! eop:*";
            foreach (var c in Find(code)) {
                var type = c.Mark.Value.Value.ToLower();

                c.Start.Value.Metadata = "float";
                c.Start.Value.Value = type;
                c.Start.Next.RemoveTo(c.Mark);

                if (c.Start.Next.Value.Metadata == null) {
                    c.Start.Next.Value.Value = c.Start.Next.Value.Value.TrimStart(' ', '-', '.', ':', '‒', '–', '—');
                }

                c.End.Value.Metadata = "end_caption";
                var niceName = NiceNameFromCaption(c.Start.Next);
                if (string.IsNullOrWhiteSpace(niceName)) {
                    niceName = Guid.NewGuid().ToString("N");
                }
                if (UsedUnreferencedNames.Contains(niceName)) {
                    Warnings.WriteLine("Caption {0} occurs multiple times.", niceName);
                    for (int i = 1; i < int.MaxValue; ++i) {
                        var newName = niceName + i.ToString();
                        if (!UsedUnreferencedNames.Contains(newName)) {
                            niceName = newName;
                            break;
                        }
                    }
                }
                UsedUnreferencedNames.Add(niceName);
                c.List.AddAfter(c.End, new Token("label", niceName));
            }
        }

        private static readonly HashSet<string> BoringWords = new HashSet<string>(new[] {
            "the", "a", "an", "and", "in", "to", "for", "as", "of", "show", "shows", "shown"
        });

        private string NiceNameFromCaption(LinkedListNode<Token> caption, string possibility = null) {
            if (string.IsNullOrWhiteSpace(possibility)) {
                var name_parts = new List<string>();
                for (var node = caption;
                    node != null && node.Value != null && node.Value.Metadata != "end_caption";
                    node = node.Next) {
                    if (node.Value.Metadata != null) continue;
                    var words = node.Value.Value.ToLower();
                    words = Regex.Replace(words, "[^a-z0-9]", " ");
                    foreach (var word in words.Split()) {
                        if (string.IsNullOrWhiteSpace(word) || BoringWords.Contains(word)) continue;
                        name_parts.Add(word.Substring(0, 1).ToUpper() + word.Substring(1));
                    }
                }
                possibility = string.Join("", name_parts);
            }

            if (string.IsNullOrWhiteSpace(possibility))
                return null;

            if (NiceReferenceNames.ContainsValue(possibility)) {
                int alt_count = 1;
                var alt = possibility + alt_count.ToString();
                while (NiceReferenceNames.ContainsValue(alt)) {
                    alt_count += 1;
                    alt = possibility + alt_count.ToString();
                }
                possibility = alt;
            }

            return possibility;
        }

        private void ConvertFloats() {
            // Floats based on tables
            var code = "float:* ! | label:* table:* ! end_table:*";
            foreach (var c in Find(code)) {
                var niceRef = NiceNameFromCaption(c.Start);
                if (!string.IsNullOrWhiteSpace(niceRef))
                    NiceReferenceNames[c.Mark.Value.Value] = niceRef;

                c.List.AddAfter(c.End, new Token("end_float", c.Start.Value.Value));
            }

            // Floats based on images
            code = "float:* ! | label:* image:* eop:*";
            foreach (var c in Find(code)) {
                var niceRef = NiceNameFromCaption(c.Start, c.Mark.Next.Value.Value);
                if (!string.IsNullOrWhiteSpace(niceRef))
                    NiceReferenceNames[c.Mark.Value.Value] = niceRef;

                c.End.Value.Metadata = "end_float";
                c.End.Value.Value = c.Start.Value.Value;
            }

            // Floats based on equation
            code = "float:* ! label:* | math_para:* eop:*";
            foreach (var c in Find(code)) {
                var niceRef = NiceNameFromCaption(c.Start);
                if (!string.IsNullOrWhiteSpace(niceRef))
                    NiceReferenceNames[c.Mark.Previous.Value.Value] = niceRef;

                c.Mark.Value.Metadata = "math_float";
                c.End.Value.Metadata = "end_float";
                c.End.Value.Value = c.Start.Value.Value;
            }

            // Floats based on paragraph style
            code = "float:* ! label:* | para_style:*";
            for (var c = FindSequence(DocumentTokens.First, code); c.IsMatch; c = FindSequence(c.Start.Next, code)) {
                var niceRef = NiceNameFromCaption(c.Start);
                if (!string.IsNullOrWhiteSpace(niceRef))
                    NiceReferenceNames[c.Mark.Previous.Value.Value] = niceRef;

                var style = c.Mark.Value;
                var c2 = c;
                for (c2 = FindSequence(c2.End, "eop:*"); c2.IsMatch; c2 = FindSequence(c2.End.Next, "eop:*")) {
                    if (c2.End.Next == null || !c2.End.Next.Value.Equals(style)) break;
                }
                if (c2.IsMatch) c.End = c2.End;
                c.List.AddAfter(c.End, new Token("end_float", c.Start.Value.Value));
            }
        }

        private static readonly Dictionary<string, string> KnownListingLanguages = new Dictionary<string, string> { 
            { "python", "python" },
            { "esdl", "esdl" }, 
            { "ruby", "ruby" }, 
            { "c++", "cpp" },
            { "pseudocode", "pseudocode" }
        };

        private void DetectCodeListings() {
            var code = "float:listing ! | end_caption:* !float:* end_float:listing";
            foreach (var c in Find(code)) {
                var caption = "";
                for (var node = c.Start.Next; node != c.Mark; node = node.Next) {
                    if (node.Value == null || node.Value.Metadata != null) {
                        continue;
                    }
                    caption += node.Value.Value;
                }

                caption = caption.Replace(',', ' ').Replace('.', ' ') + " ";
                var language = KnownListingLanguages
                    .SelectMany(p => " ,.)".Select(x => new KeyValuePair<string, string>(p.Key + x, p.Value)))
                    .Select(p => new { Value=p.Value, Index=caption.IndexOf(p.Key, StringComparison.CurrentCultureIgnoreCase) })
                    .Where(p => p.Index >= 0)
                    .OrderBy(p => p.Index)
                    .Select(p => p.Value)
                    .FirstOrDefault() ?? "unknownlanguage";

                c.Start.Value.Value = c.End.Value.Value = "listing_" + language;
            }
        }

        private void UseNiceReferenceNames() {
            for (var node = DocumentTokens.First; node != null; node = node.Next) {
                if (node == null || node.Value == null ||
                    node.Value.Metadata == null || node.Value.Value == null)
                    continue;

                if (node.Value.Metadata.Equals("ref", StringComparison.InvariantCultureIgnoreCase) ||
                    node.Value.Metadata.Equals("label", StringComparison.InvariantCultureIgnoreCase)) {
                    string niceRef;
                    if (NiceReferenceNames.TryGetValue(node.Value.Value.TrimStart('_'), out niceRef) ||
                        NiceReferenceNames.TryGetValue(node.Value.Value, out niceRef) ||
                        NiceReferenceNames.TryGetValue("_" + node.Value.Value, out niceRef)) {
                        node.Value.Value = niceRef + "_" + node.Value.Value.Trim(" _-".ToCharArray());
                    }
                }
            }
        }

        private void TidyDocumentStartEnd() {
            ReplaceSequence(DocumentTokens.First, "S<:document ! S<:body", "document:", 1);
            ReplaceSequence(DocumentTokens.First, "S>:body S>:document", "end:document", 1);

            var parts = new[] { "Title", "Author" };

            foreach (var part in parts) {
                var code = "para_style:" + part + " | \0:* eop:*";
                foreach (var c in Find(code)) {
                    c.Start.Remove();
                    c.End.Remove();
                    c.Mark.Value.Metadata = part.ToLower();
                }
            }
        }

        private void WrapDottedNames() {
            var dotsInWords = new Regex(@"(\p{Ll}|\p{Lu})\.(\p{Ll}|\p{Lu})");
            
            for (var node = DocumentTokens.First; node != null; node = node.Next) {
                if (node == null || node.Value == null ||
                    node.Value.Metadata != null || node.Value.Value == null)
                    continue;

                node.Value.Value = dotsInWords.Replace(node.Value.Value, WrapDottedNames_Match);
            }
        }

        public static string WrapDottedNames_Match(Match m) {
            //return m.Groups[1].Value + ".\u200b" + m.Groups[2].Value;
            return m.Groups[1].Value + ".\u00ad" + m.Groups[2].Value;
        }

    }
}
