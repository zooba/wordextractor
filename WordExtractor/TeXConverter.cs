using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace WordExtractor
{
    public class TeXConverter
    {
        private IList<Token> Tokens;

        public IDictionary<string, TextWriter> OutputFiles;

        public const string DefaultPreambleKey = "preamble.tex";
        public string PreambleKey { get; set; }
        public const string DefaultDocumentKey = "document.tex";
        public string DocumentKey { get; set; }
        public const string DefaultBibliographyKey = "bibliography.bib";
        public string BibliographyKey { get; set; }

        public DirectoryInfo Destination { get; set; }

        private bool AsChapter;
        private bool ForXetex;

        private IDictionary<string, Func<string, Token>> Conversions;
        private List<Tuple<string, string>> TextSubstitutions;
        private List<Tuple<string, string>> MathSubstitutions;

        public TeXConverter(IEnumerable<Token> source, bool asChapter, bool forXetex) {
            Tokens = source.ToList();
            AsChapter = asChapter;
            ForXetex = forXetex;
            HeadingOffsetLevel = asChapter ? 0 : 1;

            OutputFiles = new Dictionary<string, TextWriter>();

            PreambleKey = DefaultPreambleKey;
            DocumentKey = DefaultDocumentKey;
            BibliographyKey = DefaultBibliographyKey;

            Destination = null;

            // Initialise conversions dictionary
            Conversions = new Dictionary<string, Func<string, Token>>(DefaultConversions);
            Conversions["para_style"] = ConvertParagraphStyle;
            Conversions["eop"] = ConvertEndOfParagraph;
            if (AsChapter) {
                Conversions["document"] = (_ => null);
                Conversions["bibliography"] = (_ => null);
                Conversions["end"] = (text => text == "document" ? null : new Token("\\end{" + text + "}\r\n"));
            }

            // Initialise substitutions lists
            TextSubstitutions = ReadSubstitutions(Properties.Resources.TextSubstitutions);
            MathSubstitutions = ReadSubstitutions(Properties.Resources.MathSubstitutions);
            if (!ForXetex) {
                TextSubstitutions.AddRange(ReadSubstitutions(Properties.Resources.UnicodeSubstitutions));
                MathSubstitutions.AddRange(ReadSubstitutions(Properties.Resources.MathUnicodeSubstitutions));
            }
        }

        private List<Tuple<string, string>> ReadSubstitutions(string source) {
            var result = new List<Tuple<string, string>>();
            var reader = new System.IO.StringReader(source);
            for (var line = reader.ReadLine(); line != null; line = reader.ReadLine()) {
                var parts = line.Split('\t');
                var first = parts[0];
                if (first.StartsWith("0x")) first = char.ConvertFromUtf32(Convert.ToInt32(first.Substring(2), 16));
                var second = parts[1];
                result.Add(new Tuple<string, string>(first, second));
            }
            return result;
        }

        /// <summary>
        /// Set to zero to use "chapter" for level 1 headings.
        /// Set to one to use "section" for level 1 headings.
        /// </summary>
        public int HeadingOffsetLevel { get; set; }

        private bool InVerbatim;

        public void Run(TextWriter errors = null) {
            if (errors == null) errors = new StringWriter();

            TextWriter preamble = null, document = null, bibliography = null;
            if (!OutputFiles.TryGetValue(DocumentKey, out document)) OutputFiles[DocumentKey] = document = new StringWriter();
            if (AsChapter) {
                OutputFiles.Remove(PreambleKey);
                OutputFiles.Remove(BibliographyKey);
            } else {
                if (!OutputFiles.TryGetValue(PreambleKey, out preamble))
                    OutputFiles[PreambleKey] = preamble = new StringWriter();
                if (!OutputFiles.TryGetValue(BibliographyKey, out bibliography))
                    OutputFiles[BibliographyKey] = bibliography = new StringWriter();
            }

            if (preamble != null) {
                preamble.Write(Properties.Resources.LaTeXPreamble);
                document.WriteLine("\\input{" + PreambleKey + "}");
                document.WriteLine();
            }

            var target = document;
            StringWriter nextFloat = null;
            string inFloat = null;
            bool inAppendices = false;
            var knownFloats = new List<string>();
            InVerbatim = false;
            knownFloats.Add("equation");
            knownFloats.Add("figure");
            knownFloats.Add("listing");

            foreach (var t in DoConvert(Tokens)) {
                if (t.Metadata == null) {
                    var text = t.Value;
                    if (nextFloat != null) nextFloat.Write(text);
                    else target.Write(text);
                } else if (t.Metadata.Equals("error", StringComparison.InvariantCultureIgnoreCase)) {
                    if (errors != null) errors.WriteLine("Unexpected token: " + t.Value);
                } else if (t.Metadata.Equals("preamble", StringComparison.InvariantCultureIgnoreCase)) {
                    if (preamble != null) preamble.Write(t.Value);
                } else if (t.Metadata.StartsWith("float_", StringComparison.InvariantCultureIgnoreCase)) {
                    if (nextFloat != null) errors.WriteLine("Unexpected float within float: " + t.Value);
                    else if (inFloat != null) errors.WriteLine("Unterminated float: " + inFloat);
                    nextFloat = null;
                    inFloat = null;
                    target = document;

                    nextFloat = new StringWriter();
                    nextFloat.Write(t.Value);
                    inFloat = t.Metadata.Substring(6);

                    if (knownFloats.Contains(inFloat) == false) {
                        knownFloats.Add(inFloat);
                        if (preamble != null) {
                            preamble.WriteLine();
                            preamble.WriteLine("\\newcommand\\wxbegin" + inFloat + "[2]{\\begin{" + inFloat + "}\\caption{#1}\\label{#2}\\centering}");
                            preamble.WriteLine("\\newcommand\\wxend" + inFloat + "{\\end{" + inFloat + "}}");
                            preamble.WriteLine("\\crefname{{{0}}}{{{1}}}{{{0}s}}", inFloat, inFloat.Substring(0, 1).ToUpper() + inFloat.Substring(1).ToLower());
                        }
                    }
                } else if (t.Metadata.Equals("label", StringComparison.InvariantCultureIgnoreCase)) {
                    if (nextFloat != null) {
                        inFloat += "_" + t.Value + ".tex";

                        if (!OutputFiles.TryGetValue(inFloat, out target)) {
                            OutputFiles[inFloat] = target = new StringWriter();
                        }
                        document.WriteLine("\r\n% " + nextFloat.ToString().Replace("\r\n", "//") + "\r\n\\input{" + inFloat + "}");

                        target.Write(nextFloat.ToString());
                        target.Write("{" + t.Value + "}");
                        if (inFloat.StartsWith("listing", StringComparison.InvariantCultureIgnoreCase)) {
                            InVerbatim = true;
                            target.Write("\r\n");
                        } else if (inFloat.StartsWith("equation", StringComparison.InvariantCultureIgnoreCase)) { } else {
                            target.Write("\r\n\r\n");
                        }
                        nextFloat = null;
                    } else if (InVerbatim) {
                        if (ForXetex) {
                            target.Write("¶\\label{" + t.Value + "}¶");
                        } else {
                            target.Write("@\\label{" + t.Value + "}@");
                        }
                    } else {
                        target.Write("\\label{" + t.Value + "}");
                    }
                } else if (t.Metadata.Equals("end_float", StringComparison.InvariantCultureIgnoreCase)) {
                    if (nextFloat != null) errors.WriteLine("Float terminated early: " + nextFloat.ToString());
                    else if (inFloat == null) errors.WriteLine("Unexpected float terminator");
                    else target.Write(t.Value);

                    nextFloat = null;
                    inFloat = null;
                    InVerbatim = false;
                    target = document;
                } else if (t.Metadata.Equals("appendix", StringComparison.InvariantCultureIgnoreCase)) {
                    if (!inAppendices && !AsChapter) {
                        target.WriteLine("\\appendix");
                        inAppendices = true;
                    }
                    target.Write(t.Value);
                }
            }
        }

        private IEnumerable<Token> DoConvert(IEnumerable<Token> source) {
            int skipEop = 0;
            foreach (var t in source) {
                if (t.Metadata == "eop" && skipEop > 0) {
                    --skipEop;
                    continue;
                }

                var result = new Token(t);

                if (t.Metadata == null) {
                    if (t.Value == null) continue;
                    if (!InVerbatim) foreach (var p in TextSubstitutions) result.Value = result.Value.Replace(p.Item1, p.Item2);
                } else if (Conversions.ContainsKey(t.Metadata)) result = Conversions[t.Metadata](t.Value);
                else result = new Token("error", t.ToString());

                if (t.Metadata != null && t.Metadata.StartsWith("math", StringComparison.InvariantCultureIgnoreCase)) {
                    if (result == null || result.Value == null) continue;
                    foreach (var p in MathSubstitutions) result.Value = result.Value.Replace(p.Item1, p.Item2);
                }

                if (result == null) {
                    if (t != null && t.Metadata != null && !Conversions.ContainsKey(t.Metadata)) {
                        Console.WriteLine("Unhandled token: " + t.Metadata);
                    }
                    continue;
                }

                if (!string.IsNullOrEmpty(result.Value) && result.Value.EndsWith("\\wxnobreak")) {
                    skipEop += 1;
                    result.Value = result.Value.Remove(result.Value.Length - 10);
                }

                yield return result;
            }
        }

        private static IDictionary<string, Func<string, Token>> DefaultConversions = new Dictionary<string, Func<string, Token>>
        {
            { "document", _ => new Token("\\begin{document}\r\n\\maketitle\r\n\r\n") },
            { "begin", text => new Token("\\begin{" + text + "}\r\n") },
            { "end", text => new Token("\\end{" + text + "}\r\n") },
            { "wxbegin", text => new Token("\\wxbegin" + text + "\r\n") },
            { "wxend", text => new Token("\\wxend" + text + "\r\n") },
            
            { "title", text => new Token("preamble", "\\title{" + text + "}\r\n") },
            { "author", text => new Token("preamble", "\\author{" + text + "}\r\n") },
            
            { "ref", text => new Token("\\cref{" + text.Trim('_', ' ', '.') + "}") },
            { "Ref", text => new Token("\\Cref{" + text.Trim('_', ' ', '.') + "}") },
            { "ref_number", text => new Token("\\ref{" + text.Trim('_', ' ', '.') + "}") },
            { "label", text => new Token("label", text.Trim('_', ' ', '.')) },
            
            { "pagebreak", _ => new Token("\\pagebreak\r\n") },
            { "sectionbreak", _ => new Token("\\clearpage\\newpage\r\n") },
            { "oddsectionbreak", _ => new Token("\\cleardoublepage\\newpage\r\n") },
            { "linebreak", _ => new Token("\r\n") },
            { "hspace", _ => new Token("\\hspace{5mm}") },
            
            { "run_style", ConvertRunStyle },
            { "end_run_style", ConvertEndRunStyle },
            
            { "citation", text => new Token("\\cite{" + text + "}") },
            { "hyperlink", text => new Token("\\url{" + text + "}") },
            
            { "footnote", _ => new Token("\\footnote{") },
            { "end_footnote", _ => new Token("}") },

            { "float", ConvertFloat },
            { "end_caption", _ => new Token("}") },
            { "end_float", ConvertEndFloat },

            { "list", ConvertList },
            { "end_list", ConvertEndList },
            { "item", _ => new Token("\\item ") },

            { "math_float", text => new Token(text) },
            { "math", text => new Token("$" + text + "$") },
            { "math_para", ConvertMathPara },

            { "table", text => new Token("\\begin{tabulary}{\\columnwidth}{" + text + "}\r\n\\hline\r\n") },
            { "end_table", _ => new Token("\\end{tabulary}\r\n") },
            { "end_table_col", _ => new Token(" & ") },
            { "end_table_row", _ => new Token(" \\\\ \\hline\r\n") },

            { "image", text => new Token("\\includegraphics{" + (text ?? "") + "}") },

            { "bibliography", _ => new Token("\\bibliography{bibliography}\r\n") }
        };

        private static string[] IgnoreParagraphStyles = new[] { "BodyText" };
        private static string[] HeadingCommands = new[] { "\\chapter{", "\\section{", "\\subsection{", "\\subsubsection{" };

        private int ParagraphStyleCount;

        private Token ConvertParagraphStyle(string text) {
            if (IgnoreParagraphStyles.Contains(text)) return null;

            ParagraphStyleCount += 1;

            var match = Regex.Match(text, "^Heading([0-9]+)$");
            if (match != null) {
                int i = 0;
                if (int.TryParse(match.Groups[1].Value, out i)) {
                    i += HeadingOffsetLevel - 1;
                    if (i < 4) return new Token(HeadingCommands[i]);
                    return new Token(text.ToLowerInvariant());
                }
            }

            if (text.Equals("appendix", StringComparison.CurrentCultureIgnoreCase)) {
                return new Token("appendix", HeadingCommands[HeadingOffsetLevel]);
            }

            ParagraphStyleCount -= 1;
            return null;
        }

        private Token ConvertEndOfParagraph(string text) {
            var nl = InVerbatim ? "\r\n" : "\r\n\r\n";
            if (ParagraphStyleCount > 0) { ParagraphStyleCount -= 1; return new Token("}" + nl); }
            return new Token(nl);
        }

        private static Token ConvertRunStyle(string text) {
            var style = "";
            if (text.IndexOf("monospace", StringComparison.CurrentCultureIgnoreCase) >= 0) style += "\\texttt{";
            if (text.IndexOf("strong", StringComparison.CurrentCultureIgnoreCase) >= 0) style += "\\textbf{";
            if (text.IndexOf("emphasis", StringComparison.CurrentCultureIgnoreCase) >= 0) style += "\\textit{";
            if (text.IndexOf("small", StringComparison.CurrentCultureIgnoreCase) >= 0) style += "\\textsc{";
            if (text.IndexOf("superscript", StringComparison.CurrentCultureIgnoreCase) >= 0) style += "\\textsuperscript{";
            if (text.IndexOf("subscript", StringComparison.CurrentCultureIgnoreCase) >= 0) style += "\\textsubscript{";
            if (text.IndexOf("notes", StringComparison.CurrentCultureIgnoreCase) >= 0) style += "\\wxnotes{";
            if (string.IsNullOrEmpty(style)) return null;
            return new Token(style);
        }

        private static Token ConvertEndRunStyle(string text) {
            var style = ConvertRunStyle(text);
            if (style == null) return null;
            int styles = style.Value.Count(c => c == '{');
            return new Token(new String('}', styles));
        }

        private static Token ConvertList(string text) {
            if (text.IndexOf("bullet", StringComparison.InvariantCultureIgnoreCase) >= 0) return new Token("\\begin{itemize}\r\n");
            if (text.IndexOf("decimal", StringComparison.InvariantCultureIgnoreCase) >= 0) return new Token("\\begin{enumerate}\r\n");
            return new Token((string)null);
        }

        private static Token ConvertEndList(string text) {
            if (text.IndexOf("bullet", StringComparison.InvariantCultureIgnoreCase) >= 0) return new Token("\\end{itemize}\r\n");
            if (text.IndexOf("decimal", StringComparison.InvariantCultureIgnoreCase) >= 0) return new Token("\\end{enumerate}\r\n");
            return new Token((string)null);
        }

        private static Token ConvertMathPara(string text) {
            if (text.EndsWith(".")) {
                text = text.Remove(text.Length - 1);
                return new Token("$$" + text + "$$.");
            }
            return new Token("$$" + text + "$$\r\n\\noindent{}\\wxnobreak");
        }

        private static Token ConvertFloat(string text) {
            if (text.IndexOf("listing", StringComparison.InvariantCultureIgnoreCase) == 0) {
                var language = text.Substring(text.IndexOf('_') + 1).ToLower();
                return new Token("float_listing", "\\wxbeginlisting{" + language + "}{}{");
            } else {
                return new Token("float_" + text, "\\wxbegin" + text + "{");
            }
        }

        private static Token ConvertEndFloat(string text) {
            if (text.IndexOf("listing", StringComparison.InvariantCultureIgnoreCase) == 0) {
                var language = text.Substring(text.IndexOf('_') + 1).ToLower();
                return new Token("end_float", "\\end{" + language + "}\\wxendlisting\r\n");
            } else if (text.Equals("listing", StringComparison.InvariantCultureIgnoreCase)) {
                return new Token("end_float", "\\end{unknownlanguage}\\wxendlisting\r\n");
            } else {
                return new Token("end_float", "\\wxend" + text + "\r\n");
            }
        }
    }
}
