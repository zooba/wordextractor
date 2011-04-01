using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace WordExtractor
{
    class Program
    {
        private static bool ArgsHas(string[] args, string option)
        {
            return args.Any(s => { var m = Regex.Match(s, @"^[/\-]" + option, RegexOptions.IgnoreCase); return m != null && m.Success; });
        }

        private static string ArgsRead(string[] args, string option, string pattern)
        {
            var match = args.Select(s => Regex.Match(s, @"^[/\-]" + option + ":?(" + pattern + ")", RegexOptions.IgnoreCase)).FirstOrDefault(m => m != null && m.Success);
            if (match == null || match.Groups.Count != 2) return null;
            else return match.Groups[1].Value;
        }

        static void Main(string[] args)
        {
            if (ArgsHas(args, "\\?") || args.Length == 0)
            {
                Console.Write(Properties.Resources.CommandLineUsage);
                return;
            }

            var source = args.LastOrDefault();
            if (!File.Exists(source))
            {
                Console.Error.WriteLine("File does not exist: " + source);
                return;
            }

            var levelMatch = ArgsRead(args, "L\\:?", "[0-9]") ?? "9";
            int level;
            if (!int.TryParse(levelMatch, out level)) level = 9;

            Simplifier simplifier;
            try
            {
                var reader = new DocxReader(source);
                simplifier = new Simplifier(reader, Console.Error);
            }
            catch (IOException)
            {
                Console.Error.WriteLine("Unable to open file: " + source);
                return;
            }

            simplifier.RunAll(level);

            if (ArgsHas(args, "nooutput")) return;

            if (ArgsHas(args, "nolatex"))
            {
                Console.OutputEncoding = Encoding.UTF8;
                foreach (var t in simplifier.Document)
                {
                    Console.WriteLine(t);
                }
            }
            else
            {
                var outDirectory = ArgsRead(args, "out", ".+") ?? "out";
                var outName = ArgsRead(args, "name", "[a-zA-Z0-9_]+") ?? "document";
                var overwrite = ArgsHas(args, "over(write)?");

                Console.OutputEncoding = Encoding.ASCII;
                var tex = new TeXConverter(simplifier.Document);

                tex.DocumentKey = outName + ".tex";
                tex.Run(Console.Error);

                if (!Directory.Exists(outDirectory)) Directory.CreateDirectory(outDirectory);
                foreach (var p in tex.OutputFiles)
                {
                    var dest = Path.Combine(outDirectory, p.Key);
                    if (File.Exists(dest) && !overwrite && p.Key != tex.DocumentKey)
                    {
                        Console.Error.WriteLine("Not overwriting: " + dest);
                    }
                    else
                    {
                        using (var file = new StreamWriter(dest, false, Encoding.ASCII))
                        {
                            file.Write(p.Value.ToString());
                        }
                    }
                }
            }
        }
    }
}
