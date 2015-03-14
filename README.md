Word Extractor allows you to write, review and edit your documents using Microsoft Word 2007 (or later), complete with spelling and grammar checking, trackable changes and navigation features, before using LaTeX to perform typesetting.

Only a subset of Word's formatting features are carried through to the LaTeX document - the emphasis is on content - but most labels and cross-references are retained.

The application is currently "me"-ware, as in, it performs the minimum required functionality for the purposes of the author. It is published here to encourage others to use and improve.

Prerequisites for this program include:

  * [.NET Framework 4.0](http://go.microsoft.com/fwlink/?LinkID=186913)
  * [Microsoft Word 2007 or later](http://office.microsoft.com/en-au/word/)
  * A LaTeX distribution, such as [MiKTeX](http://miktex.org/)

#  Command-Line Usage 

The usage information shown when executing the program is:

```
Processes Word 2007 (.docx) documents and produces LaTeX code.

>> WordExtractor [/?] [/L#] [/nooutput] [/nolatex] [/out:<path>]
                 [/name:<name>] [/overwrite] SourceFile

  ?          Displays this usage information..
  L:#        Only applies filters up to the specified level (0-9).
  nooutput   Applies filters but does not produce any output.
  nolatex    Writes the intermediate representation to stdout.
  out:...    Stores generated LaTeX files in the specified
             directory.
  name:...   Specifies the filename of the main LaTeX file.
  overwrite  Recreates all float and preamble files.
  SourceFile Path to a .docx document to process.
```

This page elaborates on these items and provides some examples of common usage.

#  Extracting a Document 

The simplest way to extract a document is to simply call:

```
WordExtractor.exe "My Document.docx"
```

This will create a folder called `out` containing the extracted files. The main LaTeX source file is called `document.tex`.

Running this command a second time will display messages warning that not every file was overwritten. This allows custom modifications made to the LaTeX (discussed below) to be retained when re-extracting content changes.

To overwrite all files, include the `/overwrite` option:

```
WordExtractor.exe /overwrite "My Document.docx"
```

To specify the name of the folder containing the extracted files, use the `/out` option:

```
WordExtractor.exe /out:"LateX Files" "My Document.docx"
```

To specify the name of the main LaTeX file (`document.tex` by default), use the `/name` option:

```
WordExtractor.exe /name:doc "My Document.docx"
```

##  Automating a Document Build 

Assuming `WordExtractor.exe`, `pdflatex.exe` and `bibtex.exe` are on the path, the following batch-file simplifies the extract and build process:

```
@echo off
WordExtractor "%1"
pushd out
pdflatex -c-style-errors document
bibtex document
pdflatex -c-style-errors document > nul
pdflatex -c-style-errors document
popd
copy out\document.pdf %~dpn1.pdf

pause
```

#  The Extracted Files 

The main file is named `document.tex` or the name provided with the `/name` option. This is the file passed to `pdflatex` and `bibtex` calls, and is also the only file that is always overwritten. In general, `document.tex` only contains textual content; everything else is extracted into a separate file.

The `preamble.tex` file includes package references, some definitions and the title and author names. These last details are taken from the parts of the original document having the _Title_ and _Author_ styles, respectively. Since the `preamble.tex` file is not overwritten by subsequent extractions, it may be used to tweak style definitions as desired.

The `bibliography.bib` file is always empty in the current implementation. It should be filled with BibTeX entries for all the bibliography items, ensuring that the keys used in these items match the tag names used in Word. The `bibliography.bib` file is not overwritten by subsequent extractions.

All other `*.tex` files are floats. These are named using the internal bookmarks set by Word, prefixed by the type (for example, "figure" or "table"). These are included into `document.tex` by `\input{}` statements, and should not be renamed. The contents of each float, however, may be modified to change (or correct) the display style. Floats are not overwritten by subsequent extractions.

##  Floats 

Floats are introduced by text using the _Caption_ style, followed by a bookmarked sequence number. This is the style automatically used when a caption is inserted within Word. (If a float is unreferenced, Word does not create a bookmark. These floats are still extracted, and are assigned an arbitrary reference, which may change between extractions.)

The caption must appear first. Displaying captions below elements may be specified later (as part of `preamble.tex`) but for correct detection in the Word document, captions must be specified first. This includes numbered equations, which are labeled `Equation 1: <caption text> <newline> <display equation>` in Word, but are shown as `<equation> <equation number>` when converted.

Floats may be based on paragraphs, a table, an image or an equation.

Where the first element after the caption is a paragraph, the float continues as long as the paragraph style does not change. This may be used to specify, for example, code listings using a different paragraph style. A blank paragraph in the _Normal_ or _Body Text_ style is recommended after the last paragraph to ensure correct extraction.

Where the first element after the caption is a table, the entire table is extracted as the float. Basic conversion to a LaTeX style table is performed, however, in most cases manual styling (in the extracted `table_....tex` file) is required.

Where the first element after the caption is an image or diagram, a standard image float is created. This float requires specification of a source file (generally a PDF) and a display width - the image is not extracted (in the current version).

Where the first element after the caption is an equation (Word 2007/10 style, not a MathType object), it will be converted into a LaTeX display equation. (Inline equations are also converted, but are kept in `document.tex`.)

#  Bugs 

Everywhere. However, most of them can be relatively easily corrected (such as floats not being closed).

The extractor will display any tags that it does not recognise. These can often be ignored, and are mainly of interest to developers. (They are displayed using an internal format, but normally equate directly to the OOXML tags.)

A number of special characters are converted manually based on substitution lists, since LaTeX only supports American text in source files and Word uses Unicode internally. Not all characters are converted yet - feel free to submit an issue for any characters that are missing. This also applies to many mathematical symbols.

This document is probably missing millions of details. Again, start a discussion or post an issue if something needs to be clarified.
