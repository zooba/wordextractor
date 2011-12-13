﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.239
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace WordExtractor.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("WordExtractor.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Processes Word 2007 (.docx) documents and produces LaTeX code.
        ///
        ///&gt;&gt; WordExtractor [/?] [/L#] [/nooutput] [/nolatex] [/out:&lt;path&gt;]
        ///                 [/name:&lt;name&gt;] [/overwrite] SourceFile
        ///
        ///  ?          Displays this usage information..
        ///  L:#        Only applies filters up to the specified level (0-9).
        ///  nooutput   Applies filters but does not produce any output.
        ///  nolatex    Writes the intermediate representation to stdout.
        ///  out:...    Stores generated LaTeX files in the specified
        ///             direc [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string CommandLineUsage {
            get {
                return ResourceManager.GetString("CommandLineUsage", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to %------------------------------------------------------------------------------
        ///
        ///\documentclass[pdftex,a4paper,11pt]{article}
        ///\pdfminorversion=5
        ///\usepackage[lmargin=2cm,rmargin=2cm,tmargin=3cm,bmargin=2cm]{geometry}
        ///
        ///%------------------------------------------------------------------------------
        ///
        ///\usepackage{fancyhdr}
        ///\pagestyle{fancy}
        ///\setlength{\headheight}{14pt}
        ///%\fancyhead[L]{report number}
        ///
        ///%------------------------------------------------------------------------------
        ///
        ///\usepackage{cmap} [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string LaTeXPreamble {
            get {
                return ResourceManager.GetString("LaTeXPreamble", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to ×	\times 
        ///÷	\div 
        ///⋅	\cdot 
        ///0x2002	\:
        ///0x2003	\;
        ///0x2006	\,
        ///&amp;lt;	&lt;
        ///&amp;gt;	&gt;
        ///\mathbit	.
        /// </summary>
        internal static string MathSubstitutions {
            get {
                return ResourceManager.GetString("MathSubstitutions", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to  	~
        ///0x2002	\:
        ///0x2003	\;
        ///0x2006	\,
        ///.
        /// </summary>
        internal static string MathUnicodeSubstitutions {
            get {
                return ResourceManager.GetString("MathUnicodeSubstitutions", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to \	\textbackslash		# Replaced with \textbackslash{} below
        ///`	\textgrave			# Replaced with \`{} below
        ///&amp;	\&amp;
        ///$	\$
        ///#	\#
        ///%	\%
        ///{	\{
        ///}	\}
        ///_	\_
        ///\*	*
        ///&gt;&gt;	&gt;{}&gt;
        ///&lt;&lt;	&lt;{}&lt;
        ///\textbackslash	\textbackslash{}
        ///\textgrave	\`{}.
        /// </summary>
        internal static string TextSubstitutions {
            get {
                return ResourceManager.GetString("TextSubstitutions", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to ‘	`
        ///’	&apos;
        ///“	``
        ///”	&apos;&apos;
        ///©	\textcopyright{}
        ///®	\textregistered{}
        ///™	\texttrademark{}
        ///…	\textellipsis{}
        ///–	\textendash{}
        ///‒	\textendash{}			# Figure dash
        ///—	\textemdash{}
        ///0x200B	{}				# Zero-width space
        ///0x00A0	~				# Non-breaking space
        ///0x00AD	\-				# Soft hyphen
        ///à	\`{a}
        ///á	\&apos;{a}
        ///â	\^{a}
        ///ã	\~{a}
        ///ä	\&quot;{a}
        ///å	\aa{}
        ///æ	\ae{}
        ///ç	\c{c}
        ///è	\`{e}
        ///é	\&apos;{e}
        ///ê	\^{e}
        ///ë	\&quot;{e}
        ///ì	\`{\i}
        ///í	\&apos;{\i}
        ///î	\^{\i}
        ///ï	\&quot;{\i}
        ///ñ	\~{n}
        ///ò	\`{o}
        ///ö	\&quot;{o}
        ///ù	\`{u}
        ///ú	\&apos;{u}
        ///û	\^{u}
        ///ü	\&quot;{u}
        ///ý	\&apos;{y}
        ///ÿ	\&quot;{y}.
        /// </summary>
        internal static string UnicodeSubstitutions {
            get {
                return ResourceManager.GetString("UnicodeSubstitutions", resourceCulture);
            }
        }
    }
}
