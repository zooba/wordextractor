using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordExtractor
{
    public class Token : IEquatable<Token>
    {
        public string Metadata { get; set; }
        public string Value { get; set; }

        public Token(Token source) {
            Metadata = source.Metadata;
            Value = source.Value;
        }

        public Token(string text) {
            Metadata = null;
            Value = text;
        }

        public Token(string metadata, string text) {
            Metadata = metadata;
            if (string.IsNullOrEmpty(Metadata) || Metadata == "\0") Metadata = null;
            Value = text;
        }

        public static Token Parse(string source) {
            string m = null, v = source;
            int i = v.IndexOf(':');
            if (i >= 0) {
                m = v.Substring(0, i);
                v = v.Substring(i + 1);
            }
            if (string.IsNullOrEmpty(m) || m == "\0") m = null;
            if (string.IsNullOrEmpty(v) || v == "\0") v = null;
            return new Token(m, v);
        }

        public override string ToString() {
            if (Metadata != null) return Metadata + ":" + Value;
            else return Value;
        }

        public override bool Equals(object obj) {
            return Equals(obj as Token);
        }

        public override int GetHashCode() {
            return Metadata.GetHashCode() ^ Value.GetHashCode();
        }

        public bool Equals(Token other) {
            if (object.Equals(other, null)) return false;

            return Metadata == other.Metadata && Value == other.Value;
        }

        public bool WildcardEquals(Token other) {
            if (object.Equals(other, null)) return false;

            bool m = false, v = false;

            if (Metadata == "*" || other.Metadata == "*") m = true;
            else m = (Metadata == other.Metadata);

            if (Value == "*" || other.Value == "*") v = true;
            else v = (Value == other.Value);

            return m && v;
        }
    }
}
