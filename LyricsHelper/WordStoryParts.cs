using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace LyricsHelper {
	internal static class WordStoryParts {
		public static readonly XNamespace Transitional = @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
		public static readonly XNamespace Strict = @"http://purl.oclc.org/ooxml/wordprocessingml/main";

		private static readonly HashSet<string> StoryRootNames = new(StringComparer.Ordinal) {
			"document",
			"glossaryDocument",
			"hdr",
			"ftr",
			"footnotes",
			"endnotes",
			"comments"
		};

		public static bool IsWordNamespace(XNamespace ns) {
			return ns == Transitional || ns == Strict;
		}

		/// <summary>
		/// True for every WordprocessingML story root, including headers,
		/// footers, notes, comments and the glossary document.
		/// </summary>
		public static bool IsStoryRoot(XElement? element) {
			return element is not null
				&& IsWordNamespace(element.Name.Namespace)
				&& StoryRootNames.Contains(element.Name.LocalName);
		}

		/// <summary>
		/// True for the parts of a docx/flat-OPC package that can contain
		/// WordprocessingML text. Names are normalised to /word/... form.
		/// </summary>
		public static bool IsStoryPart(string? partName) {
			if (string.IsNullOrEmpty(partName))
				return false;

			string name = NormalizePartName(partName);
			if (!name.StartsWith("word/", StringComparison.Ordinal))
				return false;

			string file = name["word/".Length..];
			if (file.Equals("glossary/document.xml", StringComparison.Ordinal))
				return true;

			if (!file.EndsWith(".xml", StringComparison.Ordinal))
				return false;

			string stem = file[..^4];
			return stem.Equals("document", StringComparison.Ordinal)
				|| stem.Equals("footnotes", StringComparison.Ordinal)
				|| stem.Equals("endnotes", StringComparison.Ordinal)
				|| stem.Equals("comments", StringComparison.Ordinal)
				|| IsNumberedPart(stem, "header")
				|| IsNumberedPart(stem, "footer");
		}

		public static bool IsMainDocumentPart(string? partName) {
			if (string.IsNullOrEmpty(partName))
				return false;
			return string.Equals(NormalizePartName(partName), "word/document.xml", StringComparison.Ordinal);
		}

		public static string NormalizePartName(string partName) {
			string name = partName.Replace('\\', '/').TrimStart('/');
			return name;
		}

		private static bool IsNumberedPart(string stem, string prefix) {
			if (stem.Equals(prefix, StringComparison.Ordinal))
				return true;
			if (stem.Length <= prefix.Length || !stem.StartsWith(prefix, StringComparison.Ordinal))
				return false;
			return stem[prefix.Length..].All(char.IsAsciiDigit);
		}
	}
}
