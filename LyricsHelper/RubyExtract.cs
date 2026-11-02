using System.Text;
using System.Xml.Linq;
using static LyricsHelper.CharHelper;

namespace LyricsHelper {
	internal static class RubyExtract {

		internal static async Task<string> TryProcess(string[] paths) {
			string err = "";
			try {
				var res = await OfficeWordDocProc.ReadFilesAsync(paths, ProcessXml);
				err = TextUtil.JoinCapped(res, "\n\n", DocumentLimits.MaxResultChars);
			}
			catch (Exception ex) {
				err = ex.Message;
			}
			return err;
		}

		static string? ProcessXml(XElement root) {
			if (!WordStoryParts.IsStoryRoot(root))
				throw new Exception("XML not including a Word story root");

			XNamespace w = root.Name.Namespace;
			var builder = new StringBuilder();
			bool truncated = false;

			void Add(string? value) {
				if (truncated || string.IsNullOrEmpty(value))
					return;

				int room = DocumentLimits.MaxResultChars - builder.Length;
				if (value.Length <= room) {
					builder.Append(value);
					return;
				}

				if (room > 0)
					builder.Append(value, 0, room);
				builder.Append("\n[output truncated]");
				truncated = true;
			}

			void WalkRuby(XElement ruby, XElement? outerRun) {
				string? rubyText = GetContainerText(ruby.Element(w + "rt"), w);
				string? rubyBase = GetContainerText(ruby.Element(w + "rubyBase"), w);

				// Katakana/Latin/digit/symbol bases have their own output
				// policy: do not treat their (often translated or descriptive)
				// reading as the lyric text.
				if (RubyTextPolicy.TryGetSpecialExtraction(rubyBase, rubyText, out string specialText)) {
					Add(specialText);
					return;
				}

				bool mistakeColor = HasMistakeColor(ruby, outerRun, w);
				bool mistakeReading = rubyText is not null && rubyText.Any(c => !IsHiraganaReadingChar(c));

				if (mistakeColor || mistakeReading) {
					Add("[");
					Add(rubyBase);
					Add("](");
					Add(rubyText);
					Add(")");
					return;
				}

				// If the reading is missing/empty, keep the base text instead of
				// silently dropping the lyric.
				Add(!string.IsNullOrEmpty(rubyText) ? rubyText : rubyBase);
			}

			void WalkRun(XElement run) {
				foreach (XElement child in run.Elements()) {
					if (child.Name == w + "rPr")
						continue;

					if (child.Name == w + "t") {
						Add(child.Value);
					}
					else if (child.Name == w + "ruby") {
						WalkRuby(child, run);
					}
					else if (child.Name == w + "tab") {
						Add("\t");
					}
					else if (child.Name == w + "br" || child.Name == w + "cr") {
						Add("\n");
					}
					else if (child.Name == w + "noBreakHyphen") {
						Add("\u2011");
					}
					else if (child.Name == w + "softHyphen") {
						Add("\u00AD");
					}
					else if (child.Name == w + "sym") {
						Add(GetSymbolText(child, w));
					}
					else if (child.Name == w + "drawing" || child.Name == w + "pict") {
						// Text boxes inside drawings are separate paragraphs and
						// are visited by the paragraph loop below.
						continue;
					}
					else {
						WalkInline(child);
					}
				}
			}

			void WalkInline(XElement parent) {
				foreach (XElement child in parent.Elements()) {
					if (child.Name == w + "p")
						continue; // nested text-box paragraph: processed separately
					if (child.Name == w + "r")
						WalkRun(child);
					else if (child.Name == w + "ruby")
						WalkRuby(child, null);
					else
						WalkInline(child);
				}
			}

			foreach (XElement paragraph in root.Descendants(w + "p")) {
				if (truncated)
					break;

				WalkInline(paragraph);
				Add(Environment.NewLine);
			}

			return builder.ToString();
		}

		static string? GetContainerText(XElement? container, XNamespace w) {
			if (container is null)
				return null;

			string text = string.Concat(container.Descendants(w + "t").Select(text => text.Value));
			return text.Length == 0 ? null : text;
		}

		static bool HasMistakeColor(XElement ruby, XElement? outerRun, XNamespace w) {
			if (HasExplicitNonBlackColor(outerRun?.Element(w + "rPr"), w))
				return true;

			return ruby.Descendants(w + "rPr").Any(runProperties => HasExplicitNonBlackColor(runProperties, w));
		}

		static bool HasExplicitNonBlackColor(XElement? runProperties, XNamespace w) {
			XElement? color = runProperties?.Element(w + "color");
			if (color is null)
				return false;

			string? value = (string?)color.Attribute(w + "val");
			if (string.IsNullOrEmpty(value))
				return false;

			return !value.Equals("auto", StringComparison.OrdinalIgnoreCase)
				&& !value.Equals("000000", StringComparison.OrdinalIgnoreCase);
		}

		static string? GetSymbolText(XElement symbol, XNamespace w) {
			string? value = (string?)symbol.Attribute(w + "char");
			if (string.IsNullOrEmpty(value))
				return null;

			try {
				int codePoint = Convert.ToInt32(value, 16);
				return char.ConvertFromUtf32(codePoint);
			}
			catch {
				return null;
			}
		}
	}
}
