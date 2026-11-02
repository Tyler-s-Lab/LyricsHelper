using System.Xml.Linq;
using static LyricsHelper.CharHelper;

namespace LyricsHelper {
	internal static class RubyCleanerForJapanese {

		internal static async Task<string> TryProcess(string[] paths) {
			string err = "";
			try {
				var res = await OfficeWordDocProc.ModifyFilesAsync(paths, ProcessXml);
				err = TextUtil.JoinCapped(res, "\n", DocumentLimits.MaxResultChars);
			}
			catch (Exception ex) {
				err = ex.Message;
			}
			return err;
		}

		/// <summary>
		/// Rewrites only the w:ruby nodes of one story part. Everything else in
		/// the surrounding run/paragraph is left untouched, so the operation
		/// cannot drop bookmarks, breaks or other run children.
		/// </summary>
		static bool ProcessXml(XElement root) {
			if (!WordStoryParts.IsStoryRoot(root))
				throw new Exception("XML not including a Word story root");

			XNamespace w = root.Name.Namespace;
			var rubies = root.Descendants(w + "ruby").ToList();
			long rebuildBudget = DocumentLimits.MaxRebuildNodes;
			bool modified = false;

			foreach (XElement ruby in rubies) {
				if (ruby.Parent is null)
					continue;

				if (ruby.Parent is not XElement parentRun)
					continue;
				if (parentRun.Name != w + "r")
					continue;

				XElement? rubyTextContainer = ruby.Element(w + "rt");
				XElement? rubyBaseContainer = ruby.Element(w + "rubyBase");
				string? rubyText = GetContainerText(rubyTextContainer, w);
				string? rubyBase = GetContainerText(rubyBaseContainer, w);
				if (string.IsNullOrEmpty(rubyBase))
					continue;

				// Fully-katakana bases do not need furigana at all. The same
				// applies when the base is English/digits or a symbol, where
				// the ruby is usually a translation or a description rather
				// than a reading.
				if (RubyTextPolicy.IsSpecialBase(rubyBase, rubyText)) {
					ReplaceSpecialRuby(ruby, parentRun, w, rubyBase);
					modified = true;
					continue;
				}

				if (string.IsNullOrEmpty(rubyText))
					continue;

				if (rubyBase.Length > DocumentLimits.MaxRubyBaseLength
					|| rubyText.Length > DocumentLimits.MaxRubyTextLength)
					continue;

				List<XElement> rubyTextRuns = GetPlainRuns(rubyTextContainer, w);
				List<XElement> rubyBaseRuns = GetPlainRuns(rubyBaseContainer, w);
				if (rubyTextRuns.Count == 0 || rubyBaseRuns.Count == 0)
					continue;

				List<RubySegment> segments;
				try {
					var textSegments = SplitJapaneseTextManual(rubyBase);
					if (textSegments.Count < 2)
						continue;
					segments = GetRubySegments(textSegments, rubyText);
				}
				catch {
					// A malformed annotation is left exactly as Word wrote it.
					continue;
				}

				XElement? rubyProperties = ruby.Element(w + "rubyPr");
				XElement? rubyTextProperties = rubyTextRuns[0].Element(w + "rPr");
				XElement? rubyBaseProperties = rubyBaseRuns[0].Element(w + "rPr");

				long preferenceNodes = CountNodes(rubyProperties)
					+ CountNodes(rubyTextProperties)
					+ CountNodes(rubyBaseProperties);
				long rebuildCost = (long)segments.Count * (10 + preferenceNodes);
				if (rebuildCost > rebuildBudget)
					continue;
				rebuildBudget -= rebuildCost;

				// Replace the w:ruby node with run *content*, not with new
				// w:r wrappers: the ruby already lives inside its run and that
				// run keeps its rPr, bookmarks and other children.
				var replacement = new List<XElement>(segments.Count);
				foreach (RubySegment segment in segments) {
					if (segment.Type == TextType.Hiragana || string.IsNullOrEmpty(segment.Ruby)) {
						replacement.Add(CreateText(w, segment.Text));
					}
					else {
						var readingRun = new XElement(w + "r");
						if (rubyTextProperties is not null)
							readingRun.Add(new XElement(rubyTextProperties));
						readingRun.Add(CreateText(w, segment.Ruby));

						var baseRun = new XElement(w + "r");
						if (rubyBaseProperties is not null)
							baseRun.Add(new XElement(rubyBaseProperties));
						baseRun.Add(CreateText(w, segment.Text));

						var newRuby = new XElement(w + "ruby");
						if (rubyProperties is not null)
							newRuby.Add(new XElement(rubyProperties));
						newRuby.Add(new XElement(w + "rt", readingRun));
						newRuby.Add(new XElement(w + "rubyBase", baseRun));
						replacement.Add(newRuby);
					}
				}

				ruby.ReplaceWith(replacement.Cast<object>().ToArray());
				modified = true;
			}

			return modified;
		}

		static XElement CreateText(XNamespace w, string text) {
			return new XElement(w + "t",
				new XAttribute(XNamespace.Xml + "space", "preserve"),
				text);
		}

		/// <summary>
		/// Removes a special-base ruby while keeping the base run properties.
		/// When the surrounding run contains only the ruby and its rPr, the
		/// whole run is rebuilt with the base run's formatting; otherwise the
		/// ruby is replaced in place.
		/// </summary>
		static void ReplaceSpecialRuby(XElement ruby, XElement parentRun, XNamespace w, string baseText) {
			var children = parentRun.Elements().ToList();
			bool canReplaceRun = children.Count(child => child.Name == w + "ruby") == 1
				&& children.All(child => child.Name == w + "rPr" || child.Name == w + "ruby")
				&& parentRun.Nodes().All(node => node is XElement);

			if (!canReplaceRun) {
				ruby.ReplaceWith(CreateText(w, baseText));
				return;
			}

			XElement? baseProperties = ruby.Element(w + "rubyBase")
				?.Element(w + "r")
				?.Element(w + "rPr")
				?? parentRun.Element(w + "rPr");

			var run = new XElement(w + "r");
			foreach (XAttribute attribute in parentRun.Attributes())
				run.Add(new XAttribute(attribute));
			if (baseProperties is not null)
				run.Add(new XElement(baseProperties));
			run.Add(CreateText(w, baseText));

			parentRun.ReplaceWith(run);
		}

		static string? GetContainerText(XElement? container, XNamespace w) {
			if (container is null)
				return null;

			string text = string.Concat(container.Descendants(w + "t").Select(text => text.Value));
			return text.Length == 0 ? null : text;
		}

		static List<XElement> GetPlainRuns(XElement? container, XNamespace w) {
			if (container is null)
				return [];

			var runs = container.Elements(w + "r").ToList();
			foreach (XElement run in runs) {
				if (run.Elements().Any(child => child.Name != w + "rPr" && child.Name != w + "t"))
					return [];
			}

			return runs;
		}

		static int CountNodes(XElement? element) {
			return element is null ? 0 : 1 + element.Descendants().Count();
		}

		static List<RubySegment> GetRubySegments(List<TextSegment> segs, string rubyText) {
			var result = new List<RubySegment>();
			bool lastIsSet = true;
			int i = 0;
			{
				var first = segs.First();
				if (first.Type is TextType.Hiragana) {
					i = 1;
					if (rubyText.StartsWith(first.Text, StringComparison.Ordinal)) {
						result.Add(new RubySegment(first.Text, first.Type, first.Text));
						rubyText = rubyText[first.Text.Length..];
					}
					else {
						result.Add(new RubySegment(first.Text, TextType.Other, ""));
						lastIsSet = false;
					}
				}
			}
			for (int n = ((segs.Last().Type == TextType.Hiragana) ? (segs.Count - 1) : (segs.Count)); i < n; ++i) {
				var seg = segs[i];
				if (seg.Type is not TextType.Hiragana) {
					goto FailedSplitAndMoveIntoRes;
				}
				var i0 = rubyText.IndexOf(seg.Text, StringComparison.Ordinal);
				var i1 = rubyText.LastIndexOf(seg.Text, StringComparison.Ordinal);
				if (i0 == -1 || i0 != i1) {
					goto FailedSplitAndMoveIntoRes;
				}
				if (!lastIsSet) {
					var last = result.Last();
					result.RemoveAt(result.Count - 1);
					last.Ruby = rubyText[..i0];
					rubyText = rubyText[i0..];
					result.Add(last);
				}
				result.Add(new RubySegment(seg.Text, seg.Type, seg.Text));
				rubyText = rubyText[seg.Text.Length..];
				lastIsSet = true;
				continue;
			FailedSplitAndMoveIntoRes:
				{
					RubySegment last;
					if (!lastIsSet) {
						last = result.Last();
						last.Text += seg.Text;
						result.RemoveAt(result.Count - 1);
					}
					else {
						last = new RubySegment(seg.Text, seg.Type, "");
					}
					result.Add(last);
				}
				lastIsSet = false;
			}
			if (segs.Count - 1 == i) { // Is end with hiragana seg.
				var last = segs.Last();
				if (rubyText.EndsWith(last.Text, StringComparison.Ordinal)) {
					var i0 = rubyText.LastIndexOf(last.Text, StringComparison.Ordinal);
					if (!lastIsSet) {
						var last0 = result.Last();
						result.RemoveAt(result.Count - 1);
						last0.Ruby = rubyText[..i0];
						rubyText = rubyText[i0..];
						result.Add(last0);
					}
					result.Add(new RubySegment(last.Text, last.Type, last.Text));
				}
				else if (!lastIsSet) {
					var last0 = result.Last();
					result.RemoveAt(result.Count - 1);
					last0.Ruby = rubyText;
					last0.Text += last.Text;
					result.Add(last0);
				}
				else {
					result.Add(new RubySegment(last.Text, TextType.Other, rubyText));
				}
			}
			else {
				var last0 = result.Last();
				result.RemoveAt(result.Count - 1);
				last0.Ruby = rubyText;
				result.Add(last0);
			}
			return result;
		}

		static List<TextSegment> SplitJapaneseTextManual(string input) {
			var segments = new List<TextSegment>();
			if (string.IsNullOrEmpty(input)) return segments;

			int i = 0;
			while (i < input.Length) {
				TextType currentType = GetTextType(input[i]);
				int start = i;

				while (i < input.Length && GetTextType(input[i]) == currentType) {
					i++;
				}

				string segmentText = input[start..i];
				segments.Add(new TextSegment(segmentText, currentType));
			}

			return segments;
		}

		record TextSegment(string Text, TextType Type);

		struct RubySegment(string _t, TextType _y, string _r) {
			public string Text = _t;
			public TextType Type = _y;
			public string Ruby = _r;
		}
	}
}
