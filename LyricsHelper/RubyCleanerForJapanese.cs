using SharpCompress.Archives;
using SharpCompress.Archives.Zip;
using SharpCompress.Common;
using System.IO;
using System.Xml.Linq;

namespace LyricsHelper {
	internal static class RubyCleanerForJapanese {

		internal static async Task<string> TryProcess(string[] paths) {
			string err = "";
			try {
				var res = await OfficeWordDocProc.ModifyFilesAsync(paths, ProcessXml);
				err = string.Join("\n", res);
			}
			catch (Exception ex) {
				err = ex.Message;
			}
			return err;
		}

		static ValueTuple<XContainer, string> ProcessXml(string path, XContainer xml) {
			XNamespace w = @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
			if (xml.Element(w + "document") is not XElement document) {
				throw new Exception("XML not including w:document");
			}
			if (document.Element(w + "body") is not XElement body) {
				throw new Exception("XML not including w:body");
			}
			var paragraphs = body.Elements(w + "p");

			foreach (var paragraph in paragraphs) {
				var newRuns = paragraph.Elements().SelectMany(element => {
					if (element.Name != w + "r") {
						return [element];
					}

					var run = element;
					if (run.Element(w + "t") is not null) {
						return [element];
					}
					var runPreference = run.Element(w + "rPr");

					if (run.Element(w + "ruby") is not XElement ruby) {
						return [element];
					}
					var rubyPreference = ruby.Element(w + "rubyPr");

					var rubyTextRun = ruby.Element(w + "rt")?.Element(w + "r");
					var rubyTextPr = rubyTextRun?.Element(w + "rPr");

					var rubyBaseRun = ruby.Element(w + "rubyBase")?.Element(w + "r");
					var rubyBasePr = rubyBaseRun?.Element(w + "rPr");

					var rubyText = rubyTextRun?.Element(w + "t")?.Value;
					var rubyBase = rubyBaseRun?.Element(w + "t")?.Value;
					if (rubyText is null || rubyBase is null) {
						return [element];
					}

					List<RubySegment> rubyseg;
					try {
						var segs = SplitJapaneseTextManual(rubyBase);
						if (segs.Count < 2) {
							return [element];
						}
						rubyseg = GetRubySegments(segs, rubyText);
					}
					catch {
						return [element];
					}

					List<XElement> res = [];
					foreach (var rs in rubyseg) {
						res.Add(
						new XElement(w + "r",
							runPreference,
							(rs.Type == TextType.Hiragana) ?
							new XElement(w + "t",
								rs.Text
							) :
							new XElement(w + "ruby",
								rubyPreference,
								new XElement(w + "rt",
									new XElement(w + "r",
										rubyTextPr,
										new XElement(w + "t",
											rs.Ruby
										)
									)
								),
								new XElement(w + "rubyBase",
									new XElement(w + "r",
										rubyBasePr,
										new XElement(w + "t",
											rs.Text
										)
									)
								)
							)
						)
						);
					}
					return res.ToArray();
				});

				paragraph.ReplaceNodes(newRuns);
			}
			return (xml, Path.ChangeExtension(path, ".C"));
		}

		static List<RubySegment> GetRubySegments(List<TextSegment> segs, string rubyText) {
			var result = new List<RubySegment>();
			bool lastIsSet = true;
			int i = 0;
			{
				var first = segs.First();
				if (first.Type is TextType.Hiragana) {
					i = 1;
					if (rubyText.StartsWith(first.Text)) {
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
				var i0 = rubyText.IndexOf(seg.Text);
				var i1 = rubyText.LastIndexOf(seg.Text);
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
				if (rubyText.EndsWith(last.Text)) {
					var i0 = rubyText.LastIndexOf(last.Text);
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

				// 找到相同类型的连续字符
				while (i < input.Length && GetTextType(input[i]) == currentType) {
					i++;
				}

				string segmentText = input[start..i ];
				segments.Add(new TextSegment(segmentText, currentType));
			}

			return segments;
		}

		static TextType GetTextType(char c) {
			// 平假名: U+3040 - U+309F
			if (c >= '\u3040' && c <= '\u309F')
				return TextType.Hiragana;

			// 片假名: U+30A0 - U+30FF
			//if (c >= '\u30A0' && c <= '\u30FF')
			//	return TextType.Katakana;
			//
			// 汉字: U+4E00 - U+9FFF (CJK统一表意文字)
			//if (c >= '\u4E00' && c <= '\u9FFF')
			//	return TextType.Kanji;

			return TextType.Other;
		}

		enum TextType {
			Hiragana,   // 平假名
			Katakana,   // 片假名
			Kanji,      // 汉字
			Other       // 其他字符
		}

		record TextSegment(string Text, TextType Type);

		struct RubySegment(string _t, TextType _y, string _r) {
			public string Text = _t;
			public TextType Type = _y;
			public string Ruby = _r;
		}
	}
}
