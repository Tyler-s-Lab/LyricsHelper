using System.Xml.Linq;
using static LyricsHelper.CharHelper;

namespace LyricsHelper {
	internal static class RubyExtract {

		internal static async Task<string> TryProcess(string[] paths) {
			string err = "";
			try {
				var res = await OfficeWordDocProc.ReadFilesAsync(paths, ProcessXml);
				err = string.Join("\n\n", res);
			}
			catch (Exception ex) {
				err = ex.Message;
			}
			return err;
		}

		static string? ProcessXml(XContainer xml) {
			string res = "";
			XNamespace w = @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
			if (xml.Element(w + "document") is not XElement document) {
				throw new Exception("XML not including w:document");
			}
			if (document.Element(w + "body") is not XElement body) {
				throw new Exception("XML not including w:body");
			}

			foreach (var paragraph in body.Elements(w + "p")) {
				foreach (var run in paragraph.Elements(w + "r")) {
					bool isMistakeRun = false; // 带颜色的Ruby表明读音并非正确 + 包含平假名以外的ruby

					foreach (var ele in run.Elements()) {
						if (ele.Name == w + "rPr" && ele is XElement runPreference) {
							bool isBlack = runPreference.Element(w + "color")?.Attribute(w + "val")?.Value?.Equals("000000") ?? true;
							isMistakeRun = isMistakeRun || !isBlack;
						}
						else if (ele.Name == w + "t" && ele is XElement text) {
							res += text.Value;//.Replace("“", null);
						}
						else if (ele.Name == w + "ruby" && ele is XElement ruby) {
							var rubyText = ruby.Element(w + "rt")?.Element(w + "r")?.Element(w + "t")?.Value;
							var rubyBase = ruby.Element(w + "rubyBase")?.Element(w + "r")?.Element(w + "t")?.Value;

							bool isMistakeRuby = rubyText?.Any(c => GetTextType(c) != TextType.Hiragana) ?? false;

							if (isMistakeRun || isMistakeRuby) {
								res += $"[{rubyBase}]";
								res += $"({rubyText})";
							}
							else {
								res += rubyText ?? rubyBase;
							}
						}
						else if (ele.Name == w + "br") {
							res += '\n';
						}
					}

				}
				res += Environment.NewLine;
			}
			return res;
		}

	}
}
