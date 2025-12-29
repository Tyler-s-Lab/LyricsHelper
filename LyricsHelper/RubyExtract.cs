using System.Xml.Linq;

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
			var paragraphs = body.Elements(w + "p");
			foreach (var paragraph in paragraphs) {
				var runs = paragraph.Elements(w + "r");
				foreach (var run in runs) {
					bool isMistack = false;
					if (run.Element(w + "rPr") is XElement runPreference) {
						if (runPreference.Element(w + "color") != null) {
							isMistack = true;
						}
					}
					if (run.Element(w + "t") is XElement text) {
						res += text.Value;//.Replace("“", null);
					}
					else if (run.Element(w + "ruby") is XElement ruby) {
						var rubyText = ruby.Element(w + "rt")?.Element(w + "r")?.Element(w + "t");
						var rubyBase = ruby.Element(w + "rubyBase")?.Element(w + "r")?.Element(w + "t");
						if (isMistack) {
							if (rubyBase != null) {
								res += $"[{rubyBase.Value}]";
							}
							if (rubyText != null) {
								res += $"({rubyText.Value})";
							}
						}
						else {
							if (rubyText != null) {
								res += rubyText.Value;
							}
							else if (rubyBase != null) {
								res += rubyBase.Value;
							}
						}
					}
				}
				res += '\n';
			}
			return res;
		}

	}
}
