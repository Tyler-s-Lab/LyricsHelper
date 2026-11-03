using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;

namespace LyricsHelper.RegressionTests;

internal static class Program {
	private static int _failures;

	private static async Task<int> Main() {
		string root = Path.Combine(Path.GetTempPath(), "lyricshelper-tests-" + Guid.NewGuid().ToString("N"));
		Directory.CreateDirectory(root);

		try {
			await CleaningPreservesMultiRunBase(root);
			await ExtractionCoversNestedStructures(root);
			await ExtractionHandlesStrictAndEmptyReading(root);
			await ExtractionDistinguishesColorStates(root);
			await SpecialRubyRules(root);
			await SymbolReadingsUseGeneralRubyRules(root);
		}
		finally {
			try {
				Directory.Delete(root, true);
			}
			catch {
				// Best effort only.
			}
		}

		if (_failures == 0) {
			Console.WriteLine("All regression tests passed.");
			return 0;
		}

		Console.Error.WriteLine($"{_failures} regression test(s) failed.");
		return 1;
	}

	private static async Task CleaningPreservesMultiRunBase(string root) {
		string input = MakeDocx(root, "multi.docx",
			@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body><w:p><w:r><w:ruby><w:rubyPr/><w:rt><w:r><w:t>まもります</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>守り</w:t></w:r><w:r><w:t>ます</w:t></w:r></w:rubyBase></w:ruby></w:r></w:p></w:body></w:document>");

		string result = await RubyCleanerForJapanese.TryProcess([input]);
		Check(!result.Contains("Error", StringComparison.Ordinal), "multi-run cleaning reported an error: " + result);

		string output = Path.ChangeExtension(input, ".C") + ".docx";
		Check(File.Exists(output), "multi-run cleaning did not create the output file");

		string xml = ReadEntry(output, "word/document.xml");
		Check(xml.Contains(">守</w:t>", StringComparison.Ordinal), "cleaned base lost the kanji");
		Check(xml.Contains(">ります</w:t>", StringComparison.Ordinal), "cleaned base lost the okurigana from the second run");
		Check(!xml.Contains("<w:r><w:r>", StringComparison.Ordinal), "cleaning produced nested runs");

		var document = System.Xml.Linq.XDocument.Parse(xml);
		System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
		Check(!document.Descendants(w + "r").Any(run =>
			run.Element(w + "ruby") is not null && run.Element(w + "t") is not null),
			"cleaning put a ruby and text in the same run");
	}

	private static async Task ExtractionCoversNestedStructures(string root) {
		string input = MakeDocx(root, "structure.docx",
			@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>表</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:hyperlink><w:r><w:t>リンク</w:t></w:r></w:hyperlink></w:p><w:sdt><w:sdtContent><w:p><w:r><w:t>コントロール</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>前</w:t><w:tab/><w:t>後</w:t></w:r></w:p></w:body></w:document>");

		string result = await RubyExtract.TryProcess([input]);
		Check(result.Contains("表", StringComparison.Ordinal), "table text was not extracted");
		Check(result.Contains("リンク", StringComparison.Ordinal), "hyperlink text was not extracted");
		Check(result.Contains("コントロール", StringComparison.Ordinal), "content-control text was not extracted");
		Check(result.Contains("前\t後", StringComparison.Ordinal), "tab was not extracted");
	}

	private static async Task ExtractionHandlesStrictAndEmptyReading(string root) {
		string strict = MakeDocx(root, "strict.docx",
			@"<w:document xmlns:w=""http://purl.oclc.org/ooxml/wordprocessingml/main""><w:body><w:p><w:r><w:ruby><w:rt><w:r><w:t>まも</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>守</w:t></w:r></w:rubyBase></w:ruby></w:r></w:p></w:body></w:document>");
		string strictResult = await RubyExtract.TryProcess([strict]);
		Check(strictResult.Contains("まも", StringComparison.Ordinal), "Strict OOXML was not processed");

		string emptyReading = MakeDocx(root, "empty-rt.docx",
			@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body><w:p><w:r><w:ruby><w:rt><w:r><w:t></w:t></w:r></w:rt><w:rubyBase><w:r><w:t>守ります</w:t></w:r></w:rubyBase></w:ruby></w:r></w:p></w:body></w:document>");
		string emptyResult = await RubyExtract.TryProcess([emptyReading]);
		Check(emptyResult.Contains("守ります", StringComparison.Ordinal), "empty reading discarded the base text");
	}

	private static async Task ExtractionDistinguishesColorStates(string root) {
		string automatic = MakeDocx(root, "auto-color.docx",
			@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body><w:p><w:r><w:ruby><w:rt><w:r><w:rPr><w:color w:val=""auto""/></w:rPr><w:t>まも</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>守</w:t></w:r></w:rubyBase></w:ruby></w:r></w:p></w:body></w:document>");
		string automaticResult = await RubyExtract.TryProcess([automatic]);
		Check(!automaticResult.Contains("[守]", StringComparison.Ordinal), "automatic color was treated as a mistake");

		string red = MakeDocx(root, "red-color.docx",
			@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body><w:p><w:r><w:ruby><w:rt><w:r><w:rPr><w:color w:val=""FF0000""/></w:rPr><w:t>まも</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>守</w:t></w:r></w:rubyBase></w:ruby></w:r></w:p></w:body></w:document>");
		string redResult = await RubyExtract.TryProcess([red]);
		Check(redResult.Contains("[守](まも)", StringComparison.Ordinal), "explicit non-black reading color was not reported");
	}

	private static async Task SpecialRubyRules(string root) {
		Check(RubyTextPolicy.TryGetHiraganaReading("こんぴゅーたー", out string computer)
			&& computer == "こんぴゅうたあ", "katakana reading was not normalised");
		Check(RubyTextPolicy.TryGetHiraganaReading("めもりー", out string memory)
			&& memory == "めもりい", "long vowel after り was not expanded");
		Check(RubyTextPolicy.TryGetHiraganaReading("いんたーねっと", out string internet)
			&& internet == "いんたあねっと", "long vowel after た was not expanded");
		Check(!RubyTextPolicy.TryGetHiraganaReading("INTERNET", out _), "English reading was treated as hiragana");

		Check(RubyTextPolicy.TryGetSpecialExtraction("アイウエオ", "あいうえお", out string katakana)
			&& katakana == "あいうえお", "full-katakana base did not use the hiragana reading");
		Check(RubyTextPolicy.TryGetSpecialExtraction("インターネット", "INTERNET", out string internetBase)
			&& internetBase == "いんたあねっと", "full-katakana base with English reading was not converted to hiragana");
		Check(RubyTextPolicy.TryGetSpecialExtraction("エンジェル", "ANGEL", out string angelBase)
			&& angelBase == "えんじぇる", "full-katakana base with English reading was not converted to hiragana");
		Check(RubyTextPolicy.TryGetSpecialExtraction("Dream", "Ｄｒｅａｍ", out string latin)
			&& latin == "Dream", "Latin base did not keep the base text");
		Check(RubyTextPolicy.TryGetSpecialExtraction("SNS", "ＳＮＳ", out string sns)
			&& sns == "SNS", "Latin base with full-width reading did not keep the base text");
		Check(RubyTextPolicy.TryGetSpecialExtraction("123", "１２３", out string digits)
			&& digits == "123", "digit base did not keep the base text");
		Check(!RubyTextPolicy.TryGetSpecialExtraction("☆", "ほし", out _), "symbol with a real reading was treated as special");
		Check(!RubyTextPolicy.TryGetSpecialExtraction("♪", "おんぷ", out _), "music note with a real reading was treated as special");
		Check(!RubyTextPolicy.TryGetSpecialExtraction("○", "まる", out _), "circle with a real reading was treated as special");
		Check(RubyTextPolicy.TryGetSpecialExtraction("…", "・・・", out string ellipsis)
			&& ellipsis == "…", "punctuation base did not keep the symbol");
		Check(RubyTextPolicy.TryGetSpecialExtraction("【", "（", out string bracket)
			&& bracket == "【", "bracket base did not keep the symbol");
		Check(!RubyTextPolicy.TryGetSpecialExtraction("守ります", "まもります", out _), "normal kanji base was treated as special");

		static string Paragraph(string baseText, string reading) {
			return $@"<w:p><w:r><w:ruby><w:rt><w:r><w:t>{reading}</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>{baseText}</w:t></w:r></w:rubyBase></w:ruby></w:r></w:p>";
		}

		string document = @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body>"
			+ Paragraph("アイウエオ", "あいうえお")
			+ Paragraph("コンピューター", "こんぴゅーたー")
			+ Paragraph("インターネット", "INTERNET")
			+ Paragraph("Dream", "Ｄｒｅａｍ")
			+ Paragraph("…", "・・・")
			+ @"</w:body></w:document>";

		string input = MakeDocx(root, "special.docx", document);
		string extracted = await RubyExtract.TryProcess([input]);
		Check(extracted.Contains("あいうえお", StringComparison.Ordinal), "special extraction lost katakana reading");
		Check(extracted.Contains("こんぴゅうたあ", StringComparison.Ordinal), "special extraction did not expand the long vowel");
		Check(extracted.Contains("いんたあねっと", StringComparison.Ordinal), "special extraction did not convert the katakana base with an English reading");
		Check(extracted.Contains("Dream", StringComparison.Ordinal), "special extraction lost the Latin base");
		Check(extracted.Contains("…", StringComparison.Ordinal), "special extraction lost the symbol base");

		string cleaningResult = await RubyCleanerForJapanese.TryProcess([input]);
		Check(!cleaningResult.Contains("Error", StringComparison.Ordinal), "special cleaning reported an error: " + cleaningResult);

		string output = Path.ChangeExtension(input, ".C") + ".docx";
		string xml = ReadEntry(output, "word/document.xml");
		Check(!xml.Contains("<w:ruby", StringComparison.Ordinal), "special cleaning left ruby nodes behind");
		Check(xml.Contains(">アイウエオ</w:t>", StringComparison.Ordinal), "special cleaning lost the katakana base");
		Check(xml.Contains(">Dream</w:t>", StringComparison.Ordinal), "special cleaning lost the Latin base");
		Check(xml.Contains(">…</w:t>", StringComparison.Ordinal), "special cleaning lost the symbol base");
	}

	private static async Task SymbolReadingsUseGeneralRubyRules(string root) {
		static string Paragraph(string baseText, string reading) {
			return $@"<w:p><w:r><w:ruby><w:rt><w:r><w:t>{reading}</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>{baseText}</w:t></w:r></w:rubyBase></w:ruby></w:r></w:p>";
		}

		string document = @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body>"
			+ Paragraph("☆", "ほし")
			+ Paragraph("♪", "おんぷ")
			+ Paragraph("○", "まる")
			+ @"</w:body></w:document>";

		string input = MakeDocx(root, "symbol-readings.docx", document);
		string extracted = await RubyExtract.TryProcess([input]);
		Check(extracted.Contains("ほし", StringComparison.Ordinal), "real symbol reading ほし was not extracted");
		Check(extracted.Contains("おんぷ", StringComparison.Ordinal), "real symbol reading おんぷ was not extracted");
		Check(extracted.Contains("まる", StringComparison.Ordinal), "real symbol reading まる was not extracted");

		string cleaningResult = await RubyCleanerForJapanese.TryProcess([input]);
		Check(!cleaningResult.Contains("Error", StringComparison.Ordinal), "symbol-reading cleaning reported an error: " + cleaningResult);

		string output = Path.ChangeExtension(input, ".C") + ".docx";
		string xml = ReadEntry(output, "word/document.xml");
		Check(xml.Contains("<w:ruby", StringComparison.Ordinal), "real symbol readings should keep their ruby");
		Check(xml.Contains(">☆</w:t>", StringComparison.Ordinal), "symbol-reading cleaning lost the symbol base");
		Check(xml.Contains(">ほし</w:t>", StringComparison.Ordinal), "symbol-reading cleaning lost the reading");
	}

	private static string MakeDocx(string root, string name, string documentXml) {
		string path = Path.Combine(root, name);
		using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write);
		using var zip = new ZipArchive(file, ZipArchiveMode.Create);
		using Stream entry = zip.CreateEntry("word/document.xml").Open();
		byte[] bytes = Encoding.UTF8.GetBytes(documentXml);
		entry.Write(bytes, 0, bytes.Length);
		return path;
	}

	private static string ReadEntry(string path, string entryName) {
		using var zip = ZipFile.OpenRead(path);
		ZipArchiveEntry entry = zip.GetEntry(entryName)
			?? throw new InvalidOperationException($"missing zip entry '{entryName}'");
		using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
		return reader.ReadToEnd();
	}

	private static void Check(bool condition, string message) {
		if (condition)
			return;

		_failures++;
		Console.Error.WriteLine("FAIL: " + message);
	}
}
