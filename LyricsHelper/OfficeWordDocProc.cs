using SharpCompress.Archives;
using SharpCompress.Archives.Zip;
using SharpCompress.Common;
using System.IO;
using System.Xml.Linq;

namespace LyricsHelper {
	internal static class OfficeWordDocProc {

		internal struct ProcessResult() {
			public string? Message = null;

			public XContainer? ModifiedXml = null;
			public string TargetPath = "";
		}

		internal static Task<IEnumerable<string>> ReadFilesAsync(string[] paths, Func<XContainer, string?> func) {
			return Task.Run(() => ProccessFiles(paths, (_, d) => new ProcessResult() { Message = func(d) }));
		}

		internal static Task<IEnumerable<string>> ModifyFilesAsync(string[] paths, Func<string, XContainer, ValueTuple<XContainer, string>> func) {
			return Task.Run(() => ProccessFiles(paths, (p, d) => {
				var a = func(p, d); return new ProcessResult() { ModifiedXml = a.Item1, TargetPath = a.Item2 };
			}));
		}

		static IEnumerable<string> ProccessFiles(string[] paths, Func<string, XContainer, ProcessResult> func) {
			return paths.SelectMany(x => {
				try {
					if (File.Exists(x)) {
						return [TryProcessFile(x, func)];
					}
					if (Directory.Exists(x)) {
						return Directory.EnumerateFiles(x, "*.*", SearchOption.AllDirectories).Select(y => TryProcessFile(y, func));
					}
				}
				catch (Exception ex) {
					return [$"[Error processing '{x}': {ex.Message}]"];
				}
				return [$"[Unable to process '{x}']"];
			});
		}

		static string TryProcessFile(string path, Func<string, XContainer, ProcessResult> func) {
			try {
				return ProcessFile(path, func) ?? $"[Success processing file '{path}']";
			}
			catch (Exception ex) {
				return $"[Error processing file '{path}': {ex.Message}]";
			}
		}


		static string? ProcessFile(string path, Func<string, XContainer, ProcessResult> func) {
			if (Path.GetExtension(path).Equals(".docx", StringComparison.OrdinalIgnoreCase)) {
				if (!ArchiveFactory.IsArchive(path, out ArchiveType? archiveType) || archiveType != ArchiveType.Zip) {
					throw new Exception(".docx is not a valid archive file");
				}

				using IArchive archive = ArchiveFactory.Open(path);
				if (archive is not ZipArchive zip) {
					throw new Exception("failed to open zip");
				}

				var docEntries = zip.Entries.Where(x => x.Key == @"word/document.xml");
				if (docEntries.Count() != 1) {
					throw new Exception("the entry in zip not found");
				}

				var docEntry = docEntries.First();
				using var docStream = docEntry.OpenEntryStream();

				XDocument xml = XDocument.Load(docStream);

				var res = func(path, xml);

				if (res.ModifiedXml != null) {
					var new_xml = res.ModifiedXml as XDocument ?? throw new Exception("error processing xml as XDocument") ;

					zip.RemoveEntry(docEntry);

					using MemoryStream memoryStream = new();
					new_xml.Save(memoryStream, SaveOptions.DisableFormatting);
					zip.AddEntry(@"word/document.xml", memoryStream, true);

					using FileStream fileStream1 = new(res.TargetPath + ".docx", FileMode.CreateNew, FileAccess.ReadWrite);
					zip.SaveTo(fileStream1);
				}
				return res.Message;
			}
			if (Path.GetExtension(path).Equals(".xml", StringComparison.OrdinalIgnoreCase)) {
				using FileStream fileStream = new(path, FileMode.Open, FileAccess.Read);
				XDocument xml = XDocument.Load(fileStream);
				XNamespace pkg = @"http://schemas.microsoft.com/office/2006/xmlPackage";

				if (xml.Element(pkg + "package") is not XElement package) {
					throw new Exception("XML package format error");
				}

				var document_parts = package.Elements(pkg + "part").Where(x => (string?)x.Attribute(pkg + "name") == "/word/document.xml");
				if (document_parts.Take(2).Count() != 1) {
					throw new Exception("XML object not existing");
				}

				if (document_parts.First().Element(pkg + "xmlData") is not XElement xmlData) {
					throw new Exception("XML data missing");
				}

				var res = func(path, xmlData);

				if (res.ModifiedXml != null) {
					var new_xml = res.ModifiedXml as XElement ?? throw new Exception("error processing xml as XElement");

					xmlData.ReplaceWith(new_xml);

					using FileStream fileStream1 = new(res.TargetPath + ".xml", FileMode.CreateNew, FileAccess.ReadWrite);
					xml.Save(fileStream1, SaveOptions.DisableFormatting);
				}
				return res.Message;
			}
			throw new Exception("file format not supported");
		}

	}
}
