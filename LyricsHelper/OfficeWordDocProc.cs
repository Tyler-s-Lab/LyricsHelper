using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

namespace LyricsHelper {
	/// <summary>
	/// Reads and rewrites Word documents without trusting any size, count or
	/// compression metadata from the dropped file. The container work is done
	/// with the framework zip implementation; SharpCompress remains referenced
	/// by the project but is no longer on the untrusted intake path.
	/// </summary>
	internal static class OfficeWordDocProc {

		internal static Task<IEnumerable<string>> ReadFilesAsync(string[] paths, Func<XElement, string?> func) {
			return Task.Run<IEnumerable<string>>(() => ProcessFiles(paths, path => TryProcessFile(path, func, null)).ToList());
		}

		internal static Task<IEnumerable<string>> ModifyFilesAsync(string[] paths, Func<XElement, bool> func) {
			return Task.Run<IEnumerable<string>>(() => ProcessFiles(paths, path => TryProcessFile(path, null, func)).ToList());
		}

		static List<string> ProcessFiles(string[] paths, Func<string, string> process) {
			var results = new List<string>();
			foreach (string path in ExpandPaths(paths))
				results.Add(process(path));
			return results;
		}

		static IEnumerable<string> ExpandPaths(string[] paths) {
			foreach (string path in paths) {
				if (File.Exists(path)) {
					yield return path;
					continue;
				}

				if (Directory.Exists(path)) {
					int count = 0;
					foreach (string file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories)) {
						if (!IsSupportedInput(file))
							continue;
						if (++count > DocumentLimits.MaxDirectoryFiles)
							throw new DocumentLimitException($"directory contains more than {DocumentLimits.MaxDirectoryFiles} supported files");
						yield return file;
					}
					continue;
				}

				// Let ProcessFile report the unsupported path.
				yield return path;
			}
		}

		static bool IsSupportedInput(string path) {
			string extension = Path.GetExtension(path);
			if (!extension.Equals(".docx", StringComparison.OrdinalIgnoreCase)
				&& !extension.Equals(".xml", StringComparison.OrdinalIgnoreCase))
				return false;

			string fileName = Path.GetFileName(path);
			if (fileName.StartsWith("~$", StringComparison.Ordinal))
				return false;

			// Do not feed a previous Cleaning result back into the tool.
			string stem = Path.GetFileNameWithoutExtension(path);
			return !stem.EndsWith(".C", StringComparison.OrdinalIgnoreCase);
		}

		static string TryProcessFile(string path, Func<XElement, string?>? read, Func<XElement, bool>? modify) {
			try {
				return ProcessFile(path, read, modify) ?? $"[Success processing file '{path}']";
			}
			catch (Exception ex) {
				return $"[Error processing file '{path}': {ex.Message}]";
			}
		}

		static string? ProcessFile(string path, Func<XElement, string?>? read, Func<XElement, bool>? modify) {
			string extension = Path.GetExtension(path);
			if (extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
				return ProcessDocx(path, read, modify);
			if (extension.Equals(".xml", StringComparison.OrdinalIgnoreCase))
				return ProcessFlatOpc(path, read, modify);

			throw new Exception("file format not supported");
		}

		static string? ProcessDocx(string path, Func<XElement, string?>? read, Func<XElement, bool>? modify) {
			using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
			using var zip = new ZipArchive(input, ZipArchiveMode.Read);

			if (zip.Entries.Count > DocumentLimits.MaxZipEntries)
				throw new DocumentLimitException($"archive contains more than {DocumentLimits.MaxZipEntries} entries");

			int mainParts = zip.Entries.Count(entry => WordStoryParts.IsMainDocumentPart(entry.FullName));
			if (mainParts != 1)
				throw new Exception("the entry word/document.xml was not found exactly once");

			var storyEntries = zip.Entries
				.Where(entry => !entry.FullName.EndsWith("/", StringComparison.Ordinal) && WordStoryParts.IsStoryPart(entry.FullName))
				.ToList();
			if (storyEntries.Count > DocumentLimits.MaxStoryParts)
				throw new DocumentLimitException($"archive contains more than {DocumentLimits.MaxStoryParts} Word story parts");

			return read is not null
				? ReadDocx(storyEntries, read)
				: RewriteDocx(path, zip, modify!);
		}

		static string? ReadDocx(List<ZipArchiveEntry> storyEntries, Func<XElement, string?> read) {
			var budget = new ByteBudget(DocumentLimits.MaxTotalDecompressedBytes);
			var messages = new List<string>();

			foreach (var entry in storyEntries) {
				XDocument document = LoadPart(entry, budget);
				XElement root = document.Root ?? throw new Exception("XML part has no root element");
				if (!WordStoryParts.IsStoryRoot(root))
					throw new Exception($"unsupported Word story part '{entry.FullName}'");

				string? message = read(root);
				if (!string.IsNullOrEmpty(message))
					messages.Add(message!);
			}

			return messages.Count == 0 ? null : TextUtil.JoinCapped(messages, "\n", DocumentLimits.MaxResultChars);
		}

		static XDocument LoadPart(ZipArchiveEntry entry, ByteBudget budget) {
			if (entry.Length > DocumentLimits.MaxStoryPartBytes)
				throw new DocumentLimitException($"document part '{entry.FullName}' exceeds the configured size limit");

			using Stream raw = entry.Open();
			using var bounded = new BoundedReadStream(raw, DocumentLimits.MaxStoryPartBytes, budget);
			using var reader = SafeXml.CreateReader(bounded);
			return XDocument.Load(reader, LoadOptions.None);
		}

		static string? RewriteDocx(string path, ZipArchive zip, Func<XElement, bool> modify) {
			string target = GetTargetPath(path, ".docx");
			string temp = GetTempPath(target);

			try {
				using (var output = new FileStream(temp, FileMode.CreateNew, FileAccess.Write))
				using (var rewritten = new ZipArchive(output, ZipArchiveMode.Create)) {
					var budget = new ByteBudget(DocumentLimits.MaxTotalDecompressedBytes);

					foreach (ZipArchiveEntry entry in zip.Entries) {
						if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) {
							rewritten.CreateEntry(entry.FullName);
							continue;
						}

						if (WordStoryParts.IsStoryPart(entry.FullName)) {
							XDocument document = LoadPart(entry, budget);
							XElement root = document.Root ?? throw new Exception("XML part has no root element");
							if (!WordStoryParts.IsStoryRoot(root))
								throw new Exception($"unsupported Word story part '{entry.FullName}'");

							if (!modify(root)) {
								CopyEntry(entry, rewritten, budget);
								continue;
							}

							byte[] serialized = SafeXml.Save(document);
							WriteEntry(rewritten, entry, serialized);
							continue;
						}

						CopyEntry(entry, rewritten, budget);
					}
				}

				File.Move(temp, target, true);
				return null;
			}
			catch {
				TryDelete(temp);
				throw;
			}
		}

		static void CopyEntry(ZipArchiveEntry source, ZipArchive rewritten, ByteBudget budget) {
			CompressionLevel level = source.Length > 0 && source.CompressedLength >= source.Length * 0.95
				? CompressionLevel.NoCompression
				: CompressionLevel.Optimal;

			ZipArchiveEntry target = rewritten.CreateEntry(source.FullName, level);
			TrySetTimestamp(target, source);

			using Stream raw = source.Open();
			using var bounded = new BoundedReadStream(raw, DocumentLimits.MaxTotalDecompressedBytes, budget);
			using Stream destination = target.Open();
			bounded.CopyTo(destination, 81920);
		}

		static void WriteEntry(ZipArchive rewritten, ZipArchiveEntry source, byte[] content) {
			ZipArchiveEntry target = rewritten.CreateEntry(source.FullName, CompressionLevel.Optimal);
			TrySetTimestamp(target, source);
			using Stream destination = target.Open();
			destination.Write(content, 0, content.Length);
		}

		static void TrySetTimestamp(ZipArchiveEntry target, ZipArchiveEntry source) {
			try {
				target.LastWriteTime = source.LastWriteTime;
			}
			catch {
				// Some archives carry timestamps that zip cannot represent.
			}
		}

		static string? ProcessFlatOpc(string path, Func<XElement, string?>? read, Func<XElement, bool>? modify) {
			using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
			if (input.Length > DocumentLimits.MaxFlatOpcBytes)
				throw new DocumentLimitException("XML package exceeds the configured size limit");

			var budget = new ByteBudget(DocumentLimits.MaxFlatOpcBytes);
			XDocument packageDocument;
			using (var bounded = new BoundedReadStream(input, DocumentLimits.MaxFlatOpcBytes, budget))
				packageDocument = SafeXml.Load(bounded);

			XNamespace pkg = @"http://schemas.microsoft.com/office/2006/xmlPackage";
			XElement package = packageDocument.Root ?? throw new Exception("XML package format error");
			if (package.Name != pkg + "package")
				throw new Exception("XML package format error");

			var parts = package.Elements(pkg + "part")
				.Where(part => WordStoryParts.IsStoryPart((string?)part.Attribute(pkg + "name")))
				.ToList();
			if (parts.Count > DocumentLimits.MaxStoryParts)
				throw new DocumentLimitException($"XML package contains more than {DocumentLimits.MaxStoryParts} Word story parts");

			if (!parts.Any(part => WordStoryParts.IsMainDocumentPart((string?)part.Attribute(pkg + "name"))))
				throw new Exception("XML object not existing");

			if (read is not null) {
				var messages = new List<string>();
				foreach (var part in parts) {
					XElement root = GetFlatOpcStoryRoot(part, pkg);
					string? message = read(root);
					if (!string.IsNullOrEmpty(message))
						messages.Add(message!);
				}
				return messages.Count == 0 ? null : TextUtil.JoinCapped(messages, "\n", DocumentLimits.MaxResultChars);
			}

			foreach (var part in parts) {
				XElement root = GetFlatOpcStoryRoot(part, pkg);
				modify!(root);
			}

			// Always produce the .C.xml output. If nothing changed this is a
			// faithful package rewrite; if something changed the caller returns
			// success and the user receives the cleaned copy.
			WriteFlatOpc(path, packageDocument);
			return null;
		}

		static XElement GetFlatOpcStoryRoot(XElement part, XNamespace pkg) {
			XElement xmlData = part.Element(pkg + "xmlData") ?? throw new Exception("XML data missing");
			XElement root = xmlData.Elements().FirstOrDefault(WordStoryParts.IsStoryRoot)
				?? throw new Exception("XML story root missing");
			return root;
		}

		static void WriteFlatOpc(string path, XDocument packageDocument) {
			string target = GetTargetPath(path, ".xml");
			string temp = GetTempPath(target);

			try {
				using (var output = new FileStream(temp, FileMode.CreateNew, FileAccess.Write)) {
					packageDocument.Save(output, SaveOptions.DisableFormatting);
					if (output.Length > DocumentLimits.MaxFlatOpcBytes)
						throw new DocumentLimitException("resulting XML package exceeds the configured size limit");
				}

				File.Move(temp, target, true);
			}
			catch {
				TryDelete(temp);
				throw;
			}
		}

		static string GetTargetPath(string path, string extension) {
			return Path.ChangeExtension(path, ".C") + extension;
		}

		static string GetTempPath(string target) {
			string directory = Path.GetDirectoryName(target) ?? ".";
			return Path.Combine(directory, $".{Path.GetFileName(target)}.{Guid.NewGuid():N}.tmp");
		}

		static void TryDelete(string path) {
			try {
				if (File.Exists(path))
					File.Delete(path);
			}
			catch {
				// Best-effort cleanup; the original error is more useful.
			}
		}
	}
}
