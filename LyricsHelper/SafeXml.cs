using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace LyricsHelper {
	internal static class SafeXml {
		public static XmlReader CreateReader(Stream stream) {
			var settings = new XmlReaderSettings {
				// WordprocessingML does not need a DTD. Prohibiting it removes
				// the entity-expansion vector as well as external resolution.
				DtdProcessing = DtdProcessing.Prohibit,
				XmlResolver = null,
				IgnoreWhitespace = true,
				MaxCharactersInDocument = DocumentLimits.MaxXmlCharacters,
				MaxCharactersFromEntities = 1024
			};
			return XmlReader.Create(stream, settings);
		}

		public static XDocument Load(Stream stream) {
			using var reader = CreateReader(stream);
			return XDocument.Load(reader, LoadOptions.None);
		}

		public static byte[] Save(XDocument document, long maxBytes = DocumentLimits.MaxStoryPartBytes) {
			using var stream = new MemoryStream();
			document.Save(stream, SaveOptions.DisableFormatting);
			if (stream.Length > maxBytes)
				throw new DocumentLimitException("resulting document part exceeds the configured size limit");
			return stream.ToArray();
		}
	}
}
