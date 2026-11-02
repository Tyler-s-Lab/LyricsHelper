namespace LyricsHelper {
	/// <summary>
	/// Central place for every bound applied to untrusted documents and results.
	/// The values are deliberately generous for lyrics documents, but small
	/// enough that a dropped file cannot exhaust the machine.
	/// </summary>
	internal static class DocumentLimits {
		public const int MaxDirectoryFiles = 1024;
		public const int MaxZipEntries = 4096;
		public const int MaxStoryParts = 64;
		public const long MaxTotalDecompressedBytes = 128L * 1024 * 1024;
		public const long MaxStoryPartBytes = 32L * 1024 * 1024;
		public const long MaxFlatOpcBytes = 128L * 1024 * 1024;

		public const int MaxXmlCharacters = 32_000_000;

		public const int MaxResultChars = 1_000_000;
		public const int MaxDisplayChars = 1_000_000;

		public const long MaxRebuildNodes = 2_000_000;
		public const int MaxRubyBaseLength = 512;
		public const int MaxRubyTextLength = 4096;
	}
}
