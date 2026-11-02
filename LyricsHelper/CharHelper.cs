namespace LyricsHelper {
	internal static class CharHelper {

		public static TextType GetTextType(char c) {
			if (IsHiragana(c))
				return TextType.Hiragana;

			// 片假名（含片假名语音扩展与半角片假名）
			if ((c >= '\u30A0' && c <= '\u30FF')
				|| (c >= '\u31F0' && c <= '\u31FF')
				|| (c >= '\uFF66' && c <= '\uFF9D'))
				return TextType.Katakana;

			// 汉字（CJK 统一表意文字、扩展 A、兼容表意文字）
			if ((c >= '\u3400' && c <= '\u4DBF')
				|| (c >= '\u4E00' && c <= '\u9FFF')
				|| (c >= '\uF900' && c <= '\uFAFF'))
				return TextType.Kanji;

			return TextType.Other;
		}

		public static bool IsHiragana(char c) {
			return c >= '\u3040' && c <= '\u309F';
		}

		/// <summary>
		/// A hiragana reading may legitimately contain the prolonged sound mark
		/// (U+30FC), which otherwise belongs to the katakana block.
		/// </summary>
		public static bool IsHiraganaReadingChar(char c) {
			return IsHiragana(c) || c == '\u30FC';
		}

		public enum TextType {
			Hiragana,   // 平假名
			Katakana,   // 片假名
			Kanji,      // 汉字
			Other       // 其他字符
		}
	}
}
