namespace LyricsHelper {
	internal static class CharHelper {

		public static TextType GetTextType(char c) {
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

		public enum TextType {
			Hiragana,   // 平假名
			Katakana,   // 片假名
			Kanji,      // 汉字
			Other       // 其他字符
		}
	}
}
