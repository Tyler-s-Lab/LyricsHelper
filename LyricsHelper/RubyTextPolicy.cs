using static LyricsHelper.CharHelper;

namespace LyricsHelper {
	/// <summary>
	/// Classification and reading-normalisation rules for ruby annotations that
	/// Word documents use for katakana, Latin/digit and symbol bases.
	/// </summary>
	internal static class RubyTextPolicy {

		public static bool IsSpecialBase(string baseText) {
			return IsFullyKatakana(baseText)
				|| IsLatinOrDigitBase(baseText)
				|| IsSymbolBase(baseText);
		}

		/// <summary>
		/// True when the whole base consists of katakana (including the
		/// prolonged sound mark and katakana punctuation).
		/// </summary>
		public static bool IsFullyKatakana(string text) {
			if (string.IsNullOrEmpty(text))
				return false;

			bool hasKatakana = false;
			foreach (char c in text) {
				if (char.IsWhiteSpace(c))
					continue;
				if (!IsKatakana(c))
					return false;
				hasKatakana = true;
			}
			return hasKatakana;
		}

		/// <summary>
		/// True when the base is made of Latin letters/digits (ASCII or
		/// full-width), optionally with whitespace or simple punctuation.
		/// </summary>
		public static bool IsLatinOrDigitBase(string text) {
			if (string.IsNullOrEmpty(text))
				return false;

			bool hasLetterOrDigit = false;
			foreach (char c in text) {
				if (char.IsWhiteSpace(c))
					continue;
				if (IsLatinOrDigit(c)) {
					hasLetterOrDigit = true;
					continue;
				}
				if (char.IsPunctuation(c) || char.IsSymbol(c))
					continue;
				return false;
			}
			return hasLetterOrDigit;
		}

		/// <summary>
		/// True when the base is only symbols/punctuation (for example ☆, ♪,
		/// …, 【 or 】), i.e. there is no kana, kanji, letter or digit.
		/// </summary>
		public static bool IsSymbolBase(string text) {
			if (string.IsNullOrEmpty(text))
				return false;

			bool hasSymbol = false;
			foreach (char c in text) {
				if (char.IsWhiteSpace(c))
					continue;
				if (!char.IsPunctuation(c) && !char.IsSymbol(c))
					return false;
				hasSymbol = true;
			}
			return hasSymbol;
		}

		/// <summary>
		/// Returns the text that Stripping should emit for a special base.
		/// </summary>
		public static bool TryGetSpecialExtraction(string? baseText, string? reading, out string result) {
			if (string.IsNullOrEmpty(baseText)) {
				result = "";
				return false;
			}

			if (IsFullyKatakana(baseText)) {
				// Katakana is already phonetic; always emit hiragana. Prefer
				// the hiragana reading (normalising ー to an explicit vowel);
				// if the reading is English or otherwise not kana, normalise
				// the katakana base itself.
				if (TryGetHiraganaReading(reading, out string hiraganaReading)) {
					result = hiraganaReading;
				}
				else if (TryGetHiraganaReading(baseText, out string hiraganaBase)) {
					result = hiraganaBase;
				}
				else {
					result = baseText;
				}
				return true;
			}

			if (IsLatinOrDigitBase(baseText) || IsSymbolBase(baseText)) {
				result = baseText;
				return true;
			}

			result = "";
			return false;
		}

		/// <summary>
		/// Converts a katakana/hiragana reading to hiragana and expands ー to
		/// the vowel of the preceding kana. Returns false when the reading
		/// contains anything that is not kana (e.g. an English reading).
		/// </summary>
		public static bool TryGetHiraganaReading(string? reading, out string result) {
			result = "";
			if (string.IsNullOrEmpty(reading))
				return false;

			var builder = new System.Text.StringBuilder(reading.Length);
			foreach (char original in reading) {
				char c = original;

				if (c == '\u30FC') { // ー
					char? vowel = builder.Length == 0 ? null : VowelOf(builder[builder.Length - 1]);
					builder.Append(vowel ?? c);
					continue;
				}

				if (c == '\u30FB') { // ・
					builder.Append(c);
					continue;
				}

				if (char.IsWhiteSpace(c)) {
					builder.Append(c);
					continue;
				}

				if (c >= '\u30A1' && c <= '\u30FA')
					c = (char)(c - 0x60);
				else if (c >= '\u30FD' && c <= '\u30FE')
					c = (char)(c - 0x60);

				if (!IsHiragana(c))
					return false;

				builder.Append(c);
			}

			result = builder.ToString();
			return true;
		}

		private static bool IsKatakana(char c) {
			return (c >= '\u30A0' && c <= '\u30FF')
				|| (c >= '\u31F0' && c <= '\u31FF')
				|| (c >= '\uFF66' && c <= '\uFF9D');
		}

		private static bool IsLatinOrDigit(char c) {
			if (c >= 'A' && c <= 'Z') return true;
			if (c >= 'a' && c <= 'z') return true;
			if (c >= '0' && c <= '9') return true;

			// Full-width Latin letters and digits.
			if (c >= '\uFF21' && c <= '\uFF3A') return true;
			if (c >= '\uFF41' && c <= '\uFF5A') return true;
			if (c >= '\uFF10' && c <= '\uFF19') return true;

			return false;
		}

		private static char? VowelOf(char c) {
			return c switch {
				'あ' or 'か' or 'さ' or 'た' or 'な' or 'は' or 'ま' or 'や' or 'ら' or 'わ'
					or 'が' or 'ざ' or 'だ' or 'ば' or 'ぱ' or 'ぁ' or 'ゃ' or 'ゎ' => 'あ',
				'い' or 'き' or 'し' or 'ち' or 'に' or 'ひ' or 'み' or 'り'
					or 'ぎ' or 'じ' or 'ぢ' or 'び' or 'ぴ' or 'ぃ' or 'ゐ' => 'い',
				'う' or 'く' or 'す' or 'つ' or 'ぬ' or 'ふ' or 'む' or 'ゆ' or 'る'
					or 'ぐ' or 'ず' or 'づ' or 'ぶ' or 'ぷ' or 'ぅ' or 'ゅ' or 'ゔ' => 'う',
				'え' or 'け' or 'せ' or 'て' or 'ね' or 'へ' or 'め' or 'れ'
					or 'げ' or 'ぜ' or 'で' or 'べ' or 'ぺ' or 'ぇ' or 'ゑ' => 'え',
				'お' or 'こ' or 'そ' or 'と' or 'の' or 'ほ' or 'も' or 'よ' or 'ろ' or 'を'
					or 'ご' or 'ぞ' or 'ど' or 'ぼ' or 'ぽ' or 'ぉ' or 'ょ' => 'お',
				_ => null
			};
		}
	}
}
