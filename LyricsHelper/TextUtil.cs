using System.Text;

namespace LyricsHelper {
	internal static class TextUtil {
		/// <summary>
		/// Joins user-visible fragments without ever materialising more than
		/// <paramref name="maxChars"/> characters (plus a small marker).
		/// </summary>
		public static string JoinCapped(IEnumerable<string> fragments, string separator, int maxChars, string marker = "\n[output truncated]") {
			var builder = new StringBuilder(Math.Min(maxChars, 4096));

			foreach (string fragment in fragments) {
				if (string.IsNullOrEmpty(fragment))
					continue;

				if (builder.Length + fragment.Length > maxChars) {
					int room = maxChars - builder.Length;
					if (room > 0)
						builder.Append(fragment, 0, room);
					builder.Append(marker);
					break;
				}

				if (builder.Length > 0)
					builder.Append(separator);
				builder.Append(fragment);

				if (builder.Length >= maxChars) {
					builder.Append(marker);
					break;
				}
			}

			return builder.ToString();
		}
	}
}
