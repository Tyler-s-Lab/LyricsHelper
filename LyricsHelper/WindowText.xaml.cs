using System.Windows;

namespace LyricsHelper {
	/// <summary>
	/// WindowText.xaml 的交互逻辑
	/// </summary>
	public partial class WindowText : Window {
		public WindowText() {
			InitializeComponent();
		}

		public string Text {
			get {
				return textbox.Text;
			}
			set {
				textbox.Text = value;
			}
		}
	}
}
