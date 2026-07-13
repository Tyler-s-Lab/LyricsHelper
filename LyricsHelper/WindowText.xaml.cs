using System.Windows;
using System.Windows.Input;

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

		private void Textbox_MouseWheel(object sender, MouseWheelEventArgs e) {
			if (!(Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))) {
				e.Handled = false;
				return;
			}

			if (e.Delta > 0) {
				if (textbox.FontSize < 256)
					textbox.FontSize *= 1.05;
				e.Handled = true;
				return;
			}
			if (e.Delta < 0) {
				if (textbox.FontSize > 2)
					textbox.FontSize /= 1.05;
				e.Handled = true;
				return;
			}
			e.Handled = false;
			return;
		}
	}
}
