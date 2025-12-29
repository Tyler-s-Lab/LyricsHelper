using System.Windows;

namespace LyricsHelper;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window {
	public MainWindow() {
		InitializeComponent();
	}

	private async void TabItem_Drop_ToKANA(object sender, DragEventArgs e) {
		if (!e.Data.GetDataPresent(DataFormats.FileDrop) || e.Data.GetData(DataFormats.FileDrop) is not string[] paths) {
			return;
		}
		var res = await RubyExtract.TryProcess(paths);
		if (res != null) {
			App.Current.Dispatcher.Invoke(new Action(() => {
				var wint = new WindowText {
					Text = res
				};
				wint.ShowDialog();
			}));
		}
		return;
	}

	private async void TabItem_Drop_Clear(object sender, DragEventArgs e) {
		if (!e.Data.GetDataPresent(DataFormats.FileDrop) || e.Data.GetData(DataFormats.FileDrop) is not string[] paths) {
			return;
		}
		var res = await RubyCleanerForJapanese.TryProcess(paths);
		if (res != null) {
			App.Current.Dispatcher.Invoke(new Action(() => {
				var wint = new WindowText {
					Text = res
				};
				wint.ShowDialog();
			}));
		}
		return;
	}
}