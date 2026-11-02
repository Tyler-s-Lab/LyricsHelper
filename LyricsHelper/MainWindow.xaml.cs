using System.Windows;

namespace LyricsHelper;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window {
	private bool _isProcessing;

	public MainWindow() {
		InitializeComponent();
	}

	private async void TabItem_Drop_ToKANA(object sender, DragEventArgs e) {
		if (!e.Data.GetDataPresent(DataFormats.FileDrop) || e.Data.GetData(DataFormats.FileDrop) is not string[] paths) {
			return;
		}

		await ProcessAsync(paths, RubyExtract.TryProcess);
	}

	private async void TabItem_Drop_Clear(object sender, DragEventArgs e) {
		if (!e.Data.GetDataPresent(DataFormats.FileDrop) || e.Data.GetData(DataFormats.FileDrop) is not string[] paths) {
			return;
		}

		await ProcessAsync(paths, RubyCleanerForJapanese.TryProcess);
	}

	private async Task ProcessAsync(string[] paths, Func<string[], Task<string>> process) {
		if (_isProcessing)
			return;

		_isProcessing = true;
		try {
			// The result is shown even when it is empty: the window is the
			// user's signal that the task has finished (success or failure).
			string result = await process(paths);
			ShowResult(result);
		}
		finally {
			_isProcessing = false;
		}
	}

	private void ShowResult(string text) {
		var window = new WindowText {
			Owner = this,
			Text = text
		};
		window.Show();
	}
}
