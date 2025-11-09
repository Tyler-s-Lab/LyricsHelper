using SharpCompress.Archives;
using SharpCompress.Archives.Zip;
using SharpCompress.Common;
using System.IO;
using System.Windows;
using System.Xml.Linq;

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
		await ToKanaAsync(paths);
		return;
	}

	private async Task ToKanaAsync(string[] paths) {
		await Task.Run(() => { ToKana(paths); });
		return;
	}

	private static void ToKana(string[] paths) {
		string res = "";
		foreach (string path in paths) {
			try {
				if (!Path.GetExtension(path).Equals(".docx", StringComparison.OrdinalIgnoreCase)) {
					res += "no0\n";
					continue;
				}
				if (!ArchiveFactory.IsArchive(path, out ArchiveType? archiveType) || archiveType != ArchiveType.Zip) {
					res += "no1\n";
					continue;
				}
				using IArchive archive = ArchiveFactory.Open(path);
				if (archive is not ZipArchive zip) {
					res += "no2\n";
					continue;
				}

				var docEntries = zip.Entries.Where(x => x.Key == @"word/document.xml");
				if (docEntries.Count() != 1) {
					res += "no3\n";
					continue;
				}
				var docEntry = docEntries.First();
				using var docStream = docEntry.OpenEntryStream();

				XDocument xml = XDocument.Load(docStream);
				XNamespace w = @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
				if (xml.Element(w + "document") is not XElement document) {
					continue;
				}
				if (document.Element(w + "body") is not XElement body) {
					continue;
				}
				var paragraphs = body.Elements(w + "p");
				foreach (var paragraph in paragraphs) {
					var runs = paragraph.Elements(w + "r");
					foreach (var run in runs) {
						bool isMistack = false;
						if (run.Element(w + "rPr") is XElement runPreference) {
							if (runPreference.Element(w + "color") != null) {
								isMistack = true;
							}
						}
						if (run.Element(w + "t") is XElement text) {
							res += text.Value;//.Replace("“", null);
						}
						else if (run.Element(w + "ruby") is XElement ruby) {
							var rubyText = ruby.Element(w + "rt")?.Element(w + "r")?.Element(w + "t");
							var rubyBase = ruby.Element(w + "rubyBase")?.Element(w + "r")?.Element(w + "t");
							if (isMistack) {
								if (rubyBase != null) {
									res += $"[{rubyBase.Value}]";
								}
								if (rubyText != null) {
									res += $"({rubyText.Value})";
								}
							}
							else {
								if (rubyText != null) {
									res += rubyText.Value;
								}
								else if (rubyBase != null) {
									res += rubyBase.Value;
								}
							}
						}
					}
					res += '\n';
				}
			}
			catch (Exception ex) {
				res += ex.Message;
			}
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
		await RubyCleanerForJapanese.TryProcessFilesAsync(paths);
		return;
	}
}