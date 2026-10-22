using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public class SongLibraryForm : Form
    {
        private const string SettingsKey = @"SOFTWARE\WorshipHelper";
        private const string LibraryFolderValue = "SongLibraryFolder";

        private readonly TextBox searchBox = new TextBox();
        private readonly ListView songList = new ListView();
        private readonly Label statusLabel = new Label();
        private readonly Button chooseFolderButton = new Button();
        private readonly Button refreshButton = new Button();
        private readonly Button insertButton = new Button();
        private readonly Button favouriteButton = new Button();
        private readonly CheckBox favouritesOnly = new CheckBox();
        private readonly List<SongFile> songs = new List<SongFile>();
        private string libraryFolder;
        private bool inserting;

        public SongLibraryForm()
        {
            Text = "VMMC Song Library";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(520, 420);
            Size = new Size(650, 520);
            ShowInTaskbar = false;

            var topPanel = new Panel { Dock = DockStyle.Top, Height = 64, Padding = new Padding(8) };
            searchBox.Dock = DockStyle.Top;
            searchBox.Font = new Font(Font.FontFamily, 10);
            searchBox.TextChanged += delegate { ApplyFilter(); };
            searchBox.KeyDown += SearchBox_KeyDown;
            topPanel.Controls.Add(searchBox);

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            chooseFolderButton.Text = "Choose folder";
            chooseFolderButton.AutoSize = true;
            chooseFolderButton.Click += ChooseFolderButton_Click;
            refreshButton.Text = "Refresh";
            refreshButton.AutoSize = true;
            refreshButton.Click += delegate { RefreshLibrary(); };
            favouritesOnly.Text = "Favourites only";
            favouritesOnly.AutoSize = true;
            favouritesOnly.Margin = new Padding(12, 5, 0, 0);
            favouritesOnly.CheckedChanged += delegate { ApplyFilter(); };
            buttonPanel.Controls.Add(chooseFolderButton);
            buttonPanel.Controls.Add(refreshButton);
            buttonPanel.Controls.Add(favouritesOnly);
            topPanel.Controls.Add(buttonPanel);

            songList.Dock = DockStyle.Fill;
            songList.View = View.Details;
            songList.FullRowSelect = true;
            songList.HideSelection = false;
            songList.MultiSelect = false;
            songList.Columns.Add("Song", 430);
            songList.Columns.Add("Favourite", 90);
            songList.DoubleClick += delegate { InsertSelectedSong(); };
            songList.KeyDown += SongList_KeyDown;
            songList.SelectedIndexChanged += delegate { UpdateButtons(); };

            var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 48, Padding = new Padding(8) };
            statusLabel.AutoEllipsis = true;
            statusLabel.Dock = DockStyle.Fill;
            insertButton.Text = "Insert";
            insertButton.Dock = DockStyle.Right;
            insertButton.Width = 80;
            insertButton.Click += delegate { InsertSelectedSong(); };
            favouriteButton.Text = "Favourite";
            favouriteButton.Dock = DockStyle.Right;
            favouriteButton.Width = 90;
            favouriteButton.Click += delegate { ToggleFavourite(); };
            bottomPanel.Controls.Add(statusLabel);
            bottomPanel.Controls.Add(favouriteButton);
            bottomPanel.Controls.Add(insertButton);

            Controls.Add(songList);
            Controls.Add(bottomPanel);
            Controls.Add(topPanel);
            libraryFolder = GetSetting(LibraryFolderValue);
            Shown += delegate { RefreshLibrary(); };
        }

        private void ChooseFolderButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Choose the folder containing VMMC song PowerPoint files";
                dialog.SelectedPath = Directory.Exists(libraryFolder)
                    ? libraryFolder
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                libraryFolder = dialog.SelectedPath;
                SetSetting(LibraryFolderValue, libraryFolder);
                RefreshLibrary();
            }
        }

        private void RefreshLibrary()
        {
            songs.Clear();
            songList.Items.Clear();
            if (string.IsNullOrWhiteSpace(libraryFolder))
            {
                statusLabel.Text = "Choose a folder containing .ppt or .pptx files.";
                UpdateButtons();
                return;
            }

            if (!Directory.Exists(libraryFolder))
            {
                statusLabel.Text = "The configured song folder is not available.";
                UpdateButtons();
                return;
            }

            try
            {
                foreach (var path in Directory.EnumerateFiles(libraryFolder, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(IsPowerPointFile))
                {
                    songs.Add(new SongFile(path, IsFavourite(path)));
                }
                songs.Sort((left, right) => StringComparer.CurrentCultureIgnoreCase.Compare(left.Title, right.Title));
                ApplyFilter();
                statusLabel.Text = songs.Count == 0
                    ? "No .ppt or .pptx files were found in this folder."
                    : string.Format("{0} song{1} indexed.", songs.Count, songs.Count == 1 ? "" : "s");
            }
            catch (UnauthorizedAccessException)
            {
                statusLabel.Text = "VMMC cannot read this folder. Choose a folder you can access.";
            }
            catch (IOException exception)
            {
                statusLabel.Text = "VMMC could not scan the folder: " + exception.Message;
            }
            UpdateButtons();
        }

        private void ApplyFilter()
        {
            string query = Normalize(searchBox.Text);
            songList.BeginUpdate();
            try
            {
                songList.Items.Clear();
                foreach (var song in songs.Where(item =>
                    (!favouritesOnly.Checked || item.IsFavourite)
                    && (string.IsNullOrEmpty(query) || Normalize(item.Title).Contains(query))))
                {
                    var item = new ListViewItem(song.Title);
                    item.SubItems.Add(song.IsFavourite ? "Yes" : "");
                    item.Tag = song;
                    songList.Items.Add(item);
                }
                if (songList.Items.Count > 0)
                {
                    songList.Items[0].Selected = true;
                }
            }
            finally
            {
                songList.EndUpdate();
            }
            UpdateButtons();
        }

        private void SearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
            {
                MoveSelection(e.KeyCode == Keys.Down ? 1 : -1);
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                InsertSelectedSong();
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                searchBox.Clear();
                foreach (ListViewItem item in songList.SelectedItems)
                {
                    item.Selected = false;
                }
                searchBox.Focus();
                e.SuppressKeyPress = true;
            }
        }

        private void SongList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                InsertSelectedSong();
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                searchBox.Focus();
                e.SuppressKeyPress = true;
            }
        }

        private void MoveSelection(int direction)
        {
            if (songList.Items.Count == 0)
            {
                return;
            }
            int index = songList.SelectedIndices.Count == 0 ? 0 : songList.SelectedIndices[0] + direction;
            index = Math.Max(0, Math.Min(songList.Items.Count - 1, index));
            songList.Items[index].Selected = true;
            songList.Items[index].EnsureVisible();
        }

        private void InsertSelectedSong()
        {
            if (inserting || songList.SelectedItems.Count == 0)
            {
                return;
            }

            var song = songList.SelectedItems[0].Tag as SongFile;
            if (song == null || !File.Exists(song.Path))
            {
                statusLabel.Text = "This song is no longer available. Refresh the library.";
                return;
            }

            inserting = true;
            UpdateButtons();
            statusLabel.Text = "Inserting " + song.Title + "...";
            try
            {
                new SongManager().InsertSongFromFile(song.Path);
                statusLabel.Text = song.Title + " inserted.";
            }
            finally
            {
                inserting = false;
                UpdateButtons();
                searchBox.Focus();
            }
        }

        private void ToggleFavourite()
        {
            if (songList.SelectedItems.Count == 0)
            {
                return;
            }
            var song = songList.SelectedItems[0].Tag as SongFile;
            if (song == null)
            {
                return;
            }

            using (var key = Registry.CurrentUser.CreateSubKey(SettingsKey + @"\Favourites"))
            {
                if (song.IsFavourite)
                {
                    key.DeleteValue(song.Path, false);
                }
                else
                {
                    key.SetValue(song.Path, song.Path);
                }
            }
            song.IsFavourite = !song.IsFavourite;
            ApplyFilter();
        }

        private bool IsFavourite(string path)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(SettingsKey + @"\Favourites"))
            {
                return key != null && key.GetValue(path) != null;
            }
        }

        private static bool IsPowerPointFile(string path)
        {
            string extension = Path.GetExtension(path);
            return extension.Equals(".ppt", StringComparison.OrdinalIgnoreCase)
                || extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase);
        }

        private static string Normalize(string value)
        {
            return Regex.Replace(value ?? "", @"\s+", " ").Trim().ToLowerInvariant();
        }

        private static string GetSetting(string name)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(SettingsKey))
            {
                return key == null ? null : key.GetValue(name) as string;
            }
        }

        private static void SetSetting(string name, string value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(SettingsKey))
            {
                key.SetValue(name, value);
            }
        }

        private void UpdateButtons()
        {
            bool hasSelection = songList.SelectedItems.Count > 0;
            insertButton.Enabled = hasSelection && !inserting;
            favouriteButton.Enabled = hasSelection && !inserting;
            if (hasSelection)
            {
                var song = songList.SelectedItems[0].Tag as SongFile;
                favouriteButton.Text = song != null && song.IsFavourite ? "Unfavourite" : "Favourite";
            }
        }

        private sealed class SongFile
        {
            public SongFile(string path, bool isFavourite)
            {
                Path = path;
                Title = System.IO.Path.GetFileNameWithoutExtension(path);
                IsFavourite = isFavourite;
            }

            public string Path { get; private set; }
            public string Title { get; private set; }
            public bool IsFavourite { get; set; }
        }
    }
}
