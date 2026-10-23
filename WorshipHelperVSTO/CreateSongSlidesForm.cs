using System;
using System.Drawing;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public class CreateSongSlidesForm : Form
    {
        private readonly TextBox titleTextBox;
        private readonly TextBox lyricsTextBox;
        private readonly ComboBox themeComboBox;

        public CreateSongSlidesForm()
        {
            Text = "Create Song Slides";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(620, 560);
            Size = new Size(720, 620);

            var titleLabel = new Label
            {
                AutoSize = true,
                Location = new Point(20, 20),
                Text = "Song title"
            };
            titleTextBox = new TextBox
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Location = new Point(20, 42),
                Size = new Size(660, 23)
            };

            var lyricsLabel = new Label
            {
                AutoSize = true,
                Location = new Point(20, 82),
                Text = "Lyrics (leave a blank line between slides)"
            };
            lyricsTextBox = new TextBox
            {
                AcceptsReturn = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 10F),
                Location = new Point(20, 104),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Size = new Size(660, 390),
                WordWrap = true
            };

            var themeLabel = new Label
            {
                AutoSize = true,
                Location = new Point(20, 510),
                Text = "Theme"
            };
            themeComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(70, 507),
                Size = new Size(220, 23)
            };
            themeComboBox.Items.AddRange(new object[]
            {
                SongSlideTheme.Dark,
                SongSlideTheme.Light,
                SongSlideTheme.Blue
            });
            themeComboBox.SelectedIndex = 0;

            var cancelButton = new Button
            {
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.Cancel,
                Location = new Point(500, 507),
                Size = new Size(85, 28),
                Text = "Cancel"
            };
            var createButton = new Button
            {
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(595, 507),
                Size = new Size(85, 28),
                Text = "Create"
            };
            createButton.Click += CreateButton_Click;

            Controls.Add(titleLabel);
            Controls.Add(titleTextBox);
            Controls.Add(lyricsLabel);
            Controls.Add(lyricsTextBox);
            Controls.Add(themeLabel);
            Controls.Add(themeComboBox);
            Controls.Add(cancelButton);
            Controls.Add(createButton);
            AcceptButton = createButton;
            CancelButton = cancelButton;
        }

        private void CreateButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(titleTextBox.Text))
            {
                MessageBox.Show("Enter a song title.", "Create Song Slides", MessageBoxButtons.OK, MessageBoxIcon.Information);
                titleTextBox.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(lyricsTextBox.Text))
            {
                MessageBox.Show("Paste the song lyrics.", "Create Song Slides", MessageBoxButtons.OK, MessageBoxIcon.Information);
                lyricsTextBox.Focus();
                return;
            }

            try
            {
                new SongSlideManager().CreateSlides(
                    titleTextBox.Text.Trim(),
                    lyricsTextBox.Text,
                    (SongSlideTheme)themeComboBox.SelectedItem);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    $"An error occurred while creating the song slides.\n\n{exception.Message}",
                    "Create Song Slides",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
