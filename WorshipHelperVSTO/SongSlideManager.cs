using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace WorshipHelperVSTO
{
    public enum SongSlideTheme
    {
        Dark,
        Light,
        Blue
    }

    public class SongSlideManager
    {
        private readonly Application app = Globals.ThisAddIn.Application;

        public void CreateSlides(string title, string lyrics, SongSlideTheme theme)
        {
            if (app.ActivePresentation == null)
            {
                throw new InvalidOperationException("Open a PowerPoint presentation before creating song slides.");
            }

            var paragraphs = ParseParagraphs(lyrics);
            if (paragraphs.Count == 0)
            {
                throw new InvalidOperationException("The lyrics do not contain any text.");
            }

            var presentation = app.ActivePresentation;
            var insertIndex = new SelectionManager().GetNextSlideIndex();
            insertIndex = Math.Max(1, Math.Min(insertIndex, presentation.Slides.Count + 1));

            AddTextSlide(presentation, ref insertIndex, title, true, theme);
            foreach (var paragraph in paragraphs)
            {
                AddTextSlide(presentation, ref insertIndex, paragraph, false, theme);
            }
        }

        public static IList<string> ParseParagraphs(string lyrics)
        {
            var normalizedLyrics = lyrics
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');

            return SplitParagraphs(normalizedLyrics)
                .Select(paragraph => string.Join(
                    Environment.NewLine,
                    paragraph.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(line => line.Trim())
                        .Where(line => line.Length > 0)))
                .Where(paragraph => paragraph.Length > 0)
                .ToList();
        }

        private static IEnumerable<string> SplitParagraphs(string lyrics)
        {
            return Regex.Split(lyrics, @"\n\s*\n")
                .Where(paragraph => paragraph.Trim().Length > 0);
        }

        private static void AddTextSlide(
            Presentation presentation,
            ref int insertIndex,
            string text,
            bool isTitle,
            SongSlideTheme theme)
        {
            var slide = presentation.Slides.Add(insertIndex, PpSlideLayout.ppLayoutBlank);
            insertIndex++;

            var colors = GetColors(theme);
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.Solid();
            slide.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(colors.Background);

            var slideWidth = presentation.PageSetup.SlideWidth;
            var slideHeight = presentation.PageSetup.SlideHeight;
            var textBoxWidth = slideWidth - 120;
            var fontSize = GetFontSize(text, isTitle);
            var lineCount = Math.Max(1, text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length);
            var textBoxHeight = Math.Max(180F, lineCount * fontSize * 1.35F + 20F);
            textBoxHeight = Math.Min(textBoxHeight, slideHeight - 120);
            var textBox = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                (slideWidth - textBoxWidth) / 2,
                (slideHeight - textBoxHeight) / 2,
                textBoxWidth,
                textBoxHeight);
            var textRange = textBox.TextFrame.TextRange;
            textRange.Text = text;
            textRange.Font.Name = "Nirmala UI";
            textRange.Font.Size = fontSize;
            textRange.Font.Bold = MsoTriState.msoTrue;
            textRange.Font.Color.RGB = ColorTranslator.ToOle(colors.Foreground);
            textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            textBox.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            textBox.TextFrame.WordWrap = MsoTriState.msoTrue;
            textBox.TextFrame.MarginLeft = 10;
            textBox.TextFrame.MarginRight = 10;
            textBox.TextFrame.MarginTop = 10;
            textBox.TextFrame.MarginBottom = 10;
            textBox.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeNone;
            textBox.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            textBox.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
            textBox.Line.Visible = MsoTriState.msoFalse;
        }

        private static float GetFontSize(string text, bool isTitle)
        {
            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var lineCount = Math.Max(1, lines.Length);
            var longestLineLength = lines.Length == 0 ? 0 : lines.Max(line => line.Length);
            var size = isTitle ? 40F : 40F;

            if (lineCount > 4)
            {
                size -= (lineCount - 4) * 3F;
            }

            if (longestLineLength > 42)
            {
                size -= (longestLineLength - 42) * 0.2F;
            }

            return Math.Max(22F, Math.Min(40F, size));
        }

        private static SlideColors GetColors(SongSlideTheme theme)
        {
            switch (theme)
            {
                case SongSlideTheme.Light:
                    return new SlideColors(Color.White, Color.FromArgb(30, 30, 30));
                case SongSlideTheme.Blue:
                    return new SlideColors(Color.FromArgb(18, 48, 92), Color.White);
                default:
                    return new SlideColors(Color.FromArgb(20, 20, 20), Color.White);
            }
        }

        private struct SlideColors
        {
            public SlideColors(Color background, Color foreground)
            {
                Background = background;
                Foreground = foreground;
            }

            public Color Background { get; }
            public Color Foreground { get; }
        }
    }
}
