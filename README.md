# VMMC plugin for Microsoft PowerPoint

VMMC helps build Sunday morning presentations by adding songs and Bible readings.

VMMC is a desktop PowerPoint add-in. It can insert existing song presentations with their original slide formatting and add formatted Bible passages.

You must have PowerPoint 2013 or later installed.

## Creating song slides

Use **VMMC > Create Song** to generate a presentation section from pasted lyrics. Enter the song title, paste the lyrics, choose a theme, and click **Create**. The add-in creates a title slide followed by one lyric slide for each paragraph. Separate paragraphs with a blank line; lines within a paragraph remain together on the same slide.

The first version uses three local themes: dark, light, and blue. Lyrics and title text are bold, centered, and automatically fitted inside a generous slide-sized text box for projection readability. Slide creation is performed locally with PowerPoint automation, so lyrics are not sent to a web service. Very long paragraphs should still be reviewed after creation because they may need to be split manually for projection.

## Song library

Use **VMMC > Song Library** to choose the folder that contains your `.ppt` and `.pptx` song files. VMMC stores that folder for future PowerPoint sessions. Click **Refresh** after adding, removing, or renaming files.

The library indexes filenames only, so it does not open every deck while searching and remains suitable for large collections. Search is case-insensitive and ignores repeated whitespace. Focus the search box, use Up/Down to select a result, and press Enter to insert it. Escape clears the search. Double-clicking a result or using **Insert** has the same effect.

Favorites are stored in the existing per-user Windows registry location and can be toggled from the library. Existing Favorites on the ribbon continue to work.

### Slide insertion limitations

This is a VSTO desktop add-in, so it can access the local folder and use PowerPoint COM automation; a companion web service is not required. Insertion opens the source deck, copies its slides, and uses PowerPoint's **Paste Source Formatting** command. This preserves slide order and normally preserves the source deck's visual formatting. PowerPoint's support for particular embedded objects, fonts, and legacy `.ppt` content can still vary by Office version. Source files are never modified.

## Adding bible translations

VMMC comes pre-packaged with the NASB and ESV translations.

You can add more by dropping them in
`C:\Program Files (x86)\VMMC\Data\Bibles` (assuming the default install location). Bibles are in
[OpenSong format](https://opensong.org/development/file-formats/#bible-file-format-specification) (.xmm files). A limited selection can be downloaded from the [OpenSong website](https://opensong.org/downloads/#bibles).

## License

VMMC is free to use for all charitable organisations, and can also be freely modified and shared for any charitable purpose. It is designed for use by churches.

Bible data packaged with VMMC must be used only in accordance with the publisher's license agreement.
