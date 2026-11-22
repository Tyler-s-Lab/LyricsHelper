
<div align="center">

<img src="LyricsHelper-256.png" width="128px" alt="Logotype">

#  Lyrics Helper

Lyrics Helper is a utility for handling ruby text in document files,

especially useful for Japanese lyrics.

<div align="left">


## What it can

- Cleaning ruby text.　
- Stripping out ruby text.

<div align="center">
<figure>
    <img src="doc/ss2026-06-16-210750.png" width="420px" alt="Lyrics with ruby superfluous">
    <figcaption>Original</figcaption>
</figure>
<figure>
    <img src="doc/ss2026-06-16-210821.png" width="420px" alt="Lyrics with ruby just fitting">
    <figcaption>Cleaned</figcaption>
</figure>
<figure>
    <img src="doc/ss2026-06-16-210849.png" width="420px" alt="Lyrics without ruby base">
    <figcaption>Stripped</figcaption>
</figure>
<div align="left">

## How it behaves

### Stripping

Stripping normally replaces the ruby base with its reading.

### Cleaning

Cleaning only rewrites ruby annotations.

## Input support

- Microsoft® Word `.docx` packages and Flat OPC `.xml` packages (the Word 2003 XML Package format, not arbitrary WordprocessingML).

## Resource limits

Documents are treated as untrusted. The application applies the following bounds (see `LyricsHelper/DocumentLimits.cs`):

- at most 4096 zip entries, 64 Word story parts, and 1024 files per dropped directory;
- at most 32 MiB per XML story part, 128 MiB total decompressed content, and 32 million XML characters per part;
- at most 2,000,000 generated XML nodes per Cleaning file;
- at most 512 characters of ruby base and 4096 characters of ruby reading per annotation;
- at most 1,000,000 characters of extracted/displayed text.

## What it can't (yet)

- To *generate* ruby text.
- To support multilingual.
- To resolve ruby colors inherited from Word styles/themes (direct run colors are checked).
