# DOCX Exporter: Obsidian Note Export Plugin

[中文版](README_CN.md)

## Introduction
DOCX Exporter is a plugin designed for Obsidian, aiming to help users easily export note content into DOCX format.

The biggest advantage of this plugin is its **zero external dependencies**. It does not require installing additional external tools like Pandoc, allowing it to run seamlessly on all platforms supported by Obsidian (including desktop, mobile, and iPad), providing you with a consistent export experience.

## Key Features
* **Cross-Platform Support**: Works on Windows, macOS, Linux, iOS, and Android without the need for additional software.
* **Rich Text Export**: Supports exporting various Markdown formats, including headings, bold, italics, lists, hyperlinks, and code blocks.
* **Image Support**: Can export local and web images from notes, and automatically handles image formats and dimensions.
* **Compatibility**: The generated DOCX files have good compatibility with Microsoft Word.

## Known Issues and Workarounds
Currently, we have found that when the exported DOCX file is opened with **Apple Pages** or the **system's native Preview app**, the formatting may be incorrect.

* **Cause**: This is typically because these applications cannot correctly handle certain compatibility settings and font metadata within the DOCX file.
* **Workaround**: Simply open the file with **Microsoft Word** (desktop or mobile version) and save it again. Word will automatically fix and add the necessary compatibility information, after which the file can be opened normally in Pages or other applications.

## Installation
You can install the DOCX Exporter directly from the Obsidian Community Plugins market.
* Open Obsidian **Settings**.
* Click **Community plugins**, and then turn off **Safe mode**.
* In the plugin list, search for **"DOCX Exporter"**.
* Click **Install**, then **Enable** the plugin.

## How to Use
1.  Open the note you want to export.
2.  Click the export icon in the left sidebar, or run the command “**Export current note to DOCX**” via the command palette (`Ctrl/Cmd + P`).
3.  The plugin will generate a DOCX file and save it in the same folder as the current note.
