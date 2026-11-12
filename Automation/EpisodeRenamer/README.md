
# Episode Renamer

PowerShell script to normalize TV show episode filenames.

## Scripts available

### [EpisodeRenamer.ps1](EpisodeRenamer.ps1)

A Windows Forms application that scans a selected folder and renames TV show episode files using a consistent `SxxExx` naming format.

#### Requirements

- Windows.
- Windows PowerShell 5.1 or later.
- .NET Framework with Windows Forms support.

#### Features

- Graphical user interface.
- Folder selection or drag and drop.
- Recursive folder scanning (optional).
- Support for `S01E01` and `1x02` episode naming patterns.
- Filename normalization to `SxxExx`.
- Preview mode enabled by default.
- Automatic handling of filename conflicts.
- Rename operation with undo support during the current session.
- Operation log with export capability.
- Progress indicator and stop request.

#### Usage

Run the script:

```powershell
.\EpisodeRenamer.ps1
```

Select a folder containing episode files, or drag and drop a folder into the application.

Review the proposed changes in preview mode. Disable preview mode to apply the renaming operations.

#### Supported formats

The script recognizes the following episode patterns:

- `Show.S01E02.1080p.mkv` → `S01E02.mkv`
- `Show.1x02.1080p.mkv` → `S01E02.mkv`

By default, the script processes the following extensions:

- MKV
- MP4
- AVI
- MOV

Leaving the extension filter empty allows all file extensions to be processed.
