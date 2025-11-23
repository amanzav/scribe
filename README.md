# Scribe

A tiny PowerShell daemon that watches your Downloads folder, extracts the URL a file came from, and auto-sorts it into the correct course folder.

## Features

- Watches Downloads live
- Extracts HostUrl / ReferrerUrl via Zone.Identifier
- Matches course IDs to folders
- Auto-creates the folder if missing
- Moves files cleanly with collision handling

## Usage

```powershell
.\scribe.ps1 -MonitorFolder "C:\Users\<username>\Downloads"
```
