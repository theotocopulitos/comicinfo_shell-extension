# ComicInfo Shell Extension - Windows 11 Installation Guide

This shell extension displays comic book information from CBZ files by reading the embedded ComicInfo.xml metadata.

## Windows 11 Compatibility

This extension has been updated for Windows 11 compatibility with the following improvements:

- **Updated to .NET Framework 4.8** for better Windows 11 support
- **Added x64 platform configuration** to ensure proper bitness matching
- **Added application manifest** with Windows 11 compatibility declarations
- **Improved error handling** to prevent crashes from malformed comic files
- **Enhanced DPI awareness** for better display on high-DPI screens

## Installation

### Prerequisites
- Windows 11 (64-bit recommended)
- .NET Framework 4.8 (usually pre-installed on Windows 11)

### Manual Installation Steps

1. **Build the extension** (requires Visual Studio or .NET Framework SDK):
   ```
   msbuild GetComicInfo.sln /p:Configuration=Release /p:Platform=x64
   ```

2. **Register the shell extension** (run as Administrator):
   ```
   regsvr32 "path\to\GetComicInfo.dll"
   ```

3. **Alternative registration using ServerManager.exe**:
   - Navigate to the build output folder
   - Right-click on `ServerManager.exe` and "Run as administrator"
   - Use the GUI to install the shell extension

### Troubleshooting

- **Extension doesn't appear**: Ensure you're using the x64 build on 64-bit Windows 11
- **Access denied errors**: Run registration commands as Administrator
- **Missing .NET Framework**: Install .NET Framework 4.8 from Microsoft

### Uninstallation

Run as Administrator:
```
regsvr32 /u "path\to\GetComicInfo.dll"
```

## Usage

1. Right-click on any `.cbz` file in Windows Explorer
2. Select "Comic Info..." from the context menu
3. View the comic metadata in the popup dialog

## Supported CBZ Files

The extension works with CBZ (Comic Book ZIP) files that contain a `ComicInfo.xml` file with metadata including:
- Series name and volume
- Issue number and title
- Publication date
- Writer and artist information
- Story summary