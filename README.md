# Multi Serial Monitor

A Windows-based GUI application for monitoring multiple serial ports simultaneously. Displays incoming data with timestamps, color-coded by port, and provides features like export to HTML, pause/resume, and baud rate selection.

## Features

- **Multi-Port Monitoring**: Monitor up to 255 serial ports at once
- **Color Coding**: Each port has a customizable color for easy identification
- **Timestamps**: All messages include timestamps in HH:MM:SS format
- **Baud Rate Selection**: Supports common baud rates from 1200 to 921600
- **Export to HTML**: Save the monitored data as a colorized HTML file
- **Pause/Resume**: Temporarily pause scrolling to review data
- **Clear Display**: Clear the monitor window
- **Refresh Ports**: Dynamically refresh available serial ports
- **Status Bar**: Shows current monitoring status and port count

## Screenshot

![Multi Serial Monitor](https://raw.githubusercontent.com/nsssim/multiserial/refs/heads/main/screenshot.jpg)

## Requirements

- Windows operating system
- MinGW-w64 or similar GCC toolchain for Windows
- Windows SDK (for resource compilation)

## Building

### Prerequisites
- MinGW-w64 GCC compiler
- Windows SDK (for resource compilation with windres)

### Using Makefile
1. Ensure MinGW-w64 is installed and in your PATH
2. Run `make` in the project directory:

```bash
make
```

This compiles the resource file (`SerialMonitor.rc`) and links the executable (`SerialMonitor.exe`).

### Manual Compilation
If you prefer manual compilation:

1. Compile the resource file:
```bash
windres SerialMonitor.rc -o SerialMonitor_res.o
```

2. Compile and link the main program:
```bash
g++ -o SerialMonitor.exe main.cpp SerialMonitor_res.o -mwindows -lcomctl32 -lriched20 -lgdi32 -luser32 -lkernel32 -ladvapi32
```

## Usage

1. Run `SerialMonitor.exe`
2. Select the desired baud rate from the dropdown (default: 115200)
3. Click the play button (▶) to start monitoring
4. Data from connected serial ports will appear in the main window with timestamps and colors
5. Use the color palette buttons to change port colors
6. Click the export button to save data as HTML
7. Use pause/resume to control scrolling
8. Click the stop button (■) to stop monitoring

## Controls

- **Refresh**: Reload available serial ports
- **Port Selector**: Choose which port's color to modify
- **Color Palette**: 10 color options for port identification
- **Stop/Start**: Toggle monitoring on/off
- **Export**: Save current data to `output.html`
- **Clear**: Clear the display
- **Pause**: Pause/resume auto-scrolling
- **Baud Rate**: Select communication speed

## License

This project is licensed under the GNU General Public License v3.0. See the (https://www.gnu.org/licenses/gpl-3.0.en.html) file for details.

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.
