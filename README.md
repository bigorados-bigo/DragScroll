# DragScroll - Advanced Mouse Scrolling for Windows

![AutoHotkey](https://img.shields.io/badge/AutoHotkey-v1.1+-blue.svg)
![Windows](https://img.shields.io/badge/Windows-7%2B-green.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)

A powerful AutoHotkey script that transforms your mouse into a drag-to-scroll tool, similar to what you'd find on touchscreens or tablets. Hold down a configurable mouse button and drag to scroll in any direction.

## ✨ Features

### Core Functionality
- **Drag Scrolling**: Hold middle mouse button (or any configured button) and drag to scroll
- **Directional Control**: Both vertical and horizontal scrolling modes
- **High-Resolution Scrolling**: Smooth, precise wheel events with configurable sensitivity
- **Speed Control**: Adjustable speed multiplier for different scrolling preferences

### Advanced Configuration
- **GUI Settings Panel**: Easy-to-use configuration interface
- **Button Capture**: Interactive button selection for activation key
- **Process Exclusions**: Disable drag scrolling for specific applications
- **Top Guard Zones**: Prevent accidental scrolling near window title bars
- **Persistent Settings**: Configuration saved to INI file

### Smart Behavior
- **Process Detection**: Automatically disable for excluded applications
- **Fallback Clicks**: Single clicks still work when no dragging occurs
- **Performance Optimized**: Configurable scan intervals and batch processing
- **Buffer Management**: Smooth scrolling with delta accumulation

## 🚀 Quick Start

1. **Download**: Clone this repository or download `DragScroll.ahk`
2. **Install**: Requires [AutoHotkey v1.1+](https://www.autohotkey.com/)
3. **Run**: Double-click `DragScroll.ahk` or compile to `.exe`
4. **Configure**: Right-click tray icon → Settings

## 🎮 Usage

### Basic Operation
1. Hold down the middle mouse button (default)
2. Drag in any direction to scroll
3. Release to stop scrolling
4. Single clicks still work normally

### Configuration Options

| Setting | Description | Default |
|---------|-------------|---------|
| **Activation Button** | Mouse button to trigger drag scrolling | Middle Button |
| **Invert Direction** | Reverse scroll direction | Disabled |
| **Horizontal Mode** | Switch to horizontal panning | Disabled |
| **Speed Multiplier** | Adjust scrolling speed | 1.0 |
| **Wheel Sensitivity** | Delta per pixel moved | 12.0 |
| **Max Wheel Step** | Maximum scroll step size | 480 |
| **Scan Interval** | Polling frequency (ms) | 20 |

### Process Exclusions
Add applications where drag scrolling should be disabled:
```
notepad.exe
photoshop.exe
game.exe
```

### Top Guard Zones
Prevent scrolling near title bars:
```
OneCommander.exe:60
chrome.exe:40
```
Format: `process.exe:height_in_pixels`

## 🔧 Advanced Features

### Supported Mouse Buttons
- Left Button
- Right Button  
- Middle Button (default)
- XButton1 (side button)
- XButton2 (side button)

### Smart Detection
- **Process-based exclusions**: Automatically disabled for specified applications
- **Title bar protection**: Configurable guard zones prevent accidental scrolling
- **Fallback behavior**: Preserves normal click functionality

### Performance Tuning
- Configurable scan intervals (5-1000ms)
- Wheel event batching and clamping
- Efficient delta accumulation system

## 📁 Files

- `DragScroll.ahk` - Main script file
- `mouse-scroll.ini` - Configuration file (auto-generated)
- `README.md` - This documentation

## 🛠️ Development

### Requirements
- AutoHotkey v1.1 or later
- Windows 7 or later

### Building
To compile to executable:
```
Right-click DragScroll.ahk → Compile Script
```

### Configuration Format
Settings are stored in `mouse-scroll.ini`:
```ini
[Settings]
Swap=0
Horizontal=0
SpeedMultiplier=1.0
WheelSensitivity=12.0
WheelMaxStep=480
ActivationButton=MButton
ExcludedProcesses=notepad.exe,game.exe
TopGuardZones=OneCommander.exe:60
ScanInterval=20
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

### Ideas for Enhancement
- Multi-monitor support
- Gesture recognition
- Custom acceleration curves
- Per-application settings
- Touchpad integration

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Original concept by Mikhail V. (2021)
- Enhanced with GUI configuration and advanced features
- Inspired by touchscreen scrolling behavior

## 💡 Tips & Tricks

### Performance Optimization
- Lower scan intervals (5-10ms) for ultra-smooth scrolling
- Increase wheel sensitivity for faster scrolling
- Use process exclusions for resource-intensive applications

### Application-Specific Setup
- **Browsers**: Works great with default settings
- **Code Editors**: Consider horizontal mode for wide files
- **Image Viewers**: Try inverted direction for natural feel
- **Games**: Add to exclusions list to prevent interference

### Troubleshooting
- **Not working in app**: Add to excluded processes
- **Too sensitive**: Lower wheel sensitivity or speed multiplier  
- **Clicks not working**: Ensure single clicks without dragging
- **Performance issues**: Increase scan interval

---

**Made with ❤️ for productivity enthusiasts**

*Transform your mouse into a precision scrolling tool!*