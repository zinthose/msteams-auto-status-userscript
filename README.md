# MS Teams Auto Status Userscript

![Version](https://img.shields.io/badge/version-1.1.0-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-Web-orange.svg)

Automatically switch your Microsoft Teams status based on time of day and day of week. Never forget to set your status again with this intelligent userscript that runs in your browser.

## 🚀 Installation

### Step 1: Install a Userscript Manager

1. Install [Tampermonkey](https://www.tampermonkey.net/) for your browser
2. Or install [Greasemonkey](https://www.greasespot.net/) for Firefox

### Step 2: Install the Script

1. Click on the Tampermonkey extension icon in your browser
2. Select "Create a new script..."
3. Replace the default content with the code from [`msteams-auto-status.user.js`](./msteams-auto-status.user.js)
4. Press `Ctrl+S` (or `Cmd+S` on Mac) to save
5. Navigate to [Microsoft Teams](https://teams.microsoft.com/v2/)

### Step 3: Configure Your Schedules

1. Look for the **⚙️ Auto Status** button in the top-left corner of Teams
2. Click it to open the configuration panel
3. Set up your schedules (see [Usage Examples](#-usage-examples) below)

## 🌟 Features

- **⏰ Automatic Status Switching**: Set your status to change automatically based on schedules you define
- **📅 Day-Based Scheduling**: Configure different schedules for different days of the week
- **🎛️ Intuitive Configuration UI**: Easy-to-use dark-themed interface for managing all settings
- **🧪 Built-in Testing Tools**: Test status changes and schedule logic without waiting
- **🚨 Alert Hiding**: Automatically hide Teams notification popups (optional)
- **🐛 Debug Mode**: Verbose logging for troubleshooting
- **📊 Current Status Display**: See your current Teams status in real-time
- **🔄 Real-time Updates**: Changes apply immediately without page refresh

## 📋 Requirements

- **Browser**: Chrome, Firefox, Edge, or any browser that supports userscripts
- **Userscript Manager**: [Tampermonkey](https://www.tampermonkey.net/) (recommended) or [Greasemonkey](https://www.greasespot.net/)
- **Microsoft Teams**: Web version at `https://teams.microsoft.com/v2/`

## 💡 Usage Examples

### Example 1: Standard Work Hours

```yaml
Name: Work Start
Time: 08:00
Days: Mon, Tue, Wed, Thu, Fri
Status: Available
```

```yaml
Name: Work End
Time: 17:00
Days: Mon, Tue, Wed, Thu, Fri
Status: Appear offline
```

### Example 2: Lunch Break

```yaml
Name: Lunch Break
Time: 12:00
Days: Mon, Tue, Wed, Thu, Fri
Status: Be right back
```

```yaml
Name: Back from Lunch
Time: 13:00
Days: Mon, Tue, Wed, Thu, Fri
Status: Available
```

### Example 3: Meeting Block

```yaml
Name: Focus Time
Time: 14:00
Days: Mon, Wed, Fri
Status: Do not disturb
```

## 🧪 Testing Features

The configuration panel includes several testing tools:

- **Status Test Buttons**: Manually test each status type
- **Alert Hiding Test**: Test the alert hiding functionality
- **Schedule Check**: Manually trigger schedule evaluation
- **Clear History**: Reset schedule history for testing
- **Current Status Display**: Real-time status with auto-refresh

## 🐛 Troubleshooting

### Script Not Working?

1. **Check if Teams is loaded**: Wait for Teams to fully load before the script activates
2. **Enable Debug Mode**: Turn on debug mode in settings and check the browser console
3. **Verify Selectors**: Teams UI changes frequently; selectors may need updates
4. **Check Schedule Times**: Ensure schedules are within 2 minutes of current time to trigger

### Console Debugging

Open browser console (`F12`) and look for messages starting with:

- `[MS Teams Auto Status]` - General information
- `[MS Teams Auto Status DEBUG]` - Detailed debug info (when debug mode enabled)
- `[MS Teams Auto Status ERROR]` - Error messages

### Common Issues

- **Status not changing**: Check if schedule is enabled and current time matches
- **UI not appearing**: Ensure Teams page is fully loaded
- **Permissions errors**: Some browsers may block certain DOM operations

## 🔧 Customization

### Updating Selectors

If Teams changes its interface, update the `SELECTORS` object at the top of the script:

```javascript
const SELECTORS = {
    ACCOUNT_MANAGER: 'div[data-tid="me-control-avatar-presence"]',
    STATUS_MENU_ITEM: 'div[data-tid="set-presence-status-menu-item"]',
    STATUS_OPTIONS: 'div[data-tid^="me_control_presence_availability_"]',
    PRESENCE_INDICATOR: '[data-tid="presence-indicator"]',
    UI_ALERTS: 'div.ui-alert'
};
```

### Adding New Status Types

Extend the `STATUS_TYPES` object:

```javascript
const STATUS_TYPES = {
    AVAILABLE: 'Available',
    BUSY: 'Busy',
    // Add new status types here
};
```

## 📁 File Structure

```yaml
userscript_MSTeams_autostatus/
├── README.md                    # This file
├── msteams-auto-status.user.js  # Main userscript
├── description.md               # Development notes and examples
└── LICENSE                      # MIT License
```

## 🤝 Contributing

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Commit** your changes (`git commit -m 'Add amazing feature'`)
4. **Push** to the branch (`git push origin feature/amazing-feature`)
5. **Open** a Pull Request

### Development Notes

- See [`description.md`](./description.md) for development examples and selector samples
- Use the built-in debug mode for testing
- Ensure all UI elements are accessible in dark mode
- Test with different Teams interface updates

## 📝 Changelog

### v1.1.0 (Current)

- ✅ Added alert hiding functionality
- ✅ Enhanced testing interface with alert hiding test
- ✅ Improved error handling and debug logging
- ✅ Added real-time configuration updates

### v1.0.0

- ✅ Initial release with automatic status switching
- ✅ Schedule-based configuration
- ✅ Dark-themed configuration UI
- ✅ Testing and debugging tools
- ✅ Current status display

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This userscript interacts with the Microsoft Teams web interface and may stop working if Microsoft changes their interface. The script is provided as-is and the authors are not responsible for any issues that may arise from its use.

## 🙏 Acknowledgments

- Microsoft Teams for providing a web interface that can be automated
- The userscript community for tools and inspiration
- Contributors and users who provide feedback and improvements

---

## **Made with ❤️ for productivity and automation**

*Star ⭐ this repository if it helps you stay on top of your Teams status!*
