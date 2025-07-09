# MS Teams Auto Status Userscript

## About

This project is a userscript to automatically switch the user status depending on time of day and day of week.

The userscript will interact with the web version of MS Teams.
An editable config menu will be available to the user to adjust settings as needed.
Constants for relevant CSS selectors are at the top of the script for easy corrections as the script needs to change to reflect the ever-evolving web interface.

### Examples

- 8 AM every weekday: set status to "Available" for working hours start
- 5 PM every weekday: set status to "Appear offline" for working hours end

## Technical Details

- **Match URL**: `https://teams.microsoft.com/v2/*`
- **Userscript Manager**: Tampermonkey, Greasemonkey, or compatible
- **Platform**: Web browsers with userscript support

## Core Implementation

### Main Selectors

The script uses these CSS selectors to interact with Teams UI:

```javascript
const SELECTORS = {
    ACCOUNT_MANAGER: 'div[data-tid="me-control-avatar-presence"]',
    STATUS_MENU_ITEM: 'div[data-tid="set-presence-status-menu-item"]',
    STATUS_OPTIONS: 'div[data-tid^="me_control_presence_availability_"]',
    PRESENCE_INDICATOR: '[data-tid="presence-indicator"]',
    UI_ALERTS: 'div.ui-alert'
};
```

### Status Change Workflow

```javascript
// Open Account Manager Menu (if not already open)
const account_manager = document.querySelector('div[data-tid="me-control-avatar-presence"]');
account_manager.click();

// Click on status to expand context submenu
// Note: Wait for selector to ensure element is available
const status = account_manager.querySelector('div[data-tid="set-presence-status-menu-item"]');
status.click();

// Get list of statuses that can be set
// Note: Wait for selectors to load
const list_of_statuses = status.querySelectorAll('div[data-tid^="me_control_presence_availability_"]');

// Get label for a single status option
const list_status_label = list_of_statuses[0].textContent;

// Set Status
list_of_statuses[0].click();

// Get Current Status
const current_status = account_manager.getAttribute("aria-label");
```

## Additional Features

### Alert Hiding

Automatically hide Teams UI notification popups:

```javascript
// Hide all alerts every second
// Note: A MutationObserver might be more efficient than setInterval
setInterval(() => {
    document.querySelectorAll('div.ui-alert').forEach(alert => {
        alert.style.display = 'none';
    });
}, 1000);
```

### Improved Implementation (MutationObserver)

```javascript
// More efficient approach using MutationObserver
const alertObserver = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
        mutation.addedNodes.forEach((node) => {
            if (node.nodeType === 1 && node.matches('div.ui-alert')) {
                node.style.display = 'none';
            }
        });
    });
});

// Start observing
alertObserver.observe(document.body, {
    childList: true,
    subtree: true
});
```

## Development Notes

- **Selectors**: Teams frequently updates their UI, so selectors may need periodic updates
- **Timing**: Use `waitForElement` functions with timeouts for reliable element detection
- **Error Handling**: Implement try-catch blocks around DOM interactions
- **Debugging**: Use console logging with debug mode toggle for troubleshooting
- **Storage**: Use `GM_setValue`/`GM_getValue` for persistent configuration storage

## Testing

- Test all status types: Available, Busy, Do not disturb, Be right back, Appear offline
- Verify schedule triggers work within the 2-minute window
- Test alert hiding functionality
- Check console for debug messages and errors
- Validate configuration persistence across page reloads
