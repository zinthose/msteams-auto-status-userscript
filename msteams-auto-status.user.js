// ==UserScript==
// @name         MS Teams Auto Status
// @namespace    http://tampermonkey.net/
// @version      1.1.0
// @description  Automatically switch MS Teams status based on time of day and day of week
// @author       Dane Jones
// @match        https://teams.microsoft.com/v2/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=teams.microsoft.com
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// @run-at       document-idle

// ==/UserScript==

(function() {
    'use strict';

    // ===== SELECTOR CONSTANTS =====
    // Update these selectors if Teams UI changes
    const SELECTORS = {
        ACCOUNT_MANAGER: 'div[data-tid="me-control-avatar-presence"]',
        STATUS_MENU_ITEM: 'div[data-tid="set-presence-status-menu-item"]',
        STATUS_OPTIONS: 'div[data-tid^="me_control_presence_availability_"]',
        PRESENCE_INDICATOR: '[data-tid="presence-indicator"]',
        UI_ALERTS: 'div.ui-alert'
    };

    // ===== STATUS MAPPINGS =====
    const STATUS_TYPES = {
        AVAILABLE: 'Available',
        BUSY: 'Busy',
        DO_NOT_DISTURB: 'Do not disturb',
        BE_RIGHT_BACK: 'Be right back',
        APPEAR_OFFLINE: 'Appear offline'
    };

    // ===== GLOBAL VARIABLES =====
    let alertHidingInterval = null;

    // ===== DEFAULT CONFIGURATION =====
    const DEFAULT_CONFIG = {
        enabled: true,
        schedules: [
            {
                id: 'work_start',
                name: 'Work Start',
                time: '08:00',
                days: [1, 2, 3, 4, 5], // Monday to Friday
                status: STATUS_TYPES.AVAILABLE,
                enabled: true
            },
            {
                id: 'work_end',
                name: 'Work End',
                time: '17:00',
                days: [1, 2, 3, 4, 5], // Monday to Friday
                status: STATUS_TYPES.APPEAR_OFFLINE,
                enabled: true
            }
        ],
        checkInterval: 60000, // Check every minute
        lastCheck: null,
        debugMode: false,
        hideAlerts: false // Hide Teams UI alerts
    };

    // ===== UTILITY FUNCTIONS =====
    function log(message, ...args) {
        console.log(`[MS Teams Auto Status] ${message}`, ...args);
    }

    function debugLog(message, ...args) {
        const config = getConfig();
        if (config.debugMode) {
            console.debug(`[MS Teams Auto Status DEBUG] ${message}`, ...args);
        }
    }

    function errorLog(message, error, ...args) {
        console.error(`[MS Teams Auto Status ERROR] ${message}`, error, ...args);
    }

    // Monitor for Trusted Types violations
    function setupTrustedTypesMonitoring() {
        debugLog('Setting up Trusted Types monitoring');
        
        // Override innerHTML setter to catch violations
        try {
            const originalInnerHTMLSetter = Element.prototype.__lookupSetter__('innerHTML');
            if (originalInnerHTMLSetter) {
                debugLog('Overriding innerHTML setter for monitoring');
                Object.defineProperty(Element.prototype, 'innerHTML', {
                    set: function(value) {
                        debugLog('innerHTML setter called on element:', this.tagName, this.className, this.id);
                        debugLog('Value being set:', typeof value, value?.length ? value.substring(0, 100) + '...' : value);
                        try {
                            return originalInnerHTMLSetter.call(this, value);
                        } catch (error) {
                            errorLog('innerHTML assignment failed:', error);
                            errorLog('Element details:', {
                                tagName: this.tagName,
                                className: this.className,
                                id: this.id,
                                parent: this.parentElement?.tagName
                            });
                            throw error;
                        }
                    },
                    get: Element.prototype.__lookupGetter__('innerHTML')
                });
            } else {
                debugLog('innerHTML setter not found for overriding');
            }
        } catch (error) {
            errorLog('Error setting up innerHTML monitoring:', error);
        }

        // Monitor for any TrustedHTML policy violations
        if (window.trustedTypes) {
            debugLog('TrustedTypes is available');
            
            // Check what methods are available
            debugLog('TrustedTypes methods available:', Object.getOwnPropertyNames(window.trustedTypes));
            
            // Try to get existing policies if the method exists
            try {
                if (typeof window.trustedTypes.getPolicyNames === 'function') {
                    debugLog('Existing policies:', window.trustedTypes.getPolicyNames());
                } else {
                    debugLog('getPolicyNames method not available');
                }
            } catch (error) {
                debugLog('Could not get policy names:', error.message);
            }
            
            // Try to create a policy for our script
            try {
                const policy = window.trustedTypes.createPolicy('ms-teams-auto-status', {
                    createHTML: (string) => {
                        debugLog('Creating trusted HTML for:', string.substring(0, 100));
                        return string;
                    }
                });
                debugLog('Successfully created TrustedTypes policy');
                window.msTeamsAutoStatusPolicy = policy;
            } catch (error) {
                debugLog('Failed to create TrustedTypes policy (this is normal if policies are restricted):', error.message);
            }
        } else {
            debugLog('TrustedTypes not available in this environment');
        }
    }

    function waitForElement(selector, timeout = 10000) {
        return new Promise((resolve, reject) => {
            debugLog(`Waiting for element: ${selector}`);
            const element = document.querySelector(selector);
            if (element) {
                debugLog(`Element found immediately: ${selector}`);
                resolve(element);
                return;
            }

            const observer = new MutationObserver((mutations, obs) => {
                const element = document.querySelector(selector);
                if (element) {
                    debugLog(`Element found via observer: ${selector}`);
                    obs.disconnect();
                    resolve(element);
                }
            });

            observer.observe(document.body, {
                childList: true,
                subtree: true
            });

            setTimeout(() => {
                observer.disconnect();
                errorLog(`Element not found within timeout: ${selector}`);
                reject(new Error(`Element ${selector} not found within ${timeout}ms`));
            }, timeout);
        });
    }

    function getCurrentTime() {
        const now = new Date();
        return {
            hour: now.getHours(),
            minute: now.getMinutes(),
            day: now.getDay(), // 0 = Sunday, 1 = Monday, etc.
            timeString: now.toTimeString().substr(0, 5) // HH:MM format
        };
    }

    function timeToMinutes(timeString) {
        const [hours, minutes] = timeString.split(':').map(Number);
        return hours * 60 + minutes;
    }

    // ===== CONFIGURATION MANAGEMENT =====
    function getConfig() {
        const config = GM_getValue('teamsAutoStatusConfig', JSON.stringify(DEFAULT_CONFIG));
        return JSON.parse(config);
    }

    function saveConfig(config) {
        GM_setValue('teamsAutoStatusConfig', JSON.stringify(config));
        log('Configuration saved');
    }

    // ===== TEAMS STATUS MANAGEMENT =====
    async function getCurrentStatus() {
        try {
            // Try the account manager aria-label first (from description.md)
            const accountManager = document.querySelector(SELECTORS.ACCOUNT_MANAGER);
            if (accountManager) {
                const ariaLabel = accountManager.getAttribute("aria-label");
                if (ariaLabel) {
                    debugLog('Current status from account manager aria-label:', ariaLabel);
                    return ariaLabel;
                }
            }
            
            // Fallback to presence indicator
            const presenceIndicator = document.querySelector(SELECTORS.PRESENCE_INDICATOR);
            if (presenceIndicator) {
                const status = presenceIndicator.getAttribute('aria-label') || 'Unknown';
                debugLog('Current status from presence indicator:', status);
                return status;
            }
        } catch (error) {
            errorLog('Error getting current status:', error);
        }
        return 'Unknown';
    }

    async function setTeamsStatus(targetStatus) {
        try {
            log(`Attempting to set status to: ${targetStatus}`);

            // Open Account Manager Menu
            const accountManager = await waitForElement(SELECTORS.ACCOUNT_MANAGER, 5000);
            accountManager.click();
            log('Opened account manager menu');

            // Wait a bit for menu to open
            await new Promise(resolve => setTimeout(resolve, 500));

            // Click on status to expand context submenu
            const statusMenuItem = await waitForElement(SELECTORS.STATUS_MENU_ITEM, 5000);
            statusMenuItem.click();
            log('Opened status submenu');

            // Wait for status options to load
            await new Promise(resolve => setTimeout(resolve, 1000));

            // Get list of available status options
            const statusOptions = document.querySelectorAll(SELECTORS.STATUS_OPTIONS);
            log(`Found ${statusOptions.length} status options`);

            // Find and click the target status
            let statusSet = false;
            for (const option of statusOptions) {
                const statusText = option.textContent.trim();
                log(`Checking status option: "${statusText}"`);
                
                if (statusText.toLowerCase().includes(targetStatus.toLowerCase()) || 
                    targetStatus.toLowerCase().includes(statusText.toLowerCase())) {
                    option.click();
                    log(`Successfully set status to: ${statusText}`);
                    statusSet = true;
                    break;
                }
            }

            if (!statusSet) {
                log(`Status "${targetStatus}" not found in available options`);
                // Close any open menus by clicking elsewhere
                document.body.click();
                return false;
            }

            // Close any remaining menus
            await new Promise(resolve => setTimeout(resolve, 500));
            document.body.click();

            return true;

        } catch (error) {
            log('Error setting Teams status:', error);
            // Try to close any open menus
            document.body.click();
            return false;
        }
    }

    // ===== SCHEDULE CHECKING =====
    function shouldChangeStatus(config) {
        const currentTime = getCurrentTime();
        const currentMinutes = currentTime.hour * 60 + currentTime.minute;
        
        if (config.debugMode) {
            log(`Checking schedules for ${currentTime.timeString} on day ${currentTime.day}`);
        }

        for (const schedule of config.schedules) {
            if (!schedule.enabled) {
                debugLog(`Skipping disabled schedule: ${schedule.name}`);
                continue;
            }
            
            // Check if current day is in schedule
            if (!schedule.days.includes(currentTime.day)) {
                debugLog(`Schedule "${schedule.name}" not active on day ${currentTime.day}`);
                continue;
            }
            
            const scheduleMinutes = timeToMinutes(schedule.time);
            
            // Check if we're within a few minutes of the scheduled time
            const timeDiff = Math.abs(currentMinutes - scheduleMinutes);
            
            debugLog(`Schedule "${schedule.name}": time diff = ${timeDiff} minutes (trigger threshold: 2)`);
            
            // If within 2 minutes of scheduled time and we haven't processed this schedule recently
            if (timeDiff <= 2) {
                const lastProcessed = GM_getValue(`lastProcessed_${schedule.id}`, '');
                const todayKey = new Date().toDateString() + '_' + schedule.time;
                
                debugLog(`Schedule "${schedule.name}": lastProcessed = "${lastProcessed}", todayKey = "${todayKey}"`);
                
                if (lastProcessed !== todayKey) {
                    log(`🎯 Schedule "${schedule.name}" triggered at ${schedule.time} → ${schedule.status}`);
                    GM_setValue(`lastProcessed_${schedule.id}`, todayKey);
                    return schedule;
                } else {
                    debugLog(`Schedule "${schedule.name}" already processed today`);
                }
            }
        }
        
        return null;
    }

    async function checkAndUpdateStatus() {
        const config = getConfig();
        
        if (!config.enabled) {
            log('Auto status is disabled');
            return;
        }

        const triggeredSchedule = shouldChangeStatus(config);
        
        if (triggeredSchedule) {
            log(`Executing schedule: ${triggeredSchedule.name} -> ${triggeredSchedule.status}`);
            const success = await setTeamsStatus(triggeredSchedule.status);
            
            if (success) {
                log(`Status successfully changed to ${triggeredSchedule.status}`);
            } else {
                log(`Failed to change status to ${triggeredSchedule.status}`);
            }
        }
    }

    // ===== CONFIG UI =====
    function createConfigUI() {
        debugLog('Creating config UI');
        
        try {
            const config = getConfig();
            debugLog('Config loaded for UI:', config);

            // Create modal structure using DOM methods instead of innerHTML
            debugLog('Creating modal elements');
            const modal = document.createElement('div');
            modal.id = 'teams-auto-status-config';
            
            const backdrop = document.createElement('div');
            backdrop.className = 'modal-backdrop';
            
            const content = document.createElement('div');
            content.className = 'modal-content';
            
            // Create header
            debugLog('Creating header elements');
            const header = document.createElement('div');
            header.className = 'modal-header';
            
            const title = document.createElement('h2');
            title.textContent = 'MS Teams Auto Status Configuration';
            
            const closeBtn = document.createElement('button');
            closeBtn.className = 'close-btn';
            closeBtn.textContent = '×';
            
            header.appendChild(title);
            header.appendChild(closeBtn);
            
            // Create body
            debugLog('Creating body elements');
            const body = document.createElement('div');
            body.className = 'modal-body';
            
            // Enable section
            const enableSection = document.createElement('div');
            enableSection.className = 'config-section';
            
            const enableLabel = document.createElement('label');
            const enableCheckbox = document.createElement('input');
            enableCheckbox.type = 'checkbox';
            enableCheckbox.id = 'enabled';
            enableCheckbox.checked = config.enabled;
            
            enableLabel.appendChild(enableCheckbox);
            enableLabel.appendChild(document.createTextNode(' Enable Auto Status'));
            enableSection.appendChild(enableLabel);
            
            // Debug mode toggle
            const debugLabel = document.createElement('label');
            debugLabel.style.cssText = 'display: block; margin-top: 10px;';
            const debugCheckbox = document.createElement('input');
            debugCheckbox.type = 'checkbox';
            debugCheckbox.id = 'debug-mode';
            debugCheckbox.checked = config.debugMode || false;
            
            debugLabel.appendChild(debugCheckbox);
            debugLabel.appendChild(document.createTextNode(' Enable Debug Mode (verbose console logging)'));
            enableSection.appendChild(debugLabel);
            
            // Alert hiding toggle
            const alertLabel = document.createElement('label');
            alertLabel.style.cssText = 'display: block; margin-top: 10px;';
            const alertCheckbox = document.createElement('input');
            alertCheckbox.type = 'checkbox';
            alertCheckbox.id = 'hide-alerts';
            alertCheckbox.checked = config.hideAlerts || false;
            
            alertLabel.appendChild(alertCheckbox);
            alertLabel.appendChild(document.createTextNode(' Hide Teams UI Alerts (hide notification popups)'));
            enableSection.appendChild(alertLabel);
            
            // Schedules section
            debugLog('Creating schedules section');
            const schedulesSection = document.createElement('div');
            schedulesSection.className = 'config-section';
            
            const schedulesTitle = document.createElement('h3');
            schedulesTitle.textContent = 'Schedules';
            
            const schedulesContainer = document.createElement('div');
            schedulesContainer.id = 'schedules-container';
            
            // Add existing schedules
            debugLog('Adding existing schedules:', config.schedules.length);
            config.schedules.forEach((schedule, index) => {
                debugLog('Creating schedule item:', index, schedule.name);
                const scheduleItem = createScheduleItem(schedule, index);
                schedulesContainer.appendChild(scheduleItem);
            });
            
            const addButton = document.createElement('button');
            addButton.id = 'add-schedule';
            addButton.textContent = 'Add Schedule';
            
            schedulesSection.appendChild(schedulesTitle);
            schedulesSection.appendChild(schedulesContainer);
            schedulesSection.appendChild(addButton);
            
            body.appendChild(enableSection);
            body.appendChild(schedulesSection);
            
            // Test section
            debugLog('Creating test section');
            const testSection = document.createElement('div');
            testSection.className = 'config-section';
            
            const testTitle = document.createElement('h3');
            testTitle.textContent = 'Testing';
            
            const testContainer = document.createElement('div');
            testContainer.style.cssText = `
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
                gap: 10px;
                margin-bottom: 10px;
            `;
            
            // Add test buttons for each status type
            Object.values(STATUS_TYPES).forEach(status => {
                const testBtn = document.createElement('button');
                testBtn.textContent = `Test ${status}`;
                testBtn.style.cssText = `
                    padding: 6px 10px;
                    font-size: 12px;
                    background: #6264A7;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                `;
                testBtn.onclick = async () => {
                    debugLog(`Testing status change to: ${status}`);
                    testBtn.disabled = true;
                    testBtn.textContent = 'Testing...';
                    
                    try {
                        const success = await setTeamsStatus(status);
                        if (success) {
                            testBtn.style.background = '#107C10';
                            testBtn.textContent = '✓ Success';
                            setTimeout(() => {
                                testBtn.style.background = '#6264A7';
                                testBtn.textContent = `Test ${status}`;
                                testBtn.disabled = false;
                            }, 2000);
                        } else {
                            testBtn.style.background = '#D83B01';
                            testBtn.textContent = '✗ Failed';
                            setTimeout(() => {
                                testBtn.style.background = '#6264A7';
                                testBtn.textContent = `Test ${status}`;
                                testBtn.disabled = false;
                            }, 2000);
                        }
                    } catch (error) {
                        errorLog('Test failed:', error);
                        testBtn.style.background = '#D83B01';
                        testBtn.textContent = '✗ Error';
                        setTimeout(() => {
                            testBtn.style.background = '#6264A7';
                            testBtn.textContent = `Test ${status}`;
                            testBtn.disabled = false;
                        }, 2000);
                    }
                };
                testContainer.appendChild(testBtn);
            });
            
            // Add alert hiding test button
            const alertTestBtn = document.createElement('button');
            alertTestBtn.textContent = '🚨 Test Alert Hiding';
            alertTestBtn.style.cssText = `
                padding: 6px 12px;
                background: #6264A7;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 12px;
                margin: 5px 5px 5px 0;
            `;
            alertTestBtn.onclick = async () => {
                try {
                    debugLog('Testing alert hiding');
                    alertTestBtn.disabled = true;
                    alertTestBtn.textContent = 'Testing...';
                    
                    const alertsBefore = document.querySelectorAll(SELECTORS.UI_ALERTS);
                    debugLog(`Found ${alertsBefore.length} alerts before hiding`);
                    
                    hideTeamsAlerts();
                    
                    const alertsAfter = document.querySelectorAll(SELECTORS.UI_ALERTS);
                    const hiddenCount = Array.from(alertsAfter).filter(alert => alert.style.display === 'none').length;
                    
                    alertTestBtn.style.background = '#107C10';
                    alertTestBtn.textContent = `✓ Hidden: ${hiddenCount}`;
                    setTimeout(() => {
                        alertTestBtn.style.background = '#6264A7';
                        alertTestBtn.textContent = '🚨 Test Alert Hiding';
                        alertTestBtn.disabled = false;
                    }, 2000);
                } catch (error) {
                    errorLog('Alert hiding test failed:', error);
                    alertTestBtn.style.background = '#D83B01';
                    alertTestBtn.textContent = '✗ Error';
                    setTimeout(() => {
                        alertTestBtn.style.background = '#6264A7';
                        alertTestBtn.textContent = '🚨 Test Alert Hiding';
                        alertTestBtn.disabled = false;
                    }, 2000);
                }
            };
            testContainer.appendChild(alertTestBtn);
            
            // Add current status display
            const statusDisplay = document.createElement('div');
            statusDisplay.style.cssText = `
                padding: 10px;
                background: #252526;
                border: 1px solid #404040;
                border-radius: 4px;
                margin-bottom: 10px;
                font-size: 14px;
                color: #f0f0f0;
                font-weight: 500;
            `;
            statusDisplay.textContent = 'Current Status: Loading...';
            
            // Function to update current status
            const updateCurrentStatus = async () => {
                try {
                    const status = await getCurrentStatus();
                    const timestamp = new Date().toLocaleTimeString();
                    statusDisplay.textContent = `Current Status: ${status} (Updated: ${timestamp})`;
                    debugLog('Status display updated:', status);
                } catch (error) {
                    statusDisplay.textContent = 'Current Status: Error getting status';
                    errorLog('Error updating status display:', error);
                }
            };
            
            // Update current status initially and every 10 seconds
            updateCurrentStatus();
            const statusUpdateInterval = setInterval(updateCurrentStatus, 10000);
            
            // Add refresh button for manual status check
            const refreshStatusBtn = document.createElement('button');
            refreshStatusBtn.textContent = '🔄 Refresh Status';
            refreshStatusBtn.style.cssText = `
                padding: 6px 12px;
                background: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 12px;
                margin-left: 10px;
                vertical-align: middle;
            `;
            refreshStatusBtn.onclick = () => {
                debugLog('Manual status refresh triggered');
                updateCurrentStatus();
                refreshStatusBtn.textContent = '🔄 Refreshed!';
                setTimeout(() => {
                    refreshStatusBtn.textContent = '🔄 Refresh Status';
                }, 1000);
            };
            
            // Create status container
            const statusContainer = document.createElement('div');
            statusContainer.style.cssText = `
                display: flex;
                align-items: center;
                margin-bottom: 10px;
            `;
            statusContainer.appendChild(statusDisplay);
            statusContainer.appendChild(refreshStatusBtn);
            
            // Add schedule test button
            const scheduleTestBtn = document.createElement('button');
            scheduleTestBtn.textContent = 'Test Schedule Check';
            scheduleTestBtn.style.cssText = `
                padding: 8px 16px;
                background: #8661C5;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                width: 100%;
                margin-bottom: 10px;
            `;
            scheduleTestBtn.onclick = () => {
                debugLog('Manual schedule check triggered');
                checkAndUpdateStatus();
                scheduleTestBtn.textContent = 'Schedule Check Triggered (see console)';
                setTimeout(() => {
                    scheduleTestBtn.textContent = 'Test Schedule Check';
                }, 2000);
            };
            
            // Add clear history button for testing
            const clearHistoryBtn = document.createElement('button');
            clearHistoryBtn.textContent = 'Clear Schedule History (for testing)';
            clearHistoryBtn.style.cssText = `
                padding: 8px 16px;
                background: #D83B01;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                width: 100%;
            `;
            clearHistoryBtn.onclick = () => {
                const config = getConfig();
                config.schedules.forEach(schedule => {
                    GM_setValue(`lastProcessed_${schedule.id}`, '');
                });
                log('Schedule history cleared - schedules can trigger again today');
                clearHistoryBtn.textContent = 'History Cleared!';
                setTimeout(() => {
                    clearHistoryBtn.textContent = 'Clear Schedule History (for testing)';
                }, 2000);
            };
            
            testSection.appendChild(testTitle);
            testSection.appendChild(statusDisplay);
            testSection.appendChild(testContainer);
            testSection.appendChild(scheduleTestBtn);
            testSection.appendChild(clearHistoryBtn);
            
            body.appendChild(testSection);
            
            // Create footer
            debugLog('Creating footer elements');
            const footer = document.createElement('div');
            footer.className = 'modal-footer';
            
            const saveBtn = document.createElement('button');
            saveBtn.id = 'save-config';
            saveBtn.textContent = 'Save';
            
            const cancelBtn = document.createElement('button');
            cancelBtn.id = 'cancel-config';
            cancelBtn.textContent = 'Cancel';
            
            footer.appendChild(saveBtn);
            footer.appendChild(cancelBtn);
            
            // Assemble modal
            debugLog('Assembling modal structure');
            content.appendChild(header);
            content.appendChild(body);
            content.appendChild(footer);
            backdrop.appendChild(content);
            modal.appendChild(backdrop);

            // Add styles
            debugLog('Adding styles');
            GM_addStyle(`
                #teams-auto-status-config {
                    position: fixed;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100%;
                    z-index: 10000;
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                }
                
                .modal-backdrop {
                    background: rgba(0, 0, 0, 0.7);
                    width: 100%;
                    height: 100%;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }
                
                .modal-content {
                    background: #2d2d30;
                    color: #f0f0f0;
                    border-radius: 8px;
                    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.5);
                    max-width: 600px;
                    max-height: 80vh;
                    overflow-y: auto;
                    width: 90%;
                    border: 1px solid #404040;
                }
                
                .modal-header {
                    padding: 20px;
                    border-bottom: 1px solid #404040;
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    background: #1e1e1e;
                }
                
                .modal-header h2 {
                    margin: 0;
                    color: #ffffff;
                    font-weight: 600;
                }
                
                .close-btn {
                    background: none;
                    border: none;
                    font-size: 24px;
                    cursor: pointer;
                    color: #cccccc;
                    padding: 4px 8px;
                    border-radius: 4px;
                    transition: background-color 0.2s;
                }
                
                .close-btn:hover {
                    background: #404040;
                    color: #ffffff;
                }
                
                .modal-body {
                    padding: 20px;
                    background: #2d2d30;
                }
                
                .config-section {
                    margin-bottom: 20px;
                    padding: 15px;
                    background: #1e1e1e;
                    border-radius: 6px;
                    border: 1px solid #404040;
                }
                
                .config-section h3 {
                    margin: 0 0 15px 0;
                    color: #ffffff;
                    font-weight: 600;
                    font-size: 16px;
                }
                
                .config-section label {
                    color: #f0f0f0;
                    font-size: 14px;
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    margin-bottom: 8px;
                    cursor: pointer;
                }
                
                .config-section input[type="checkbox"] {
                    margin: 0;
                    transform: scale(1.1);
                }
                
                .schedule-item {
                    border: 1px solid #404040;
                    border-radius: 6px;
                    margin-bottom: 15px;
                    padding: 15px;
                    background: #252526;
                }
                
                .schedule-header {
                    display: flex;
                    gap: 10px;
                    align-items: center;
                    margin-bottom: 15px;
                }
                
                .schedule-name {
                    flex: 1;
                    padding: 8px;
                    border: 1px solid #404040;
                    border-radius: 4px;
                    background: #1e1e1e;
                    color: #f0f0f0;
                    font-size: 14px;
                }
                
                .schedule-name:focus {
                    outline: none;
                    border-color: #0078d4;
                    box-shadow: 0 0 0 1px #0078d4;
                }
                
                .schedule-details {
                    display: grid;
                    grid-template-columns: 1fr 1fr;
                    gap: 15px;
                    align-items: center;
                }
                
                .schedule-details label {
                    color: #f0f0f0;
                    font-size: 14px;
                    font-weight: 500;
                }
                
                .days-selector {
                    grid-column: 1 / -1;
                    display: flex;
                    gap: 12px;
                    align-items: center;
                    flex-wrap: wrap;
                    margin-top: 10px;
                }
                
                .days-selector span {
                    color: #f0f0f0;
                    font-weight: 500;
                    margin-right: 5px;
                }
                
                .days-selector label {
                    display: flex;
                    align-items: center;
                    gap: 4px;
                    font-size: 13px;
                    color: #e0e0e0;
                    background: #1e1e1e;
                    padding: 4px 8px;
                    border-radius: 4px;
                    border: 1px solid #404040;
                    cursor: pointer;
                    transition: background-color 0.2s;
                }
                
                .days-selector label:hover {
                    background: #404040;
                }
                
                .days-selector input[type="checkbox"]:checked + span {
                    font-weight: bold;
                }
                
                .modal-footer {
                    padding: 20px;
                    border-top: 1px solid #404040;
                    display: flex;
                    gap: 10px;
                    justify-content: flex-end;
                    background: #1e1e1e;
                }
                
                button {
                    padding: 10px 16px;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 14px;
                    font-weight: 500;
                    transition: all 0.2s;
                }
                
                button:hover {
                    transform: translateY(-1px);
                    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
                }
                
                button:disabled {
                    opacity: 0.6;
                    cursor: not-allowed;
                    transform: none;
                }
                
                #save-config {
                    background: #0078d4;
                    color: white;
                }
                
                #save-config:hover:not(:disabled) {
                    background: #106ebe;
                }
                
                #cancel-config, .delete-schedule {
                    background: #424242;
                    color: #f0f0f0;
                    border: 1px solid #606060;
                }
                
                #cancel-config:hover, .delete-schedule:hover {
                    background: #525252;
                }
                
                #add-schedule {
                    background: #107c10;
                    color: white;
                }
                
                #add-schedule:hover:not(:disabled) {
                    background: #0e6e0e;
                }
                
                input[type="text"], input[type="time"], select {
                    padding: 8px;
                    border: 1px solid #404040;
                    border-radius: 4px;
                    background: #1e1e1e;
                    color: #f0f0f0;
                    font-size: 14px;
                }
                
                input[type="text"]:focus, input[type="time"]:focus, select:focus {
                    outline: none;
                    border-color: #0078d4;
                    box-shadow: 0 0 0 1px #0078d4;
                }
                
                select option {
                    background: #1e1e1e;
                    color: #f0f0f0;
                }
                
                /* Custom scrollbar for dark theme */
                .modal-content::-webkit-scrollbar {
                    width: 8px;
                }
                
                .modal-content::-webkit-scrollbar-track {
                    background: #1e1e1e;
                }
                
                .modal-content::-webkit-scrollbar-thumb {
                    background: #404040;
                    border-radius: 4px;
                }
                
                .modal-content::-webkit-scrollbar-thumb:hover {
                    background: #505050;
                }
            `);

            debugLog('Appending modal to body');
            document.body.appendChild(modal);

            // Event handlers
            debugLog('Setting up event handlers');
            closeBtn.onclick = () => {
                debugLog('Close button clicked');
                modal.remove();
            };
            cancelBtn.onclick = () => {
                debugLog('Cancel button clicked');
                modal.remove();
            };
            backdrop.onclick = (e) => {
                if (e.target === backdrop) {
                    debugLog('Backdrop clicked');
                    modal.remove();
                }
            };

            saveBtn.onclick = () => {
                debugLog('Save button clicked');
                try {
                    const newConfig = { ...config };
                    newConfig.enabled = enableCheckbox.checked;
                    newConfig.debugMode = debugCheckbox.checked;
                    newConfig.hideAlerts = alertCheckbox.checked;
                    newConfig.schedules = [];

                    schedulesContainer.querySelectorAll('.schedule-item').forEach(item => {
                        const index = item.dataset.index;
                        const days = Array.from(item.querySelectorAll('.schedule-day:checked')).map(cb => parseInt(cb.value));
                        
                        newConfig.schedules.push({
                            id: config.schedules[index]?.id || `schedule_${Date.now()}`,
                            name: item.querySelector('.schedule-name').value,
                            time: item.querySelector('.schedule-time').value,
                            days: days,
                            status: item.querySelector('.schedule-status').value,
                            enabled: item.querySelector('.schedule-enabled').checked
                        });
                    });

                    saveConfig(newConfig);
                    
                    // Update alert hiding based on new config
                    if (newConfig.hideAlerts) {
                        startAlertHiding();
                    } else {
                        stopAlertHiding();
                    }
                    
                    modal.remove();
                    log('Configuration updated');
                } catch (error) {
                    errorLog('Error saving config:', error);
                }
            };

            // Add schedule button
            addButton.onclick = () => {
                debugLog('Add schedule button clicked');
                try {
                    const newIndex = schedulesContainer.children.length;
                    const newScheduleData = {
                        id: `schedule_${Date.now()}`,
                        name: 'New Schedule',
                        time: '09:00',
                        days: [1, 2, 3, 4, 5], // Weekdays by default
                        status: STATUS_TYPES.AVAILABLE,
                        enabled: true
                    };
                    
                    const newScheduleItem = createScheduleItem(newScheduleData, newIndex);
                    schedulesContainer.appendChild(newScheduleItem);
                    debugLog('New schedule item added');
                } catch (error) {
                    errorLog('Error adding schedule:', error);
                }
            };
            
            debugLog('Config UI created successfully');
            
        } catch (error) {
            errorLog('Error creating config UI:', error);
        }
    }

    // ===== CONFIG UI HELPERS =====
    function createScheduleItem(schedule, index) {
        debugLog('Creating schedule item:', index, schedule);
        
        try {
            const scheduleItem = document.createElement('div');
            scheduleItem.className = 'schedule-item';
            scheduleItem.dataset.index = index;
            
            // Schedule header
            const header = document.createElement('div');
            header.className = 'schedule-header';
            
            const nameInput = document.createElement('input');
            nameInput.type = 'text';
            nameInput.value = schedule.name;
            nameInput.placeholder = 'Schedule Name';
            nameInput.className = 'schedule-name';
            
            const enabledLabel = document.createElement('label');
            const enabledCheckbox = document.createElement('input');
            enabledCheckbox.type = 'checkbox';
            enabledCheckbox.className = 'schedule-enabled';
            enabledCheckbox.checked = schedule.enabled;
            
            enabledLabel.appendChild(enabledCheckbox);
            enabledLabel.appendChild(document.createTextNode(' Enabled'));
            
            const deleteBtn = document.createElement('button');
            deleteBtn.className = 'delete-schedule';
            deleteBtn.textContent = 'Delete';
            deleteBtn.onclick = () => {
                debugLog('Delete button clicked for schedule:', index);
                scheduleItem.remove();
            };
            
            header.appendChild(nameInput);
            header.appendChild(enabledLabel);
            header.appendChild(deleteBtn);
            
            // Schedule details
            const details = document.createElement('div');
            details.className = 'schedule-details';
            
            // Time input
            const timeLabel = document.createElement('label');
            timeLabel.textContent = 'Time: ';
            const timeInput = document.createElement('input');
            timeInput.type = 'time';
            timeInput.value = schedule.time;
            timeInput.className = 'schedule-time';
            timeLabel.appendChild(timeInput);
            
            // Status select
            const statusLabel = document.createElement('label');
            statusLabel.textContent = 'Status: ';
            const statusSelect = document.createElement('select');
            statusSelect.className = 'schedule-status';
            
            Object.values(STATUS_TYPES).forEach(status => {
                const option = document.createElement('option');
                option.value = status;
                option.textContent = status;
                option.selected = schedule.status === status;
                statusSelect.appendChild(option);
            });
            statusLabel.appendChild(statusSelect);
            
            // Days selector
            const daysDiv = document.createElement('div');
            daysDiv.className = 'days-selector';
            
            const daysSpan = document.createElement('span');
            daysSpan.textContent = 'Days:';
            daysDiv.appendChild(daysSpan);
            
            ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].forEach((day, dayIndex) => {
                const dayLabel = document.createElement('label');
                const dayCheckbox = document.createElement('input');
                dayCheckbox.type = 'checkbox';
                dayCheckbox.value = dayIndex;
                dayCheckbox.className = 'schedule-day';
                dayCheckbox.checked = schedule.days.includes(dayIndex);
                
                dayLabel.appendChild(dayCheckbox);
                dayLabel.appendChild(document.createTextNode(' ' + day));
                daysDiv.appendChild(dayLabel);
            });
            
            details.appendChild(timeLabel);
            details.appendChild(statusLabel);
            details.appendChild(daysDiv);
            
            scheduleItem.appendChild(header);
            scheduleItem.appendChild(details);
            
            debugLog('Schedule item created successfully:', index);
            return scheduleItem;
            
        } catch (error) {
            errorLog('Error creating schedule item:', error);
            throw error;
        }
    }

    // ===== ALERT HIDING =====
    function hideTeamsAlerts() {
        try {
            const alerts = document.querySelectorAll(SELECTORS.UI_ALERTS);
            if (alerts.length > 0) {
                debugLog(`Found ${alerts.length} alert(s), hiding them`);
                alerts.forEach(alert => {
                    alert.style.display = 'none';
                });
            }
        } catch (error) {
            errorLog('Error hiding alerts:', error);
        }
    }

    function startAlertHiding() {
        if (alertHidingInterval) {
            clearInterval(alertHidingInterval);
        }
        log('Starting alert hiding interval');
        alertHidingInterval = setInterval(() => {
            try {
                hideTeamsAlerts();
            } catch (error) {
                errorLog('Error in alert hiding interval:', error);
            }
        }, 1000); // Check every second as per the example
    }

    function stopAlertHiding() {
        if (alertHidingInterval) {
            log('Stopping alert hiding interval');
            clearInterval(alertHidingInterval);
            alertHidingInterval = null;
        }
    }

    // ===== MAIN EXECUTION =====
    function init() {
        log('MS Teams Auto Status script initialized');
        
        // Setup debugging and monitoring
        setupTrustedTypesMonitoring();
        
        debugLog('Document readyState:', document.readyState);
        debugLog('Current URL:', window.location.href);
        debugLog('TrustedTypes available:', !!window.trustedTypes);

        try {
            // Add configuration button to page
            const configButton = document.createElement('button');
            configButton.textContent = '⚙️ Auto Status'; // Use textContent instead of innerHTML
            configButton.style.cssText = `
                position: fixed;
                top: 10px;
                left: 10px;
                z-index: 9999;
                background: #0078d4;
                color: white;
                border: none;
                padding: 8px 12px;
                border-radius: 4px;
                cursor: pointer;
                font-size: 12px;
                box-shadow: 0 2px 10px rgba(0, 0, 0, 0.2);
            `;
            configButton.onclick = () => {
                debugLog('Config button clicked');
                try {
                    createConfigUI();
                } catch (error) {
                    errorLog('Error creating config UI:', error);
                }
            };
            
            debugLog('Appending config button to body');
            document.body.appendChild(configButton);
            debugLog('Config button added successfully');

            // Start checking for status changes
            const config = getConfig();
            debugLog('Loaded config:', config);
            
            setInterval(() => {
                try {
                    checkAndUpdateStatus();
                } catch (error) {
                    errorLog('Error in status check interval:', error);
                }
            }, config.checkInterval);
            
            // Start alert hiding if enabled
            if (config.hideAlerts) {
                startAlertHiding();
            }
            
            // Run initial check after a delay to let Teams load
            setTimeout(() => {
                debugLog('Running initial status check');
                try {
                    checkAndUpdateStatus();
                } catch (error) {
                    errorLog('Error in initial status check:', error);
                }
            }, 5000);
            
            log('Auto status monitoring started');
            
        } catch (error) {
            errorLog('Error during initialization:', error);
        }
    }

    // Wait for Teams to load before initializing
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

})();
