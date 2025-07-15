/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

/**
 * Apply Fiery Red theme to the email
 * @param event {Office.AddinCommands.Event}
 */
function applyFieryRed(event) {
  applyColorTheme('fiery-red', event);
}

/**
 * Apply Cool Blue theme to the email
 * @param event {Office.AddinCommands.Event}
 */
function applyCoolBlue(event) {
  applyColorTheme('cool-blue', event);
}

/**
 * Apply Earth Green theme to the email
 * @param event {Office.AddinCommands.Event}
 */
function applyEarthGreen(event) {
  applyColorTheme('earth-green', event);
}

/**
 * Apply Sunshine Yellow theme to the email
 * @param event {Office.AddinCommands.Event}
 */
function applySunshineYellow(event) {
  applyColorTheme('sunshine-yellow', event);
}

/**
 * Apply the selected color theme to the email
 * @param {string} colorTheme - The selected color theme
 * @param {Office.AddinCommands.Event} event - The event object
 */
function applyColorTheme(colorTheme, event) {
  let themeMessage = '';
  let backgroundColor = '';
  let textColor = '';

  switch(colorTheme) {
    case 'fiery-red':
      themeMessage = 'Applied Fiery Red theme to your message!';
      backgroundColor = '#FFE6E6';
      textColor = '#CC0000';
      break;
    case 'cool-blue':
      themeMessage = 'Applied Cool Blue theme to your message!';
      backgroundColor = '#E6F3FF';
      textColor = '#0066CC';
      break;
    case 'earth-green':
      themeMessage = 'Applied Earth Green theme to your message!';
      backgroundColor = '#E6F7E6';
      textColor = '#006600';
      break;
    case 'sunshine-yellow':
      themeMessage = 'Applied Sunshine Yellow theme to your message!';
      backgroundColor = '#FFF9E6';
      textColor = '#CC9900';
      break;
    default:
      themeMessage = 'Unknown color theme selected.';
      // Still complete the event even for unknown themes
      event.completed();
      return;
  }

  // Apply the theme to the email body
  Office.context.mailbox.item.body.setAsync(
    `<div style="background-color: ${backgroundColor}; color: ${textColor}; padding: 15px; border-radius: 5px; border: 2px solid ${textColor};">
      <p>This message is composed with ${colorTheme.replace('-', ' ')} theme!</p>
      <p>Add your content here...</p>
    </div>`,
    { coercionType: Office.CoercionType.Html },
    function(result) {
      // Show success notification regardless of result
      const message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: themeMessage,
        icon: "Icon.80x80",
        persistent: false,
      };

      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "ColorThemeNotification",
        message
      );

      // CRITICAL: Always call event.completed() to close the menu
      event.completed();
    }
  );
}

// Register the functions with Office.
Office.actions.associate("action", action);
Office.actions.associate("applyFieryRed", applyFieryRed);
Office.actions.associate("applyCoolBlue", applyCoolBlue);
Office.actions.associate("applyEarthGreen", applyEarthGreen);
Office.actions.associate("applySunshineYellow", applySunshineYellow);
