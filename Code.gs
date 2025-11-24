/**
 * Runs when the add-on is installed.
 * @param {Object} e The event parameter for onInstall.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when the spreadsheet is opened.
 * @param {Object} e The event parameter for onOpen.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Apply Color Scale', 'showSidebar') // This will call showSidebar when clicked
      .addToUi();
}

/**
 * Displays a sidebar in the Google Sheet.
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Color Scale Settings')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}



"/**
 * Applies a color scale to rows based on a specified column's value.
 * @param {string} columnLetter The letter of the column to base the gradient on (e.g., "B").
 * @param {string} minColorHex The hex code for the minimum value color (e.g., "#FF0000").
 * @param {string} maxColorHex The hex code for the maximum value color (e.g., "#00FF00").
 * @returns {string} A status message.
 */
function applyColorScale(columnLetter, minColorHex, maxColorHex) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { // Need at least a header and one data row
      return 'No data found to apply color scale.';
    }

    // Convert hex colors to RGB objects for the API
    var minColor = hexToRgb(minColorHex);
    var maxColor = hexToRgb(maxColorHex);

    // Get the column index (A=1, B=2, etc.)
    var columnIndex = columnLetter.toUpperCase().charCodeAt(0) - 'A'.charCodeAt(0) + 1;

    // Get all values from the specified column to determine min/max for the gradient
    var columnRange = sheet.getRange(2, columnIndex, lastRow - 1, 1); // Start from row 2 (assuming header in row 1)
    var columnValues = columnRange.getValues().flat();

    // Filter out non-numeric values and find the actual min/max
    var numericValues = columnValues.filter(function(value) {
      return typeof value === 'number' && !isNaN(value);
    });

    if (numericValues.length === 0) {
      return 'No numeric data found in the specified column to base the gradient on.';
    }

    var minValue = Math.min.apply(null, numericValues);
    var maxValue = Math.max.apply(null, numericValues);

    if (minValue === maxValue) {
      return 'All values in the specified column are the same. Cannot apply a gradient.';
    }

    // Clear existing conditional formatting rules to avoid conflicts
    var rules = sheet.getConditionalFormatRules();
    var newRules = [];
    for (var i = 0; i < rules.length; i++) {
      // Keep rules that are not related to our add-on's formatting
      // For simplicity, we'll just clear all for now, but in a real add-on, you might want to be more selective.
    }
    sheet.setConditionalFormatRules(newRules);


    // Create a new conditional formatting rule for each row
    for (var r = 2; r <= lastRow; r++) { // Iterate through each data row (assuming header in row 1)
      var cellValue = sheet.getRange(r, columnIndex).getValue();

      if (typeof cellValue === 'number' && !isNaN(cellValue)) {
        // Calculate the interpolation factor for the current cell's value
        var factor = (cellValue - minValue) / (maxValue - minValue);

        // Interpolate RGB values
        var interpolatedR = Math.round(minColor.red + factor * (maxColor.red - minColor.red));
        var interpolatedG = Math.round(minColor.green + factor * (maxColor.green - minColor.green));
        var interpolatedB = Math.round(minColor.blue + factor * (maxColor.blue - minColor.blue));

        var interpolatedHex = rgbToHex(interpolatedR, interpolatedG, interpolatedB);

        // Apply the background color to the entire row
        sheet.getRange(r, 1, 1, sheet.getLastColumn()).setBackground(interpolatedHex);
      }
    }

    return 'Color scale applied successfully!';

  } catch (e) {
    return 'Error: ' + e.message;
  }
}

/**
 * Converts a hex color string to an RGB object.
 * @param {string} hex The hex color string (e.g., "#FF0000").
 * @returns {Object} An object with red, green, blue properties (0-255).
 */
function hexToRgb(hex) {
  var r = parseInt(hex.substring(1, 3), 16);
  var g = parseInt(hex.substring(3, 5), 16);
  var b = parseInt(hex.substring(5, 7), 16);
  return { red: r, green: g, blue: b };
}

/**
 * Converts RGB values to a hex color string.
 * @param {number} r Red value (0-255).
 * @param {number} g Green value (0-255).
 * @param {number} b Blue value (0-255).
 * @returns {string} The hex color string (e.g., "#FF0000").
 */
function rgbToHex(r, g, b) {
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}
"
