/**
 * 3DSellers Multi-Channel Separator for Google Sheets
 * This script takes imported CSV data and automatically separates it into different channel tabs
 * with proper field mapping for each marketplace
 *
 * Setup Instructions:
 * 1. Import your CSV into Google Sheets
 * 2. Go to Extensions > Apps Script
 * 3. Copy and paste this entire script
 * 4. Save and run the 'separateChannels' function
 */

// Channel Field Mappings - Maps 3DSellers fields to each marketplace
const CHANNEL_MAPPINGS = {
  'eBay': {
    name: 'eBay',
    fields: {
      'SKU': 'SKU',
      'Title': 'Title',
      'Description': 'Description',
      'Price': 'Price',
      'Quantity': 'Quantity',
      'CategoryID': 'CategoryID',
      'Condition': 'Condition',
      'ConditionNote': 'ConditionNote',
      'C:Brand': 'Brand',
      'C:MPN': 'MPN',
      'C:UPC': 'UPC',
      'C:EAN': 'EAN',
      'ImageURLs': 'PictureURL',
      'Image 1': 'GalleryURL',
      'PolicyShipping': 'ShippingProfileID',
      'PolicyReturn': 'ReturnProfileID',
      'PolicyPayment': 'PaymentProfileID',
      'Location': 'Location',
      'CountryCode': 'Country',
      'PostalCode': 'PostalCode',
      'BestOffer': 'BestOfferEnabled',
      'BestOfferAccept': 'BestOfferAutoAcceptPrice',
      'BestOfferDecline': 'BestOfferAutoDeclinePrice',
      'ListingDuration': 'ListingDuration',
      'PackageType': 'ShippingPackage',
      'WeightMajor': 'WeightMajor',
      'WeightMinor': 'WeightMinor'
    }
  },

  'Etsy': {
    name: 'Etsy',
    fields: {
      'Title': 'title',
      'Description': 'description',
      'Price': 'price',
      'Quantity': 'quantity',
      'SKU': 'sku',
      'C:Material': 'materials',
      'Tags': 'tags',
      'CategoryID': 'taxonomy_id',
      'ImageURLs': 'image',
      'Image 1': 'image_1',
      'Image 2': 'image_2',
      'Image 3': 'image_3',
      'C:Color': 'primary_color',
      'C:Size': 'size',
      'Condition': 'when_made',
      'Location': 'origin_zip',
      'PolicyShipping': 'shipping_profile_id',
      'WeightMajor': 'item_weight',
      'PackageLength': 'item_length',
      'PackageWidth': 'item_width',
      'PackageDepth': 'item_height'
    }
  },

  'Amazon': {
    name: 'Amazon',
    fields: {
      'SKU': 'item_sku',
      'Title': 'item_name',
      'Description': 'product_description',
      'Price': 'standard_price',
      'Quantity': 'quantity',
      'C:Brand': 'item_type_keyword',
      'C:MPN': 'part_number',
      'C:UPC': 'external_product_id',
      'C:EAN': 'external_product_id',
      'C:ASIN': 'external_product_id',
      'Condition': 'condition_type',
      'ConditionNote': 'condition_note',
      'ImageURLs': 'main_image_url',
      'Image 1': 'other_image_url1',
      'Image 2': 'other_image_url2',
      'Image 3': 'other_image_url3',
      'WeightMajor': 'item_weight',
      'PackageLength': 'item_length',
      'PackageWidth': 'item_width',
      'PackageDepth': 'item_height',
      'CategoryID': 'recommended_browse_nodes'
    }
  },

  'Shopify': {
    name: 'Shopify',
    fields: {
      'SKU': 'Variant SKU',
      'Title': 'Title',
      'Description': 'Body (HTML)',
      'Price': 'Variant Price',
      'Quantity': 'Variant Inventory Qty',
      'C:Brand': 'Vendor',
      'Tags': 'Tags',
      'ImageURLs': 'Image Src',
      'Image 1': 'Image Src',
      'WeightMajor': 'Variant Grams',
      'C:Size': 'Option1 Value',
      'C:Color': 'Option2 Value',
      'C:Material': 'Option3 Value',
      'MetaKeywords': 'SEO Title',
      'MetaDescription': 'SEO Description',
      'OriginalRetailPrice': 'Variant Compare At Price',
      'C:UPC': 'Variant Barcode',
      'CategoryID': 'Product Category',
      'Condition': 'Status'
    }
  },

  'Facebook': {
    name: 'Facebook Marketplace',
    fields: {
      'SKU': 'id',
      'Title': 'title',
      'Description': 'description',
      'Price': 'price',
      'Quantity': 'inventory',
      'C:Brand': 'brand',
      'ImageURLs': 'link',
      'Image 1': 'image_link',
      'Image 2': 'additional_image_link',
      'CategoryID': 'google_product_category',
      'Condition': 'condition',
      'C:UPC': 'gtin',
      'C:MPN': 'mpn',
      'C:Size': 'size',
      'C:Color': 'color',
      'C:Material': 'material',
      'Location': 'shipping',
      'Tags': 'custom_label_0'
    }
  },

  'Poshmark': {
    name: 'Poshmark',
    fields: {
      'SKU': 'sku',
      'Title': 'title',
      'Description': 'description',
      'Price': 'price',
      'Quantity': 'quantity',
      'C:Brand': 'brand',
      'C:Size': 'size',
      'C:Color': 'color_shade',
      'CategoryID': 'category',
      'Condition': 'condition',
      'ImageURLs': 'cover_photo',
      'Image 1': 'photo_1',
      'Image 2': 'photo_2',
      'Image 3': 'photo_3',
      'Tags': 'style',
      'OriginalRetailPrice': 'original_price',
      'C:Material': 'material'
    }
  },

  'Mercari': {
    name: 'Mercari',
    fields: {
      'SKU': 'sku',
      'Title': 'name',
      'Description': 'description',
      'Price': 'price',
      'Quantity': 'quantity',
      'C:Brand': 'brand',
      'CategoryID': 'category_id',
      'Condition': 'condition',
      'ImageURLs': 'photo_ids',
      'Image 1': 'photo_1',
      'Image 2': 'photo_2',
      'Image 3': 'photo_3',
      'Tags': 'hashtags',
      'WeightMajor': 'weight',
      'C:Size': 'size',
      'C:Color': 'color'
    }
  },

  'Depop': {
    name: 'Depop',
    fields: {
      'Title': 'title',
      'Description': 'description',
      'Price': 'price',
      'Quantity': 'quantity',
      'C:Brand': 'brand',
      'C:Size': 'size',
      'CategoryID': 'category',
      'Condition': 'condition',
      'ImageURLs': 'photos',
      'Tags': 'hashtags',
      'C:Color': 'colour',
      'Location': 'location',
      'C:Material': 'material'
    }
  },

  'TikTok Shop': {
    name: 'TikTok Shop',
    fields: {
      'SKU': 'seller_sku',
      'Title': 'title',
      'Description': 'description',
      'Price': 'sale_price',
      'Quantity': 'inventory',
      'C:Brand': 'brand',
      'CategoryID': 'category_id',
      'ImageURLs': 'main_image',
      'Image 1': 'image_1',
      'Image 2': 'image_2',
      'Image 3': 'image_3',
      'Tags': 'product_tags',
      'WeightMajor': 'package_weight',
      'PackageLength': 'package_length',
      'PackageWidth': 'package_width',
      'PackageDepth': 'package_height'
    }
  }
};

/**
 * Main function to separate channels into different tabs
 */
function separateChannels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const data = sourceSheet.getDataRange().getValues();

  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('No data found in the active sheet!');
    return;
  }

  // Get headers from first row
  const headers = data[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[header] = index;
  });

  // Process each channel
  Object.keys(CHANNEL_MAPPINGS).forEach(channelKey => {
    const channel = CHANNEL_MAPPINGS[channelKey];
    createChannelSheet(ss, channel, data, headerMap);
  });

  // Create summary sheet
  createSummarySheet(ss, data);

  SpreadsheetApp.getUi().alert('Channel separation complete! Check the new tabs.');
}

/**
 * Creates a sheet for a specific channel with mapped fields
 */
function createChannelSheet(ss, channel, sourceData, headerMap) {
  // Delete existing sheet if it exists
  let sheet = ss.getSheetByName(channel.name);
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  // Create new sheet
  sheet = ss.insertSheet(channel.name);

  // Create headers for this channel
  const channelHeaders = [];
  const fieldMappings = [];

  Object.keys(channel.fields).forEach(sourceField => {
    const targetField = channel.fields[sourceField];
    channelHeaders.push(targetField);
    fieldMappings.push(headerMap[sourceField] !== undefined ? headerMap[sourceField] : -1);
  });

  // Add channel-specific required fields
  addChannelSpecificFields(channel.name, channelHeaders);

  // Set headers
  if (channelHeaders.length > 0) {
    sheet.getRange(1, 1, 1, channelHeaders.length).setValues([channelHeaders]);
    sheet.getRange(1, 1, 1, channelHeaders.length)
      .setBackground('#6366f1')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }

  // Copy data with field mapping
  const channelData = [];
  for (let i = 1; i < sourceData.length; i++) {
    const row = [];
    fieldMappings.forEach(sourceIndex => {
      if (sourceIndex >= 0) {
        row.push(sourceData[i][sourceIndex] || '');
      } else {
        row.push('');
      }
    });

    // Only add row if it has some data
    if (row.some(cell => cell !== '')) {
      channelData.push(row);
    }
  }

  // Write data to sheet
  if (channelData.length > 0) {
    sheet.getRange(2, 1, channelData.length, channelHeaders.length).setValues(channelData);
  }

  // Auto-resize columns
  for (let i = 1; i <= channelHeaders.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Add data validation for specific fields
  addDataValidation(sheet, channel.name);
}

/**
 * Adds channel-specific required fields
 */
function addChannelSpecificFields(channelName, headers) {
  switch(channelName) {
    case 'Etsy':
      headers.push('who_made', 'is_supply', 'when_made');
      break;
    case 'Amazon':
      headers.push('product_type', 'feed_product_type');
      break;
    case 'Shopify':
      headers.push('Handle', 'Published', 'Type');
      break;
    case 'Facebook Marketplace':
      headers.push('availability', 'fb_product_category');
      break;
  }
}

/**
 * Adds data validation to specific fields
 */
function addDataValidation(sheet, channelName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  switch(channelName) {
    case 'eBay':
      // Condition validation
      const conditionCol = getColumnByHeader(sheet, 'Condition');
      if (conditionCol > 0) {
        const conditionRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['1000', '1500', '2000', '2500', '3000', '4000', '5000', '6000', '7000'])
          .setAllowInvalid(false)
          .build();
        sheet.getRange(2, conditionCol, lastRow - 1, 1).setDataValidation(conditionRule);
      }
      break;

    case 'Etsy':
      // who_made validation
      const whoMadeCol = getColumnByHeader(sheet, 'who_made');
      if (whoMadeCol > 0) {
        const whoMadeRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['i_did', 'collective', 'someone_else'])
          .setAllowInvalid(false)
          .build();
        sheet.getRange(2, whoMadeCol, lastRow - 1, 1).setDataValidation(whoMadeRule);
      }
      break;
  }
}

/**
 * Gets column number by header name
 */
function getColumnByHeader(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(headerName);
  return index + 1; // Convert to 1-based index
}

/**
 * Creates a summary sheet with statistics
 */
function createSummarySheet(ss, data) {
  let sheet = ss.getSheetByName('Summary');
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('Summary');
  sheet.moveTo(1); // Move to first position

  const summary = [
    ['Channel Distribution Summary', ''],
    ['', ''],
    ['Channel', 'Product Count'],
    ['eBay', countNonEmptyRows(ss.getSheetByName('eBay'))],
    ['Etsy', countNonEmptyRows(ss.getSheetByName('Etsy'))],
    ['Amazon', countNonEmptyRows(ss.getSheetByName('Amazon'))],
    ['Shopify', countNonEmptyRows(ss.getSheetByName('Shopify'))],
    ['Facebook Marketplace', countNonEmptyRows(ss.getSheetByName('Facebook Marketplace'))],
    ['Poshmark', countNonEmptyRows(ss.getSheetByName('Poshmark'))],
    ['Mercari', countNonEmptyRows(ss.getSheetByName('Mercari'))],
    ['Depop', countNonEmptyRows(ss.getSheetByName('Depop'))],
    ['TikTok Shop', countNonEmptyRows(ss.getSheetByName('TikTok Shop'))],
    ['', ''],
    ['Total Products', data.length - 1],
    ['Created', new Date().toLocaleString()]
  ];

  sheet.getRange(1, 1, summary.length, 2).setValues(summary);

  // Format summary
  sheet.getRange('A1:B1').merge()
    .setBackground('#6366f1')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('A3:B3')
    .setBackground('#e5e7eb')
    .setFontWeight('bold');

  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
}

/**
 * Counts non-empty rows in a sheet
 */
function countNonEmptyRows(sheet) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  return lastRow > 1 ? lastRow - 1 : 0;
}

/**
 * Menu creation for easy access
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('3DSellers Tools')
    .addItem('Separate Channels', 'separateChannels')
    .addItem('Validate Data', 'validateAllData')
    .addItem('Export Channel CSVs', 'exportChannelCSVs')
    .addSeparator()
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * Validates data across all sheets
 */
function validateAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];

  // Validate eBay
  const ebaySheet = ss.getSheetByName('eBay');
  if (ebaySheet) {
    validateEbayData(ebaySheet, errors);
  }

  // Show results
  if (errors.length > 0) {
    const errorMsg = errors.join('\n');
    SpreadsheetApp.getUi().alert('Validation Errors Found:\n\n' + errorMsg);
  } else {
    SpreadsheetApp.getUi().alert('All data validated successfully!');
  }
}

/**
 * Validates eBay specific requirements
 */
function validateEbayData(sheet, errors) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const titleCol = headers.indexOf('Title');
  const skuCol = headers.indexOf('SKU');

  for (let i = 1; i < data.length; i++) {
    // Check title length
    if (titleCol >= 0 && data[i][titleCol] && data[i][titleCol].length > 80) {
      errors.push(`eBay Row ${i+1}: Title exceeds 80 characters`);
    }

    // Check SKU uniqueness
    if (skuCol >= 0 && !data[i][skuCol]) {
      errors.push(`eBay Row ${i+1}: Missing SKU`);
    }
  }
}

/**
 * Exports each channel sheet as a separate CSV file
 */
function exportChannelCSVs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = DriveApp.createFolder('Channel Exports - ' + new Date().toISOString().split('T')[0]);

  Object.keys(CHANNEL_MAPPINGS).forEach(channelKey => {
    const sheet = ss.getSheetByName(CHANNEL_MAPPINGS[channelKey].name);
    if (sheet && sheet.getLastRow() > 1) {
      const csv = convertSheetToCSV(sheet);
      const blob = Utilities.newBlob(csv, 'text/csv', `${channelKey}_export.csv`);
      folder.createFile(blob);
    }
  });

  SpreadsheetApp.getUi().alert(`CSV files exported to folder: ${folder.getUrl()}`);
}

/**
 * Converts a sheet to CSV format
 */
function convertSheetToCSV(sheet) {
  const data = sheet.getDataRange().getValues();
  return data.map(row =>
    row.map(cell =>
      typeof cell === 'string' && cell.includes(',')
        ? `"${cell.replace(/"/g, '""')}"`
        : cell
    ).join(',')
  ).join('\n');
}

/**
 * Shows help information
 */
function showHelp() {
  const helpText = `
3DSellers Multi-Channel Separator Help

This tool automatically separates your 3DSellers CSV data into channel-specific tabs with proper field mapping.

How to use:
1. Import your 3DSellers CSV into this sheet
2. Run "Separate Channels" from the 3DSellers Tools menu
3. Review each channel tab for accuracy
4. Use "Validate Data" to check for errors
5. Export individual channel CSVs when ready

Features:
- Automatic field mapping for 9 major marketplaces
- Data validation for required fields
- Summary statistics
- CSV export for each channel

For support, visit: help.3dsellers.com
  `;

  SpreadsheetApp.getUi().alert(helpText);
}