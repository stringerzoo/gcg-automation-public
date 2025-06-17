/**
 * Configuration for GCG File-Based Automation (Production Version)
 */

/**
 * Default configuration
 */
const DEFAULT_CONFIG = {
  // Google Drive folder where exports are stored
  DRIVE_FOLDER_ID: '1P7JKNiYgFcQh6TtzHS1q9DofASu6pfYh',
  
  // Google Sheets ID (your working sheet)
  SHEET_ID: '1H_bKbWbSTCBJWffd4bbGiRpbxjcxIhq2hb7ICUt1-M0',
  
  // Smart file patterns that work with any date format
  FILE_PATTERNS: {
    ACTIVE_MEMBERS: {
      contains: 'immanuelky-people',
      excludes: [],
      description: 'Active Members Export'
    },
    TAGS_EXPORT: {
      contains: 'immanuelky-tags',
      excludes: [],
      description: 'Tags Export'
    },
    CURRENT_SHEET: {
      contains: 'GCG Placement',
      excludes: [],
      description: 'Current Working Sheet'
    }
  },
  
  // File selection strategy
  FILE_SELECTION: {
    strategy: 'latest',
    dateFormats: [
      'MM-dd-yyyy',
      'MM-dd-yy', 
      'yyyy-MM-dd',
      'dd-MM-yyyy'
    ]
  },
  
  // Email notifications
  NOTIFICATIONS: {
    ADMIN_EMAIL: 'sstringer@immanuelky.org',
    SEND_CHANGE_NOTIFICATIONS: true,
    SEND_ERROR_NOTIFICATIONS: true
  },
  
  // Expected column structures
  COLUMNS: {
    ACTIVE_MEMBERS: {
      PERSON_ID: 'Breeze ID',
      FIRST_NAME: 'First Name', 
      LAST_NAME: 'Last Name',
      STREET_ADDRESS: 'Street Address',
      CITY: 'City',
      STATE: 'State',
      ZIP: 'Zip'
    },
    
    GCG_TAGS: {
      PERSON_ID: 'Person ID',
      FIRST_NAME: 'First Name',
      LAST_NAME: 'Last Name'
    }
  }
};

/**
 * Get configuration from Properties Service or defaults
 */
function getConfig() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const configData = properties.getProperty('GCG_FILE_CONFIG');
    
    if (configData) {
      const savedConfig = JSON.parse(configData);
      return { ...DEFAULT_CONFIG, ...savedConfig };
    }
  } catch (error) {
    console.warn('No saved config found, using defaults:', error.message);
  }
  
  return DEFAULT_CONFIG;
}

/**
 * Save configuration to Properties Service
 */
function saveConfig(config) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('GCG_FILE_CONFIG', JSON.stringify(config));
    console.log('✅ Configuration saved to Properties Service');
  } catch (error) {
    console.error('❌ Failed to save configuration:', error);
    throw error;
  }
}

/**
 * Setup configuration with smart file detection
 */
function setupConfig() {
  const config = {
    ...DEFAULT_CONFIG,
    DRIVE_FOLDER_ID: '1P7JKNiYgFcQh6TtzHS1q9DofASu6pfYh',
    SHEET_ID: '1H_bKbWbSTCBJWffd4bbGiRpbxjcxIhq2hb7ICUt1-M0'
  };
  
  saveConfig(config);
  console.log('✅ Configuration setup completed');
  return config;
}

/**
 * Test configuration values
 */
function testConfig() {
  console.log('🧪 Testing configuration...');
  
  const config = getConfig();
  
  console.log('Configuration values:');
  console.log(`- Drive Folder ID: ${config.DRIVE_FOLDER_ID}`);
  console.log(`- Sheet ID: ${config.SHEET_ID}`);
  console.log(`- Admin Email: ${config.NOTIFICATIONS.ADMIN_EMAIL}`);
  console.log(`- File Patterns: ${JSON.stringify(config.FILE_PATTERNS)}`);
  
  // Test Drive access
  try {
    const folder = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);
    console.log(`✅ Drive access: ${folder.getName()}`);
  } catch (error) {
    console.log(`❌ Drive access failed: ${error.message}`);
  }
  
  // Test Sheets access
  try {
    const ss = SpreadsheetApp.openById(config.SHEET_ID);
    console.log(`✅ Sheets access: ${ss.getName()}`);
  } catch (error) {
    console.log(`❌ Sheets access failed: ${error.message}`);
  }
  
  return config;
}
