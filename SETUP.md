# GCG Automation Setup Guide

Complete installation and configuration instructions for the GCG Automation System.

## 📋 Prerequisites

### Required Access
- [ ] Google Workspace account with Apps Script permissions
- [ ] Access to Breeze Church Management System
- [ ] Google Drive folder for automation files
- [ ] Google Sheets with GCG Placement data

### Required Knowledge
- Basic Google Sheets navigation
- Understanding of your church's GCG structure
- Familiarity with Breeze export process

## 🚀 Installation Steps

### Step 1: Prepare Google Drive Folder

1. **Create or identify** a Google Drive folder for automation files
2. **Note the folder ID** from the URL: `https://drive.google.com/drive/folders/[FOLDER_ID]`
3. **Ensure permissions** allow the automation script to access files

### Step 2: Prepare Google Sheet

1. **Open your GCG Placement sheet** or create a copy for testing
2. **Note the sheet ID** from the URL: `https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit`
3. **Ensure the sheet has these tabs**:
   - `GCG Members` (with Person ID in column A)
   - `Not in a GCG`
   - `Group List`

### Step 3: Create Container-Bound Apps Script

1. **In your Google Sheet**: Go to `Extensions → Apps Script`
2. **This creates a new bound project** attached to your sheet
3. **Delete the default `Code.gs`** file

### Step 4: Add the Code Files

Copy each file from the `/src/` folder into your Apps Script project:

1. **Create `config.js`**: Copy content from repository
2. **Create `smart-file-detection.js`**: Copy content from repository  
3. **Create `google-sheets-parser.js`**: Copy content from repository
4. **Create `comparison-engine.js`**: Copy content from repository
5. **Create `preview-report.js`**: Copy content from repository
6. **Create `menu-system.js`**: Copy content from repository

### Step 5: Configure Script Properties

1. **In Apps Script**: Go to `Project Settings` (⚙️ gear icon)
2. **Scroll to "Script Properties"**
3. **Add script property**:
   - **Property**: `GCG_FILE_CONFIG`
   - **Value**: See configuration template below

#### Configuration Template
```json
{
  "DRIVE_FOLDER_ID": "YOUR_GOOGLE_DRIVE_FOLDER_ID",
  "SHEET_ID": "YOUR_GOOGLE_SHEET_ID",
  "FILE_PATTERNS": {
    "ACTIVE_MEMBERS": {
      "contains": "immanuelky-people",
      "excludes": [],
      "description": "Active Members Export"
    },
    "TAGS_EXPORT": {
      "contains": "immanuelky-tags",
      "excludes": [],
      "description": "Tags Export"
    }
  },
  "FILE_SELECTION": {
    "strategy": "latest",
    "dateFormats": ["MM-dd-yyyy", "MM-dd-yy", "yyyy-MM-dd", "dd-MM-yyyy"]
  },
  "NOTIFICATIONS": {
    "ADMIN_EMAIL": "your-email@church.org",
    "SEND_CHANGE_NOTIFICATIONS": true,
    "SEND_ERROR_NOTIFICATIONS": true
  }
}
```

**Replace these values**:
- `YOUR_GOOGLE_DRIVE_FOLDER_ID`: The folder ID from Step 1
- `YOUR_GOOGLE_SHEET_ID`: The sheet ID from Step 2  
- `your-email@church.org`: Your admin email address

### Step 6: Test the Installation

1. **Save the project** (Ctrl+S)
2. **Run the test function**: Execute `testConfig()` from `config.js`
3. **Expected output**:
   ```
   🧪 Testing configuration...
   ✅ Drive access: [Your folder name]
   ✅ Sheets access: [Your sheet name]
   ```

### Step 7: Setup the Menu

1. **In your Google Sheet**: Refresh the page
2. **Look for "Breeze Update" menu** in the menu bar
3. **If no menu appears**: Run `onOpen()` manually from Apps Script
4. **Test menu items**: Try the Tutorial and Health Check

## 🔧 Configuration Details

### File Naming Conventions

The system automatically detects files with these patterns:

- **Active Members**: Must contain `immanuelky-people`
  - Example: `immanuelky-people-redacted-06-13-2025`
- **Tags Export**: Must contain `immanuelky-tags`
  - Example: `immanuelky-tags-06-13-2025`

### Sheet Structure Requirements

#### GCG Members Tab
- **Column A**: Person ID (hidden after setup)
- **Column B**: First Name
- **Column C**: Last Name  
- **Column D**: Group
- **Columns E+**: Deacon, Pastor, Team, etc.
- **Action Steps column**: Used for inactive member detection

#### Not in a GCG Tab  
- **Standard family-grouped format**
- **Headers in row 3, data starts row 4**

#### Group List Tab
- **Column A**: Group Leader (sorted A-Z)
- **Columns B+**: Co-Leader, Deacon, Pastor, etc.

## 🧪 Testing Your Setup

### Run Health Check
1. **Breeze Update → System Health Check**
2. **Expected result**: All green checkmarks
3. **If issues**: See troubleshooting section below

### Test File Detection
1. **Upload test export files** to your Google Drive folder
2. **Generate Preview Report**
3. **Review results** for accuracy

### Verify Menu Functions
- [ ] Tutorial displays correctly
- [ ] Health Check shows system status
- [ ] Configuration shows your settings
- [ ] Preview generation works (even if no changes)

## 🚨 Troubleshooting

### Menu Not Appearing
- **Check**: Apps Script is container-bound to your sheet
- **Try**: Run `onOpen()` manually from Apps Script
- **Verify**: All 6 code files are present and saved

### Health Check Failures

#### Drive Access Failed
- **Check**: Folder ID is correct in configuration
- **Verify**: You have access to the Google Drive folder
- **Try**: Open the folder directly in Google Drive

#### Sheet Access Failed  
- **Check**: Sheet ID is correct in configuration
- **Verify**: Sheet contains required tabs
- **Try**: Open the sheet directly

#### Export Files Not Found
- **Check**: Files are uploaded to correct Google Drive folder
- **Verify**: File names contain required patterns
- **Try**: Upload files with correct naming convention

#### Data Parsing Failed
- **Check**: Export files are in Google Sheets format (not Excel)
- **Verify**: Files contain expected data structure
- **Try**: Open files directly to check content

### Common Configuration Issues

#### Wrong Folder/Sheet IDs
```bash
# Get Folder ID from URL
https://drive.google.com/drive/folders/1P7JKNiYgFcQh6TtzHS1q9DofASu6pfYh
                                        ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                                        This is your FOLDER_ID

# Get Sheet ID from URL  
https://docs.google.com/spreadsheets/d/1H_bKbWbSTCBJWffd4bbGiRpbxjcxIhq2hb7ICUt1-M0/edit
                                       ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                                       This is your SHEET_ID
```

#### JSON Formatting Errors
- **Use a JSON validator** before pasting configuration
- **Check for trailing commas** (not allowed in JSON)
- **Ensure proper quote marks** (double quotes only)

## 🔄 Updating the System

### Code Updates
1. **Pull latest changes** from the repository
2. **Replace code files** in your Apps Script project
3. **Test functionality** with Health Check
4. **Update documentation** if needed

### Configuration Changes
1. **Modify Script Properties** in Apps Script project settings
2. **Test changes** with `testConfig()` function
3. **Verify** with Health Check

## 📞 Getting Help

### Before Contacting Support
- [ ] Run Health Check and note specific error messages
- [ ] Verify file names and locations
- [ ] Check Script Properties configuration  
- [ ] Try accessing files/sheets manually

### What to Include in Support Requests
- Health Check output (copy/paste full results)
- Screenshot of error messages
- Description of what you were trying to do
- Recent changes to files or configuration

### Contact Information
- **System Administrator**: sstringer@immanuelky.org
- **Include**: "GCG Automation" in email subject

---

## 🎉 Success Checklist

Once setup is complete, you should be able to:

- [ ] See "Breeze Update" menu in your Google Sheet
- [ ] Run Health Check with all green results
- [ ] Generate preview reports (even with no changes)
- [ ] Upload new export files and detect them automatically
- [ ] View tutorial and configuration through the menu

**Congratulations! Your GCG Automation system is ready for use.** 🚀
