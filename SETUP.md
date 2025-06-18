# GCG Automation Setup Guide

Complete installation and configuration instructions for the GCG Automation System.

## 📋 Prerequisites

### Required Access
- [ ] Google Workspace account with Apps Script permissions
- [ ] Access to Breeze Church Management System
- [ ] Google Drive folder for automation files
- [ ] Google Sheets with GCG Placement data structure

### Required Knowledge
- Basic Google Sheets navigation
- Understanding of your church's GCG structure
- Familiarity with Breeze export process
- Basic understanding of Person IDs in Breeze

### Required Data Structure
Your Google Sheet must have these tabs with proper formatting:
- [ ] **GCG Members tab**: Person ID in column A, proper headers
- [ ] **Not in a GCG tab**: Family-grouped format, headers in row 3
- [ ] **Group List tab**: Alphabetical group leader listing

## 🚀 Installation Steps

### Step 1: Prepare Google Drive Folder

1. **Create or identify** a Google Drive folder for automation files
2. **Note the folder ID** from the URL: 
   ```
   https://drive.google.com/drive/folders/[FOLDER_ID]
   ```
3. **Ensure permissions** allow the automation script to access files
4. **Test access** by uploading a sample file to the folder

### Step 2: Prepare Google Sheet Structure

1. **Open your GCG Placement sheet** or create a copy for testing
2. **Note the sheet ID** from the URL: 
   ```
   https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit
   ```
3. **Verify required tabs exist**:

#### GCG Members Tab
- **Person ID in column A** (critical for matching)
- **Headers**: First, Last, Group, Deacon, Pastor, Team, Active in GCG, Serving, Action Steps/Comments, Assigned to
- **Action Steps column**: Used for inactive member detection (keywords: "inactive", "moved away", "left church")

#### Not in a GCG Tab  
- **Family-grouped format** with one representative per family
- **Headers in row 3, data starts row 4**
- **Family ID and Family Role columns** for proper family grouping

#### Group List Tab
- **Column A**: Group Leader names (sorted A-Z)
- **Additional columns**: Co-Leader, Deacon, Pastor assignments

### Step 3: Create Container-Bound Apps Script

1. **In your Google Sheet**: Go to `Extensions → Apps Script`
2. **This creates a bound project** attached to your sheet
3. **Delete the default `Code.gs`** file (we'll replace with our files)
4. **Note the Apps Script project ID** for reference

### Step 4: Add the Code Files

Copy each file from the `/src/` folder into your Apps Script project:

#### Required Files (copy in this order):
1. **Create `config.js`**: Configuration management and script properties
2. **Create `smart-file-detection.js`**: Automatic file finding with date patterns  
3. **Create `google-sheets-parser.js`**: Data parsing from Breeze exports
4. **Create `comparison-engine.js`**: Core comparison logic with normalization
5. **Create `preview-report.js`**: Comprehensive preview report generation
6. **Create `menu-system.js`**: User interface and safety features

#### Copy Process:
- Create each file using the `+` button in Apps Script
- Copy the entire content from the repository file
- Save each file before proceeding to the next
- Ensure all 6 files are present and saved

### Step 5: Configure Script Properties

1. **In Apps Script**: Go to `Project Settings` (⚙️ gear icon)
2. **Scroll to "Script Properties"**
3. **Add these properties**:

```json
Property: GCG_FILE_CONFIG
Value: {
  "DRIVE_FOLDER_ID": "your_folder_id_here",
  "SHEET_ID": "your_sheet_id_here",
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
  "NOTIFICATIONS": {
    "ADMIN_EMAIL": "your_email_here",
    "SEND_CHANGE_NOTIFICATIONS": true,
    "SEND_ERROR_NOTIFICATIONS": true
  }
}
```

4. **Replace placeholders** with your actual values:
   - `your_folder_id_here`: Your Google Drive folder ID
   - `your_sheet_id_here`: Your Google Sheet ID  
   - `your_email_here`: Your email for notifications

### Step 6: Set Up Permissions

1. **Run initial setup**: In Apps Script, run the `setupConfig()` function
2. **Grant permissions** when prompted:
   - Google Drive access (to read export files)
   - Google Sheets access (to update your sheet)
   - Google Apps Script access (for the menu system)
3. **Test permissions** by running `testConfig()`

## 🧪 Testing Your Setup

### Initial Health Check
1. **In your Google Sheet**: Look for the "Breeze Update" menu (may need to refresh)
2. **Run**: Breeze Update → System Health Check
3. **Expected result**: All green checkmarks ✅
4. **If issues**: See troubleshooting section below

### Test File Detection
1. **Export sample files** from Breeze:
   - Export "Members - Active" to Google Sheets
   - Export "Tags" to Google Sheets
   - Name them with date format: `immanuelky-people-MM-DD-YYYY`
2. **Upload to your Google Drive folder**
3. **Run**: Breeze Update → Generate Preview Report
4. **Review results** for accuracy and completeness

### Verify Menu Functions
Test each menu item:
- [ ] Tutorial displays correctly with preparation steps
- [ ] Health Check shows green status for all components
- [ ] Configuration shows your folder and sheet IDs
- [ ] Preview generation works (even if no changes detected)

## 🔧 Breeze Export Setup

### Required Export Settings

#### Active Members Export
- **Export to**: Google Sheets (not Excel)
- **Include these fields**: 
  - Person ID, First Name, Last Name
  - Family ID, Family Role (for family grouping)
  - Address fields (Street, City, State, ZIP)
  - Any other fields you track
- **Naming convention**: `immanuelky-people-MM-DD-YYYY`

#### Tags Export  
- **Export to**: Google Sheets (not Excel)
- **Include all GCG tags** (the system will filter automatically)
- **Naming convention**: `immanuelky-tags-MM-DD-YYYY`

### Monthly Export Process
1. **Export both files** from Breeze on the same day
2. **Upload to designated Google Drive folder**
3. **System automatically detects** latest files by modification date
4. **No code changes needed** for different date formats

## 🚨 Troubleshooting

### Menu Not Appearing
- **Check**: Apps Script is container-bound to your sheet
- **Try**: Refresh the Google Sheet page
- **Manually run**: `onOpen()` function from Apps Script
- **Verify**: All 6 code files are present and saved

### Health Check Failures

#### Drive Access Failed
- **Check**: Folder ID is correct in script properties
- **Verify**: You have edit access to the Google Drive folder
- **Try**: Open the folder URL directly in a browser
- **Solution**: Update `DRIVE_FOLDER_ID` in script properties

#### Sheet Access Failed  
- **Check**: Sheet ID is correct in script properties
- **Verify**: Sheet contains all required tabs
- **Try**: Open the sheet URL directly in a browser
- **Solution**: Update `SHEET_ID` in script properties

#### Export Files Not Found
- **Check**: Files are uploaded to correct Google Drive folder
- **Verify**: File names contain required patterns (`immanuelky-people`, `immanuelky-tags`)
- **Try**: Upload files with standard naming convention
- **Solution**: Files must be in Google Sheets format, not Excel

#### Data Parsing Failed
- **Check**: Export files are in Google Sheets format (not Excel)
- **Verify**: Files contain expected data structure and headers
- **Try**: Open files directly to verify content and format
- **Solution**: Re-export from Breeze with correct settings

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

#### JSON Formatting Errors in Script Properties
- **Use a JSON validator** before pasting configuration
- **Check for trailing commas** (not allowed in JSON)
- **Ensure proper quote marks** (double quotes only)
- **Verify bracket matching** and proper nesting

#### Permission Issues
- **Re-run authorization**: Delete and re-create the script properties
- **Check account permissions**: Ensure you have edit access to all resources
- **Clear browser cache**: Sometimes helps with permission issues

## 🔄 Updating the System

### Code Updates
1. **Pull latest changes** from the repository
2. **Replace code files** in your Apps Script project one by one
3. **Save each file** after replacement
4. **Test functionality** with Health Check
5. **Verify menu operations** work correctly

### Configuration Changes
1. **Modify Script Properties** in Apps Script project settings
2. **Test changes** with `testConfig()` function
3. **Run Health Check** to verify new settings
4. **Update documentation** if you change file patterns

### Data Structure Updates
1. **Backup your current sheet** before making changes
2. **Update column mappings** in the code if needed
3. **Test with preview report** before applying changes
4. **Verify family logic** works with your data structure

## 📚 Best Practices

### Monthly Maintenance
- **Run Health Check** before each monthly update
- **Generate Preview Report** and review carefully before applying changes
- **Mark inactive members** in Action Steps column before running updates
- **Keep backup copies** of your GCG Placement sheet

### Data Quality
- **Ensure Person IDs** are populated in column A of GCG Members tab
- **Use consistent keywords** for inactive members ("inactive", "moved away", "left church")
- **Verify family data** includes Family ID and Family Role from Breeze
- **Review export completeness** before uploading to Drive

### Safety
- **Never apply updates** without reviewing the preview report first
- **Use test sheet** for initial setup and major changes
- **Contact administrator** if you see unexpected results
- **Keep original Breeze data** as source of truth

## 🎯 Success Indicators

Your setup is working correctly when:
- [ ] Health Check shows all green checkmarks
- [ ] Preview report generates without errors
- [ ] Family grouping shows one representative per family
- [ ] Group name normalization handles co-leader formats
- [ ] Inactive members are properly filtered
- [ ] Updates preserve pastoral notes and care data

## 🤝 Getting Help

If you encounter issues:
1. **Run Health Check** for diagnostic information
2. **Review troubleshooting section** above
3. **Check the tutorial** in the Breeze Update menu
4. **Contact system administrator**: sstringer@immanuelky.org

Remember: The system is designed to be safe and thorough. When in doubt, generate a preview report and review it carefully before applying any changes.
