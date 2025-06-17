# GCG Automation Troubleshooting Guide

Common issues and solutions for the GCG Automation System.

## 🚨 Quick Diagnostic Steps

When something isn't working:

1. **Run Health Check**: `Breeze Update → System Health Check`
2. **Check Recent Changes**: Did you upload new files or change settings?
3. **Verify Access**: Can you manually open the Drive folder and Google Sheet?
4. **Review Error Messages**: Look for specific error details in Apps Script logs

## 🔧 Common Issues & Solutions

### Menu Not Appearing

#### Symptom
"Breeze Update" menu is missing from Google Sheets menu bar.

#### Solutions
1. **Refresh the page** (Ctrl+F5 or Cmd+Shift+R)
2. **Check script binding**:
   - Go to `Extensions → Apps Script`
   - Should open your GCG project (not a new untitled project)
   - If it opens a new project, you need to create a container-bound script
3. **Run manually**: Execute `onOpen()` function from Apps Script
4. **Authorization**: Grant permissions if prompted

#### Root Causes
- Script not bound to the sheet
- Missing authorization
- Code errors preventing menu creation

---

### Health Check Failures

#### Drive Access Failed
```
❌ Google Drive access failed: Exception: Invalid argument
```

**Solutions**:
- **Check Folder ID**: Verify `DRIVE_FOLDER_ID` in Script Properties
- **Check Permissions**: Ensure you can access the folder manually
- **Update Configuration**: Use the correct folder ID from the URL

#### Sheet Access Failed  
```
❌ Google Sheet access failed: Exception: Invalid argument
```

**Solutions**:
- **Check Sheet ID**: Verify `SHEET_ID` in Script Properties  
- **Check Permissions**: Ensure you can edit the sheet manually
- **Verify Tabs**: Confirm sheet has "GCG Members", "Not in a GCG", "Group List" tabs

#### Export Files Not Found
```
❌ Export files not found: No files found matching pattern "immanuelky-people"
```

**Solutions**:
- **Check File Names**: Files must contain "immanuelky-people" and "immanuelky-tags"
- **Check Location**: Files must be in the configured Google Drive folder
- **Check Format**: Files should be Google Sheets, not Excel
- **Re-upload**: Try uploading files again with correct naming

---

### Preview Report Issues

#### No Changes Showing (When You Expect Changes)
```
📊 Total real changes: 0
```

**Possible Causes**:
1. **Already Synced**: Your sheet might already be up to date
2. **Missing Person IDs**: Check if column A has Person IDs populated
3. **Name Mismatches**: Export vs sheet might have different name formats
4. **Inactive Filtering**: People marked inactive are automatically excluded

**Solutions**:
- **Check Person IDs**: Ensure column A of "GCG Members" has Person IDs
- **Review Action Steps**: Look for people marked as "inactive"
- **Compare Manually**: Check a few names between export and sheet

#### Too Many Changes Showing
```
📊 Total real changes: 200+
```

**Possible Causes**:
1. **Missing Person IDs**: System falling back to name matching
2. **Data Mismatch**: Export from wrong time period
3. **Group Name Differences**: Format differences not being normalized

**Solutions**:
- **Populate Person IDs**: Ensure all members have Person IDs in column A
- **Check Export Dates**: Verify you're using current export files
- **Review Group Names**: Look for format differences like "John Smith" vs "John Smith & Co-Leader"

---

### Data Parsing Errors

#### Synthetic Members Warning
```
📋 4 GCG members not in Active Members list
```

**This is Normal**: This indicates people in GCG tags who aren't in the Active Members export.

**Action Needed**:
1. **Review Preview Report Section 5**: See who these people are
2. **Check in Breeze**: Are they truly active or inactive?
3. **Update Breeze**: Either add to Active Members tag OR remove from GCG tag

#### Sheet Structure Errors
```
❌ Required columns (First Name, Last Name) not found
```

**Solutions**:
- **Check Headers**: Ensure "GCG Members" tab has proper column headers
- **Check Row Position**: Headers should be in row 1 or 2
- **Check Spelling**: Headers must contain "First" and "Last" (case insensitive)

---

### Permission Issues

#### Authorization Required
```
Exception: Authorization required
```

**Solutions**:
1. **Run from Apps Script**: Execute any function to trigger authorization
2. **Grant Permissions**: Click "Review Permissions" and approve
3. **Check Scopes**: Ensure script has Drive and Sheets permissions

#### Sharing Restrictions
```
Exception: You don't have permission to access this file
```

**Solutions**:
- **Check Drive Sharing**: Ensure folder is shared with your account
- **Check Sheet Sharing**: Ensure sheet is editable by your account
- **Use Consistent Account**: Make sure you're signed in with the correct Google account

---

### File Format Issues

#### Excel vs Google Sheets
The system expects **Google Sheets format**, not Excel files.

**Converting Excel to Google Sheets**:
1. **Upload Excel file** to Google Drive
2. **Right-click** → "Open with Google Sheets"
3. **File** → "Save as Google Sheets"
4. **Rename** with proper naming convention

#### Missing Tabs in Export
Some exports might be missing expected tabs.

**Solutions**:
- **Check Breeze Export Settings**: Ensure all tags are included
- **Verify Tag Names**: Look for GCG tags starting with "Gcg "
- **Manual Review**: Open the tags export to verify content

---

### Performance Issues

#### Slow Preview Generation
Preview reports can take 30-60 seconds for large datasets.

**Normal Behavior**:
- 35 GCG groups with 500+ members
- Multiple file reads and comparisons
- Complex data processing

**If Extremely Slow** (>2 minutes):
- **Check File Sizes**: Very large exports might cause timeouts
- **Check Internet**: Slow connection affects Google Drive access
- **Retry**: Sometimes a simple retry resolves temporary issues

#### Timeout Errors
```
Exception: Exceeded maximum execution time
```

**Solutions**:
- **Reduce Data**: Try with smaller test files first
- **Optimize Exports**: Export only necessary data from Breeze
- **Retry**: Apps Script timeouts can be intermittent

---

## 🔍 Debugging Tools

### Enable Detailed Logging
Add this to any function for detailed debugging:
```javascript
console.log('Debug point 1: Starting process');
// ... your code ...
console.log('Debug point 2: Completed step X');
```

### Manual Function Testing
Test individual components:
- `testConfig()` - Configuration and access
- `findLatestFile('ACTIVE_MEMBERS')` - File detection
- `parseRealGCGDataWithGCGMembers()` - Data parsing

### Check Script Properties
Verify configuration in Apps Script:
1. **Project Settings** → **Script Properties**
2. **Look for**: `GCG_FILE_CONFIG`
3. **Validate JSON**: Use a JSON validator tool

---

## 📞 When to Contact Support

### Before Contacting Support
Complete these steps:
- [ ] Run Health Check and save full output
- [ ] Try manual access to Drive folder and Google Sheet
- [ ] Check Apps Script execution transcript for errors
- [ ] Document exact steps that lead to the issue

### Information to Include
1. **Health Check Results**: Complete output from System Health Check
2. **Error Messages**: Exact text of any error messages
3. **Steps to Reproduce**: What you were doing when the issue occurred
4. **Recent Changes**: Any files uploaded or settings changed
5. **Screenshots**: Of any error dialogs or unexpected behavior

### Contact Information
- **Email**: sstringer@immanuelky.org
- **Subject**: "GCG Automation Support - [Brief Description]"
- **Include**: All information listed above

### Response Expectations
- **Acknowledgment**: Within 24 hours
- **Resolution**: Simple issues within 48 hours, complex issues may take longer
- **Follow-up**: You'll receive updates on progress for complex issues

---

## 🛠️ Advanced Troubleshooting

### Resetting the System
If everything is broken and you need to start fresh:

1. **Backup Current Data**: Export your current Google Sheet
2. **Create New Apps Script Project**: Start with clean container-bound script
3. **Reconfigure**: Set up Script Properties from scratch
4. **Test Step by Step**: Verify each component before proceeding

### Manual Data Verification
To check if the automation is working correctly:

1. **Pick a Test Person**: Find someone in both export and sheet
2. **Check Person ID**: Verify same ID in both places
3. **Check Group**: Verify same group assignment
4. **Check Inactive Status**: Look at Action Steps column

### Log Analysis
To understand what the system is doing:

1. **Open Apps Script Editor**
2. **View → Logs** or **View → Stackdriver Logging**
3. **Look for patterns** in error messages or processing steps
4. **Note timing** of when issues occur

---

*Remember: Most issues are configuration-related and can be resolved by carefully checking the setup steps.* 🔧
