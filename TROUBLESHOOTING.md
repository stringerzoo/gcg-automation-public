# GCG Automation Troubleshooting Guide

Comprehensive troubleshooting guide for the GCG Automation System.

## 🚨 Quick Diagnostic Steps

### 1. Run Health Check First
Before troubleshooting, always run the system health check:
1. **Open your Google Sheet**
2. **Go to**: Breeze Update → System Health Check
3. **Review results**: Look for ❌ red X marks indicating issues
4. **Focus on failures**: Address critical issues first

### 2. Check System Status
Use the built-in diagnostics to understand what's working:
- ✅ **Green**: Component working properly
- ⚠️ **Yellow**: Warning or minor issue
- ❌ **Red**: Critical failure requiring attention

## 🔧 Common Issues and Solutions

### Menu System Issues

#### "Breeze Update" Menu Not Appearing
**Symptoms**: No custom menu visible in Google Sheets

**Causes & Solutions**:
1. **Apps Script not bound to sheet**
   - Check: Go to Extensions → Apps Script
   - Solution: Ensure the script is opened FROM your Google Sheet, not standalone
   - Test: URL should show your sheet ID in the Apps Script interface

2. **Missing or incomplete code files**
   - Check: Verify all 6 files exist in Apps Script project
   - Solution: Copy missing files from `/src/` folder
   - Required files: `config.js`, `smart-file-detection.js`, `google-sheets-parser.js`, `comparison-engine.js`, `preview-report.js`, `menu-system.js`

3. **Code errors preventing menu creation**
   - Check: Look for error messages in Apps Script logs
   - Solution: Fix syntax errors, save all files
   - Test: Manually run `onOpen()` function

4. **Browser caching issues**
   - Solution: Refresh the Google Sheet page (F5 or Ctrl+R)
   - Alternative: Open sheet in incognito/private browsing mode

#### Menu Items Not Working
**Symptoms**: Menu appears but items don't respond or show errors

**Solutions**:
1. **Check function names match exactly** in menu-system.js
2. **Verify permissions** are granted for all required services
3. **Review Apps Script logs** for specific error messages
4. **Re-run authorization** by executing any function manually

### File Detection Issues

#### "Export Files Not Found" Error
**Symptoms**: Health check shows files not found, or preview generation fails

**Troubleshooting Steps**:
1. **Verify file location**
   ```
   Check: Are files in the correct Google Drive folder?
   Solution: Upload files to folder specified in DRIVE_FOLDER_ID
   ```

2. **Check file naming patterns**
   ```
   Required patterns:
   - Active Members: must contain "immanuelky-people"
   - Tags Export: must contain "immanuelky-tags"
   
   Examples:
   ✅ immanuelky-people-06-13-2025.xlsx
   ✅ immanuelky-tags-06-13-2025.xlsx  
   ❌ active-members-export.xlsx
   ❌ church-tags.xlsx
   ```

3. **Verify file format**
   ```
   Required: Google Sheets format (converted from Excel)
   Not supported: Raw .xlsx files uploaded to Drive
   Solution: Export directly to Google Sheets from Breeze
   ```

4. **Check folder permissions**
   ```
   Test: Can you open the folder URL manually?
   Solution: Ensure your account has edit access to the folder
   ```

#### "Files Found But Can't Read Data" Error
**Symptoms**: Files detected but parsing fails

**Solutions**:
1. **Verify export completeness**
   - Open files manually and check they contain expected data
   - Ensure headers are present and match expected format
   - Verify Person IDs are populated

2. **Check export settings in Breeze**
   - Export to Google Sheets (not Excel download then upload)
   - Include all required fields
   - Ensure no corruption during export

### Data Parsing Issues

#### "No Person IDs Found" Error
**Symptoms**: Parser can't find Person ID column

**Solutions**:
1. **Check Active Members export**
   ```
   Required: "Breeze ID" or "Person ID" column
   Location: Should be first column (Column A)
   Format: Numeric values (e.g., 29771102)
   ```

2. **Verify header format**
   ```
   Expected headers: "Breeze ID", "First Name", "Last Name"
   Case sensitive: No, but spelling must be close
   Extra spaces: System handles automatically
   ```

#### "No GCG Groups Found" Error
**Symptoms**: Tags export contains no recognizable GCG data

**Solutions**:
1. **Check GCG naming convention**
   ```
   Expected format: Sheets named "Gcg [Leader Name]"
   Examples:
   ✅ "Gcg Aaron White"
   ✅ "Gcg Gene Cone & Scott Stringer"
   ❌ "Aaron White Group"
   ❌ "GCG Leadership" (excluded as admin)
   ```

2. **Verify export includes GCG tags**
   - Ensure Breeze export includes all GCG-related tags
   - Check that individual GCG sheets exist in the exported file
   - Verify sheets contain member data

#### "Family Data Missing" Warning
**Symptoms**: Preview shows "TBD" for Family ID/Role

**Solutions**:
1. **Update Breeze export settings**
   ```
   Required fields in Active Members export:
   - Family ID
   - Family Role (Head of Household, Spouse, Adult, Child)
   ```

2. **Check family data completeness**
   - Not all members need family assignments
   - Single individuals may show empty family data (normal)
   - Only affects "Not in GCG" family grouping

### Preview Report Issues

#### "Preview Report Generation Failed" Error
**Symptoms**: Preview report doesn't generate or shows incomplete data

**Solutions**:
1. **Check data consistency**
   ```
   Common issues:
   - Missing Person IDs in GCG Members tab
   - Mismatched data between exports and current sheet
   - Corrupted export files
   ```

2. **Verify sheet structure**
   ```
   Required tabs in Google Sheet:
   - "GCG Members" (with Person ID in column A)
   - "Not in a GCG" (with proper family structure)
   - "Group List" (with group leader names)
   ```

3. **Test with simplified data**
   - Try with a smaller export file to isolate issues
   - Check specific groups that might have formatting problems

#### Preview Report Shows Unexpected Results
**Symptoms**: Large numbers of changes, incorrect family grouping, or "TBD" values

**Common Causes & Solutions**:

1. **Group name format differences**
   ```
   Problem: Current sheet has "Gene Cone", export has "Gene Cone & Scott Stringer"
   Solution: System should normalize automatically
   Expected: These should NOT show as updates
   Check: Look for pattern of co-leader additions in preview
   ```

2. **Inactive member handling**
   ```
   Problem: Preview shows people who should be filtered out
   Solution: Mark inactive members in "Action Steps / Comments" column
   Keywords: "inactive", "moved away", "left church", "transferred"
   ```

3. **Family grouping issues**
   ```
   Problem: Multiple family members listed instead of one representative
   Cause: Missing Family ID or Family Role data
   Solution: Ensure Breeze export includes family information
   ```

### Data Comparison Issues

#### "Too Many Changes Detected" Warning
**Symptoms**: Preview shows hundreds of changes when expecting few

**Troubleshooting Steps**:
1. **Check for duplicate functions**
   ```
   Problem: Multiple versions of comparison functions in code
   Solution: Search for duplicate function names in Apps Script
   Remove: Old or test versions of functions
   ```

2. **Verify normalization working**
   ```
   Test: Look for patterns in the changes
   Expected: Most "updates" should be filtered out by normalization
   Problem: Co-leader format differences counted as real changes
   ```

3. **Review data freshness**
   ```
   Check: Are you comparing current exports with current sheet?
   Problem: Comparing old export data with updated sheet
   Solution: Ensure exports are recent and complete
   ```

#### "Person ID Mismatches" Error
**Symptoms**: Same people showing as both additions and deletions

**Solutions**:
1. **Check Person ID consistency**
   ```
   Verify: Person IDs match exactly between exports and sheet
   Format: Should be numeric strings (e.g., "29771102")
   Common issue: Leading zeros or formatting differences
   ```

2. **Review data export quality**
   ```
   Test: Open exports manually and verify data integrity
   Check: No duplicate Person IDs or missing values
   Solution: Re-export from Breeze if data issues found
   ```

### Update Application Issues

#### "Updates Failed to Apply" Error
**Symptoms**: Preview works but applying updates fails

**Solutions**:
1. **Check sheet permissions**
   ```
   Required: Edit access to the Google Sheet
   Test: Can you manually edit cells in the sheet?
   Solution: Ensure proper permissions for your account
   ```

2. **Verify sheet structure integrity**
   ```
   Check: All required tabs exist and have expected headers
   Problem: Headers moved or renamed since setup
   Solution: Restore proper header structure
   ```

3. **Review Apps Script quotas**
   ```
   Issue: Large updates may hit execution time limits
   Solution: Break updates into smaller batches
   Alternative: Run updates during low-usage periods
   ```

## 🔍 Advanced Debugging

### Enable Detailed Logging
For complex issues, enable detailed logging:

1. **Add debug logging to functions**
   ```javascript
   console.log('🔍 Debug point: Variable value =', variableName);
   ```

2. **Run functions manually in Apps Script**
   - Select function from dropdown
   - Click Run button
   - Review execution transcript for errors

3. **Check execution transcript**
   - Look for specific error messages
   - Note line numbers where failures occur
   - Review stack traces for function call paths

### Test Individual Components

#### Test File Detection
```javascript
// Run this in Apps Script to test file detection
function testFileDetection() {
  try {
    const activeFile = findLatestFile('ACTIVE_MEMBERS');
    const tagsFile = findLatestFile('TAGS_EXPORT');
    console.log('✅ Active Members file:', activeFile.getName());
    console.log('✅ Tags file:', tagsFile.getName());
  } catch (error) {
    console.error('❌ File detection failed:', error.message);
  }
}
```

#### Test Data Parsing
```javascript
// Run this to test data parsing specifically
function testDataParsing() {
  try {
    const exportData = parseRealGCGDataWithGCGMembers();
    console.log('✅ Parsed data summary:', exportData.summary);
  } catch (error) {
    console.error('❌ Data parsing failed:', error.message);
  }
}
```

#### Test Comparison Logic
```javascript
// Run this to test comparison without preview generation
function testComparison() {
  try {
    const exportData = parseRealGCGDataWithGCGMembers();
    const changes = enhancedCompareWithPersonIds(exportData);
    console.log('✅ Comparison results:', changes);
  } catch (error) {
    console.error('❌ Comparison failed:', error.message);
  }
}
```

## 📞 Getting Additional Help

### When to Contact Support
Contact the system administrator when:
- Health Check shows multiple critical failures
- Data appears corrupted or inconsistent
- Updates are producing unexpected results
- Error messages mention system configuration issues

### Information to Provide
When reporting issues, include:
1. **Health Check results** (copy/paste full output)
2. **Specific error messages** from Apps Script logs
3. **Steps taken** leading to the issue
4. **Expected vs actual behavior**
5. **Timing** (when did this work last?)

### Support Contacts
- **Primary**: sstringer@immanuelky.org
- **Emergency**: Contact church IT administrator
- **Documentation**: Review GitHub repository for latest updates

## 🛡️ Prevention Tips

### Regular Maintenance
1. **Monthly Health Checks**: Run before each update cycle
2. **Keep Backups**: Maintain copies of working configurations
3. **Test Changes**: Use preview reports to validate before applying
4. **Monitor Data Quality**: Watch for trends in inconsistencies

### Best Practices
1. **Consistent Export Process**: Use same Breeze export settings each time
2. **File Naming**: Follow established date naming conventions
3. **Inactive Marking**: Consistently mark inactive members before updates
4. **Review Previews**: Always examine preview reports thoroughly

### Early Warning Signs
Watch for these indicators of potential issues:
- **Increasing number of "synthetic members"** (GCG members not in Active list)
- **Growing family grouping inconsistencies**
- **Preview reports showing unexpected change volumes**
- **Health Check warnings becoming more frequent**

## 📈 Performance Optimization

### For Large Data Sets
If working with very large membership data:

1. **Increase timeout settings**
   ```javascript
   // Add to config.js if needed
   const PROCESSING_TIMEOUT = 300; // 5 minutes
   ```

2. **Process in batches**
   - Handle large GCG groups separately
   - Split family processing across multiple operations
   - Use pagination for large preview reports

3. **Optimize frequency**
   - Run full updates monthly, not weekly
   - Use quick checks for interim status
   - Focus on changed groups only when possible

This troubleshooting guide should help resolve most issues you'll encounter. Remember that the system includes multiple safety checks and validation steps to prevent data loss, so when in doubt, review the preview report carefully and contact support if needed.
