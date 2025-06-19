# 🔄 Inactive Member Enhancement - Complete Implementation

## 🎯 Overview
This enhancement adds comprehensive inactive member support to your GCG Automation system, providing:
- **Cleaner data analysis** - distinguish "moved to inactive" from "data error"
- **Enhanced preview reports** - new inactive member sections and statistics
- **Smarter processing** - automatic filtering and better action items

## 📋 Implementation Plan

### Phase 1: File Detection Enhancement
**File:** `smart-file-detection.js`

**Add inactive file pattern to config:**
```javascript
// Add to updateConfigForSmartDetection() function
FILE_PATTERNS: {
  ACTIVE_MEMBERS: {
    contains: 'immanuelky-people-active',
    excludes: [],
    description: 'Active Members Export'
  },
  INACTIVE_MEMBERS: {  // 🆕 NEW
    contains: 'immanuelky-people-inactive',
    excludes: [],
    description: 'Inactive Members Export'
  },
  TAGS_EXPORT: {
    contains: 'immanuelky-tags',
    excludes: [],
    description: 'Tags Export'
  }
}
```

**Add new function:**
```javascript
/**
 * Parse full GCG data including inactive members
 * @returns {Object} Complete dataset with active, inactive, and GCG data
 */
function parseRealGCGDataWithInactiveMembers() {
  console.log('🎯 Parsing FULL GCG data (active + inactive + tags)...');
  
  try {
    // Parse active members
    const activeMembersFile = findLatestFile('ACTIVE_MEMBERS');
    const activeMembersResult = parseActiveMembersSheet(activeMembersFile);
    
    // Parse inactive members (gracefully handle missing file)
    let inactiveMembersResult = { members: [], totalCount: 0 };
    try {
      const inactiveMembersFile = findLatestFile('INACTIVE_MEMBERS');
      inactiveMembersResult = parseInactiveMembersSheet(inactiveMembersFile);
      console.log(`✅ Parsed ${inactiveMembersResult.totalCount} inactive members`);
    } catch (error) {
      console.log(`⚠️ No inactive members file found (continuing without): ${error.message}`);
    }
    
    // Parse GCG tags
    const tagsFile = findLatestFile('TAGS_EXPORT');
    const tagsResult = parseTagsSheet(tagsFile);
    
    // Combine and analyze
    const allMembers = [...activeMembersResult.members, ...inactiveMembersResult.members];
    
    return {
      activeMembers: activeMembersResult.members,
      inactiveMembers: inactiveMembersResult.members,
      allMembers: allMembers,
      groups: tagsResult.groups,
      assignments: tagsResult.assignments,
      summary: {
        totalActiveMembers: activeMembersResult.totalCount,
        totalInactiveMembers: inactiveMembersResult.totalCount,
        totalMembers: allMembers.length,
        totalGroups: tagsResult.totalGroups,
        activeMembersInGCG: activeMembersResult.members.filter(m => tagsResult.assignments[m.personId]).length,
        inactiveMembersInGCG: inactiveMembersResult.members.filter(m => tagsResult.assignments[m.personId]).length
      }
    };
    
  } catch (error) {
    console.error('❌ Full data parsing failed:', error.message);
    throw error;
  }
}
```

### Phase 2: Parser Enhancement
**File:** `google-sheets-parser.js`

**Add inactive member parsing function:**
```javascript
/**
 * Parse Inactive Members Google Sheet  
 * @param {GoogleAppsScript.Drive.File} file - Google Sheets file
 * @returns {Object} Parsed inactive members data
 */
function parseInactiveMembersSheet(file) {
  console.log('📊 Parsing Inactive Members Google Sheet...');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const sheets = spreadsheet.getSheets();
    
    console.log(`📄 Sheet: ${spreadsheet.getName()}`);
    console.log(`📝 Found ${sheets.length} tabs: ${sheets.map(s => s.getName()).join(', ')}`);
    
    const dataSheet = sheets[0];
    const data = dataSheet.getDataRange().getValues();
    
    if (data.length === 0) {
      throw new Error('No data found in Inactive Members sheet');
    }
    
    // First row should be headers
    const headers = data[0];
    console.log(`📋 Headers: ${headers.slice(0, 8).join(', ')}... (showing first 8)`);
    
    // Find important column indices
    const columnMap = {
      personId: findColumnIndex(headers, 'Breeze ID'),
      firstName: findColumnIndex(headers, 'First Name'),
      lastName: findColumnIndex(headers, 'Last Name'),
      nickname: findColumnIndex(headers, 'Nickname'),
      streetAddress: findColumnIndex(headers, 'Street Address'),
      city: findColumnIndex(headers, 'City'),
      state: findColumnIndex(headers, 'State'),
      zip: findColumnIndex(headers, 'Zip'),
      inactiveReason: findColumnIndex(headers, 'Status') // or 'Inactive Reason'
    };
    
    // Validate required columns
    if (columnMap.personId === -1 || columnMap.firstName === -1 || columnMap.lastName === -1) {
      throw new Error('Required columns (Breeze ID, First Name, Last Name) not found');
    }
    
    console.log('🔍 Column mapping:', JSON.stringify(columnMap));
    
    // Process inactive member data
    const members = [];
    let nicknameCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row[columnMap.personId]) continue;
      
      const nickname = columnMap.nickname >= 0 ? (row[columnMap.nickname] || '') : '';
      if (nickname.trim()) nicknameCount++;
      
      const member = {
        personId: String(row[columnMap.personId]),
        firstName: row[columnMap.firstName] || '',
        lastName: row[columnMap.lastName] || '',
        nickname: nickname,
        fullName: `${row[columnMap.firstName] || ''} ${row[columnMap.lastName] || ''}`.trim(),
        address: {
          street: row[columnMap.streetAddress] || '',
          city: row[columnMap.city] || '',
          state: row[columnMap.state] || '',
          zip: row[columnMap.zip] || ''
        },
        isActiveMember: false,  // 🆕 Mark as inactive
        isInactiveMember: true, // 🆕 Explicit inactive flag
        inactiveReason: columnMap.inactiveReason >= 0 ? (row[columnMap.inactiveReason] || 'Unknown') : 'Unknown',
        sourceRow: i + 1
      };
      
      members.push(member);
    }
    
    console.log(`✅ Parsed ${members.length} inactive members`);
    console.log(`📝 Found ${nicknameCount} inactive members with nicknames`);
    
    return {
      members: members,
      totalCount: members.length,
      nicknameCount: nicknameCount,
      headers: headers,
      sheetName: dataSheet.getName(),
      lastUpdated: file.getLastUpdated()
    };
    
  } catch (error) {
    console.error('❌ Error parsing Inactive Members sheet:', error.message);
    throw error;
  }
}
```

### Phase 3: Comparison Engine Enhancement
**File:** `comparison-engine.js`

**Enhance the main comparison function:**
```javascript
/**
 * Enhanced comparison with inactive member awareness
 * @param {Object} exportData - Full export data including inactive members
 * @returns {Object} Changes needed with inactive-aware processing
 */
function enhancedCompareWithInactiveAwareness(exportData) {
  console.log('🔍 Enhanced comparison with inactive awareness...');
  
  try {
    // Get standard GCG member changes (active members only)
    const gcgChanges = fixedCompareWithInactiveFiltering(exportData);
    
    // Calculate enhanced "Not in GCG" changes with inactive filtering
    const notInGCGChanges = calculateNotInGCGChangesWithInactiveFiltering(exportData);
    
    // Find inactive members currently in GCGs (should be cleaned up)
    const inactiveInGCGs = findInactiveMembersInGCGs(exportData);
    
    console.log('\n📊 INACTIVE-AWARE COMPARISON RESULTS:');
    console.log(`🔄 GCG Member changes (active only): ${gcgChanges.additions.length + gcgChanges.updates.length + gcgChanges.removals.length}`);
    console.log(`👥 Not in GCG changes (filtered): ${notInGCGChanges.additions.length + notInGCGChanges.deletions.length}`);
    console.log(`⚠️ Inactive members in GCGs: ${inactiveInGCGs.length}`);
    
    return {
      ...gcgChanges,
      notInGCGChanges: notInGCGChanges,
      inactiveInGCGs: inactiveInGCGs,
      inactiveProcessing: {
        totalInactiveMembers: exportData.summary.totalInactiveMembers,
        inactiveMembersInGCG: exportData.summary.inactiveMembersInGCG,
        inactiveFilteredFromNotInGCG: calculateInactiveFilteredCount(exportData)
      }
    };
    
  } catch (error) {
    console.error('❌ Enhanced inactive comparison failed:', error.message);
    throw error;
  }
}

/**
 * Find inactive members currently in GCGs
 * @param {Object} exportData - Full export data
 * @returns {Array} Inactive members with GCG assignments
 */
function findInactiveMembersInGCGs(exportData) {
  return exportData.inactiveMembers.filter(member => {
    const assignment = exportData.assignments[member.personId];
    return assignment; // Has GCG assignment
  }).map(member => ({
    ...member,
    gcgAssignment: exportData.assignments[member.personId]
  }));
}

/**
 * Enhanced "Not in GCG" calculation that filters out inactive members
 * @param {Object} exportData - Full export data
 * @returns {Object} Changes needed for "Not in GCG" section
 */
function calculateNotInGCGChangesWithInactiveFiltering(exportData) {
  // Get people not in GCGs from ACTIVE members only (exclude inactive)
  const activeNotInGCG = exportData.activeMembers.filter(member => {
    const gcgAssignment = exportData.assignments[member.personId];
    return !gcgAssignment; // Not in any GCG
  });
  
  // Apply family grouping logic to active members only
  const familyGroupedResults = applyFamilyGroupingLogic(activeNotInGCG);
  
  // Get current "Not in GCG" from sheet
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.SHEET_ID);
  const currentNotInGCG = getNotInGCGMembers(ss);
  
  // Calculate changes (inactive members automatically excluded)
  const additions = familyGroupedResults.filter(person => 
    !currentNotInGCG.some(current => current.personId === person.personId)
  );
  
  const deletions = currentNotInGCG.filter(current => {
    // Remove if: 1) Now in GCG, 2) No longer active, or 3) Now inactive
    const inGCG = exportData.assignments[current.personId];
    const stillActive = exportData.activeMembers.some(active => active.personId === current.personId);
    const nowInactive = exportData.inactiveMembers.some(inactive => inactive.personId === current.personId);
    
    return inGCG || !stillActive || nowInactive;
  });
  
  return {
    additions: additions,
    deletions: deletions,
    inactiveFilteredCount: exportData.inactiveMembers.filter(m => !exportData.assignments[m.personId]).length
  };
}
```

### Phase 4: Preview Report Enhancement  
**File:** `preview-report.js`

**Add inactive member sections to preview report:**
```javascript
/**
 * Enhanced preview report with inactive member insights
 * @param {Object} comparisonData - Comparison results with inactive data
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Preview report
 */
function generateEnhancedPreviewReport(comparisonData) {
  console.log('📊 Generating enhanced preview report with inactive insights...');
  
  // Create or get preview report spreadsheet (existing logic)
  const reportSS = createOrUpdatePreviewReport();
  
  // Add new "Inactive Summary" sheet
  addInactiveSummarySheet(reportSS, comparisonData);
  
  // Enhance existing GCG Groups sheet with inactive columns
  enhanceGCGGroupsSheetWithInactive(reportSS, comparisonData);
  
  // Enhance "Not in GCG" sheet with filtering explanation
  enhanceNotInGCGSheetWithFiltering(reportSS, comparisonData);
  
  console.log('✅ Enhanced preview report generated');
  return reportSS;
}

/**
 * Add new "Inactive Summary" sheet
 */
function addInactiveSummarySheet(reportSS, comparisonData) {
  let sheet;
  try {
    sheet = reportSS.getSheetByName('Inactive Summary');
    sheet.clear();
  } catch (error) {
    sheet = reportSS.insertSheet('Inactive Summary');
  }
  
  let currentRow = 1;
  
  // Header
  sheet.getRange(currentRow, 1).setValue('Inactive Member Analysis');
  sheet.getRange(currentRow, 1).setFontWeight('bold').setFontSize(14);
  currentRow += 2;
  
  // Statistics
  sheet.getRange(currentRow, 1).setValue('📊 Inactive Member Statistics');
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  currentRow++;
  
  const stats = comparisonData.inactiveProcessing || {};
  sheet.getRange(currentRow, 1).setValue(`Total Inactive Members: ${stats.totalInactiveMembers || 0}`);
  currentRow++;
  sheet.getRange(currentRow, 1).setValue(`Inactive Members in GCGs: ${stats.inactiveMembersInGCG || 0}`);
  currentRow++;
  sheet.getRange(currentRow, 1).setValue(`Inactive Filtered from "Not in GCG": ${stats.inactiveFilteredFromNotInGCG || 0}`);
  currentRow += 2;
  
  // Inactive members currently in GCGs (need cleanup)
  if (comparisonData.inactiveInGCGs && comparisonData.inactiveInGCGs.length > 0) {
    sheet.getRange(currentRow, 1).setValue('⚠️ Inactive Members Still in GCGs (Recommended for Removal)');
    sheet.getRange(currentRow, 1).setFontWeight('bold').setBackground('#ffe0b3');
    currentRow++;
    
    // Headers
    sheet.getRange(currentRow, 1).setValue('Person ID');
    sheet.getRange(currentRow, 2).setValue('First Name');
    sheet.getRange(currentRow, 3).setValue('Last Name');
    sheet.getRange(currentRow, 4).setValue('GCG Group');
    sheet.getRange(currentRow, 5).setValue('Inactive Reason');
    sheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold');
    currentRow++;
    
    comparisonData.inactiveInGCGs.forEach(member => {
      sheet.getRange(currentRow, 1).setValue(member.personId);
      sheet.getRange(currentRow, 2).setValue(member.firstName);
      sheet.getRange(currentRow, 3).setValue(member.lastName);
      sheet.getRange(currentRow, 4).setValue(member.gcgAssignment.groupName);
      sheet.getRange(currentRow, 5).setValue(member.inactiveReason || 'Unknown');
      currentRow++;
    });
  } else {
    sheet.getRange(currentRow, 1).setValue('✅ No inactive members found in GCGs');
    currentRow++;
  }
}

/**
 * Enhance existing GCG Groups sheet with inactive count column
 */
function enhanceGCGGroupsSheetWithInactive(reportSS, comparisonData) {
  try {
    const sheet = reportSS.getSheetByName('GCG Groups');
    if (!sheet) return;
    
    // Add "Inactive Count" column header (assuming it's column E)
    sheet.getRange(1, 5).setValue('Inactive Count');
    sheet.getRange(1, 5).setFontWeight('bold');
    
    // Add inactive counts for each group
    const groups = comparisonData.groups || [];
    const inactiveMembers = comparisonData.inactiveInGCGs || [];
    
    groups.forEach((group, index) => {
      const row = index + 2; // Assuming data starts at row 2
      const inactiveInThisGroup = inactiveMembers.filter(member => 
        member.gcgAssignment.groupName === group.displayName
      ).length;
      
      sheet.getRange(row, 5).setValue(inactiveInThisGroup);
      
      // Color-code if inactive members present
      if (inactiveInThisGroup > 0) {
        sheet.getRange(row, 5).setBackground('#ffe0b3'); // Light orange
      }
    });
    
  } catch (error) {
    console.warn('⚠️ Could not enhance GCG Groups sheet:', error.message);
  }
}
```

### Phase 5: Menu System Update
**File:** `menu-system.js`

**Update health check to include inactive file detection:**
```javascript
/**
 * Enhanced health check including inactive member support
 */
function performEnhancedHealthCheck() {
  const results = {
    configValid: false,
    filesFound: false,
    dataParsing: false,
    inactiveSupport: false,  // 🆕 NEW
    issues: 0,
    criticalIssues: 0,
    details: ''
  };
  
  try {
    // ... existing health checks ...
    
    // Test inactive file detection (new)
    try {
      const inactiveFile = findLatestFile('INACTIVE_MEMBERS');
      results.inactiveSupport = true;
      results.details += `✅ Inactive members file found: ${inactiveFile.getName()}\n`;
    } catch (error) {
      results.details += `⚠️ Inactive members file not found (optional): ${error.message}\n`;
      results.details += `   → System will work without inactive file\n`;
    }
    
    // ... rest of health check ...
    
  } catch (error) {
    results.issues++;
    results.criticalIssues++;
    results.details += `❌ Health check error: ${error.message}\n`;
  }
  
  return results;
}
```

## 🚀 Implementation Steps

### Step 1: Update File Detection
1. **Modify** `smart-file-detection.js`:
   - Add `INACTIVE_MEMBERS` pattern to config
   - Add `parseRealGCGDataWithInactiveMembers()` function

### Step 2: Add Inactive Parser
1. **Modify** `google-sheets-parser.js`:
   - Add `parseInactiveMembersSheet()` function

### Step 3: Enhance Comparison Engine
1. **Modify** `comparison-engine.js`:
   - Add `enhancedCompareWithInactiveAwareness()` function
   - Add `findInactiveMembersInGCGs()` function
   - Add `calculateNotInGCGChangesWithInactiveFiltering()` function

### Step 4: Enhance Preview Reports
1. **Modify** `preview-report.js`:
   - Add `generateEnhancedPreviewReport()` function
   - Add `addInactiveSummarySheet()` function
   - Add `enhanceGCGGroupsSheetWithInactive()` function

### Step 5: Update Menu System
1. **Modify** `menu-system.js`:
   - Update health check to detect inactive files
   - Update main processing to use enhanced functions

## 🧪 Testing Plan

### Phase 1: File Detection
1. Export inactive members from Breeze as `immanuelky-people-inactive-MM-DD-YYYY`
2. Upload to Google Drive folder
3. Run health check - should detect inactive file

### Phase 2: Data Processing
1. Run enhanced preview report
2. Verify inactive member statistics
3. Check "Inactive Summary" sheet

### Phase 3: Comparison Logic
1. Verify inactive members are filtered from "Not in GCG" additions
2. Check that inactive members in GCGs are flagged for cleanup
3. Confirm family grouping still works for active members

## 🎯 Benefits After Implementation

### ✅ **Cleaner Data Analysis**
- Distinguish "moved to inactive" from "data error"
- Reduce false positives in change recommendations
- Better action items for pastoral team

### 📊 **Enhanced Preview Reports**
- New "Inactive Summary" sheet with key insights
- "Inactive Count" column in GCG Groups summary
- Clear identification of cleanup needed

### 🧠 **Smarter Processing**
- Automatic filtering of inactive members from "Not in GCG"
- Clear flagging of inactive members who should be removed from GCGs
- Graceful handling when inactive file is missing

### 🔄 **Future-Proof Design**
- System works with or without inactive member files
- Easy to extend with additional member status types
- Maintains compatibility with existing workflow

This enhancement transforms your system from managing just active members to providing comprehensive member lifecycle management! 🚀