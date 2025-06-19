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

**Modify the existing preview report generation to include inactive member insights:**

```javascript
/**
 * Enhanced preview report with inactive member insights integrated into Report Sheet
 * Update the existing generatePreviewReport() function to include these enhancements
 */

// 1. UPDATE SECTION 3 (Statistics) - Add inactive member statistics
// Modify the existing statistics section to include inactive stats in rows 7-9

function addInactiveStatsToSection3(sheet, comparisonData) {
  const stats = comparisonData.inactiveProcessing || {};
  
  // Add inactive statistics to existing Section 3 (rows 7-9)
  sheet.getRange(7, 1).setValue(`Total Inactive Members: ${stats.totalInactiveMembers || 0}`);
  sheet.getRange(8, 1).setValue(`Inactive Members in GCGs: ${stats.inactiveMembersInGCG || 0}`);
  sheet.getRange(9, 1).setValue(`Inactive Filtered from "Not in GCG": ${stats.inactiveFilteredFromNotInGCG || 0}`);
}

// 2. ADD NEW SECTION 5 - Inactive Members Still in GCGs
// Insert this BEFORE the current Section 5 (Data Inconsistencies)

function addSection5InactiveMembersInGCGs(sheet, comparisonData, startRow) {
  let currentRow = startRow;
  
  // Section 5 Header - Light orange background like other warning sections
  sheet.getRange(currentRow, 6, 1, 5).setValue('Section 5: Inactive Members Still in GCGs (Recommended for Removal)');
  sheet.getRange(currentRow, 6, 1, 5).setFontWeight('bold').setBackground('#ffe0b3'); // Light orange
  currentRow++;
  
  // Check if there are inactive members in GCGs
  if (comparisonData.inactiveInGCGs && comparisonData.inactiveInGCGs.length > 0) {
    // Column headers
    sheet.getRange(currentRow, 6).setValue('Person ID');
    sheet.getRange(currentRow, 7).setValue('First Name');
    sheet.getRange(currentRow, 8).setValue('Last Name');
    sheet.getRange(currentRow, 9).setValue('GCG Group');
    sheet.getRange(currentRow, 10).setValue('Inactive Reason');
    sheet.getRange(currentRow, 6, 1, 5).setFontWeight('bold');
    currentRow++;
    
    // Data rows
    comparisonData.inactiveInGCGs.forEach(member => {
      sheet.getRange(currentRow, 6).setValue(member.personId);
      sheet.getRange(currentRow, 7).setValue(member.firstName);
      sheet.getRange(currentRow, 8).setValue(member.lastName);
      sheet.getRange(currentRow, 9).setValue(member.gcgAssignment.groupName);
      sheet.getRange(currentRow, 10).setValue(member.inactiveReason || 'Unknown');
      currentRow++;
    });
  } else {
    // No inactive members in GCGs
    sheet.getRange(currentRow, 6).setValue('✅ No inactive members found in GCGs - no cleanup needed');
    currentRow++;
  }
  
  // Return the next available row (with buffer)
  return currentRow + 1; // +1 for buffer row
}

// 3. UPDATE THE MAIN PREVIEW REPORT FUNCTION
// Modify your existing generatePreviewReport function to call these new functions

/**
 * Enhanced preview report generation - UPDATE YOUR EXISTING FUNCTION
 * Add these calls to your existing generatePreviewReport() function:
 */
function enhanceExistingPreviewReport(comparisonData) {
  // Your existing preview report logic...
  
  // After creating the Report sheet and adding Sections 1-3:
  
  // 1. Enhance Section 3 with inactive statistics (if inactive data available)
  if (comparisonData.inactiveProcessing) {
    addInactiveStatsToSection3(reportSheet, comparisonData);
  }
  
  // 2. After Section 4 "Proposed Updates to Active Members Not in a GCG":
  // Calculate where Section 5 should start (after Section 4 ends + 1 buffer row)
  const section4EndRow = calculateSection4EndRow(comparisonData); // You'll need to implement this
  const section5StartRow = section4EndRow + 2; // +1 for buffer, +1 for actual start
  
  // 3. Add new Section 5 (Inactive Members in GCGs)
  const section6StartRow = addSection5InactiveMembersInGCGs(reportSheet, comparisonData, section5StartRow);
  
  // 4. Move existing "Data Inconsistencies" to Section 6
  // Update your existing data inconsistencies logic to start at section6StartRow
  addSection6DataInconsistencies(reportSheet, comparisonData, section6StartRow);
}

// 4. GCG GROUPS SHEET ENHANCEMENT - SKIPPED FOR NOW
// Note: Column D already contains "Inactive count" 
// Need to clarify what current Column D represents before adding enhancements
// Focus on Report Sheet integration (Sections 3 and 5) for now
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
