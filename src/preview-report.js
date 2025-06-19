/**
 * Preview Report System for GCG Automation
 * Generates comprehensive preview before applying updates
 * Updated with inactive member enhancements
 */

/**
 * Generate comprehensive preview report with inactive member support
 * Main function that creates the preview sheet with all sections
 */
function generatePreviewReport() {
  console.log('📊 Generating comprehensive preview report...');
  
  try {
    // Get the latest data and changes - USE ENHANCED VERSION WITH INACTIVE SUPPORT
    const exportData = parseRealGCGDataWithInactiveMembers();
    const changes = enhancedCompareWithInactiveAwareness(exportData);
    
    // Create preview sheet
    const config = getConfig();
    const ss = SpreadsheetApp.openById(config.SHEET_ID);
    
    // Delete existing preview sheet if it exists
    try {
      const existingSheet = ss.getSheetByName('Breeze Update Preview');
      ss.deleteSheet(existingSheet);
    } catch (e) {
      // Sheet doesn't exist, that's fine
    }
    
    // Create new preview sheet
    const previewSheet = ss.insertSheet('Breeze Update Preview');
    
    // Add overall header with title and timestamp
    addReportHeader(previewSheet);
    
    // Build all sections of the report (shifted down by 1 row)
    buildGCGSummaryReport(previewSheet, exportData, changes);
    buildGroupByGroupReport(previewSheet, exportData, changes);
    buildEnhancedStatsReport(previewSheet, exportData, changes); // ENHANCED with inactive stats
    buildNotInGCGReport(previewSheet, exportData, changes); // Updated to use enhanced changes
    buildInactiveMembersInGCGsReport(previewSheet, exportData, changes); // NEW SECTION 5
    buildDataInconsistenciesReport(previewSheet, exportData); // NOW SECTION 6
    
    // Format the sheet for readability
    formatPreviewSheet(previewSheet);
    
    // Activate the preview sheet to keep focus on it - try multiple approaches
    ss.setActiveSheet(previewSheet);
    previewSheet.activate();
    
    console.log('✅ Preview report generated successfully');
    console.log('📋 Check the "Breeze Update Preview" tab to review all changes');
    
    return previewSheet;
    
  } catch (error) {
    console.error('❌ Failed to generate preview report:', error.message);
    throw error;
  }
}

/**
 * Add overall report header with title and timestamp
 */
function addReportHeader(sheet) {
  const now = new Date();
  const timestamp = now.toLocaleString();
  
  // Title in merged cells A1:I1
  sheet.getRange('A1:I1').merge();
  sheet.getRange('A1').setValue('Breeze Update Preview Report');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground('#e6f3ff');
  
  // Timestamp in merged cells A2:I2
  sheet.getRange('A2:I2').merge();
  sheet.getRange('A2').setValue(`Generated: ${timestamp}`);
  sheet.getRange('A2').setHorizontalAlignment('center')
    .setFontStyle('italic').setBackground('#f0f9ff');
}

/**
 * Build GCG Summary Report (Section 1)
 */
function buildGCGSummaryReport(sheet, exportData, changes) {
  const summaryData = [];
  
  exportData.groups.forEach(group => {
    const groupChanges = getGroupChanges(group, changes);
    summaryData.push([
      group.displayName,
      group.memberCount,
      groupChanges.additions.length,
      groupChanges.updates.length,
      groupChanges.removals.length
    ]);
  });
  
  // Headers for summary
  sheet.getRange(4, 1).setValue('Section 1: GCG Groups Summary');
  sheet.getRange(4, 1).setFontWeight('bold').setFontSize(12);
  
  const headerRow = 5;
  const headers = ['Group Name', 'Current Members', 'Additions', 'Updates', 'Removals'];
  headers.forEach((header, index) => {
    sheet.getRange(headerRow, index + 1).setValue(header);
    sheet.getRange(headerRow, index + 1).setFontWeight('bold').setBackground('#d4edda');
  });
  
  // Data rows
  if (summaryData.length > 0) {
    const dataRange = sheet.getRange(headerRow + 1, 1, summaryData.length, 5);
    dataRange.setValues(summaryData);
  }
}

/**
 * Build Group-by-Group Detailed Report (Section 2)
 */
function buildGroupByGroupReport(sheet, exportData, changes) {
  let currentRow = 10 + exportData.groups.length; // Dynamic positioning
  
  sheet.getRange(currentRow, 1).setValue('Section 2: Group-by-Group Changes');
  sheet.getRange(currentRow, 1).setFontWeight('bold').setFontSize(12);
  currentRow += 2;
  
  exportData.groups.forEach(group => {
    const groupChanges = getGroupChanges(group, changes);
    
    // Group header
    sheet.getRange(currentRow, 1).setValue(`${group.displayName} (${group.memberCount} members)`);
    sheet.getRange(currentRow, 1).setFontWeight('bold').setBackground('#e9ecef');
    currentRow++;
    
    // Additions section
    sheet.getRange(currentRow, 1).setValue('Additions');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow++;
    
    if (groupChanges.additions.length > 0) {
      // Add header row for additions
      sheet.getRange(currentRow, 1).setValue('Person ID');
      sheet.getRange(currentRow, 2).setValue('First');
      sheet.getRange(currentRow, 3).setValue('Last');
      sheet.getRange(currentRow, 1, 1, 3).setFontWeight('bold');
      currentRow++;
      
      groupChanges.additions.forEach(person => {
        sheet.getRange(currentRow, 1).setValue(person.personId);
        sheet.getRange(currentRow, 2).setValue(person.firstName);
        sheet.getRange(currentRow, 3).setValue(person.lastName);
        currentRow++;
      });
    } else {
      sheet.getRange(currentRow, 1).setValue('None');
      currentRow++;
    }
    
    // Deletions section
    sheet.getRange(currentRow, 1).setValue('Deletions');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow++;
    
    if (groupChanges.deletions.length > 0) {
      // Add header row for deletions
      sheet.getRange(currentRow, 1).setValue('Person ID');
      sheet.getRange(currentRow, 2).setValue('First');
      sheet.getRange(currentRow, 3).setValue('Last');
      sheet.getRange(currentRow, 1, 1, 3).setFontWeight('bold');
      currentRow++;
      
      groupChanges.deletions.forEach(person => {
        sheet.getRange(currentRow, 1).setValue(person.personId);
        sheet.getRange(currentRow, 2).setValue(person.firstName);
        sheet.getRange(currentRow, 3).setValue(person.lastName);
        currentRow++;
      });
    } else {
      sheet.getRange(currentRow, 1).setValue('None');
      currentRow++;
    }
    
    // Inactive section - updated to light orange background
    sheet.getRange(currentRow, 1).setValue('Currently Listed as Inactive');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setBackground('#ffe0b3'); // Light orange
    currentRow++;
    
    if (groupChanges.inactive && groupChanges.inactive.length > 0) {
      // Add header row for inactive
      sheet.getRange(currentRow, 1).setValue('Person ID');
      sheet.getRange(currentRow, 2).setValue('First');
      sheet.getRange(currentRow, 3).setValue('Last');
      sheet.getRange(currentRow, 4).setValue('Reason');
      sheet.getRange(currentRow, 1, 1, 4).setFontWeight('bold');
      currentRow++;
      
      groupChanges.inactive.forEach(person => {
        sheet.getRange(currentRow, 1).setValue(person.personId);
        sheet.getRange(currentRow, 2).setValue(person.firstName);
        sheet.getRange(currentRow, 3).setValue(person.lastName);
        sheet.getRange(currentRow, 4).setValue(person.inactiveReason);
        currentRow++;
      });
    } else {
      sheet.getRange(currentRow, 1).setValue('None');
      currentRow++;
    }
    
    // Blank row between groups
    currentRow++;
  });
}

/**
 * Build Enhanced Statistics Report with Inactive Member Data (Section 3)
 */
function buildEnhancedStatsReport(sheet, exportData, changes) {
  sheet.getRange(7, 1).setValue('Section 3: Statistics');
  sheet.getRange(7, 1).setFontWeight('bold').setFontSize(12);
  
  // ENHANCED: Add inactive member statistics to rows 7-9
  if (changes.inactiveProcessing) {
    const stats = changes.inactiveProcessing;
    sheet.getRange(7, 1).setValue(`Total Active Members: ${exportData.summary?.totalActiveMembers || 0}`);
    sheet.getRange(8, 1).setValue(`Total Inactive Members: ${stats.totalInactiveMembers || 0}`);
    sheet.getRange(9, 1).setValue(`Inactive Members in GCGs: ${stats.inactiveMembersInGCG || 0} (${stats.alreadyKnownInactive || 0} already flagged + ${stats.newlyDiscoveredInactive || 0} new)`);
    sheet.getRange(10, 1).setValue(`Inactive Filtered from "Not in GCG": ${stats.inactiveFilteredFromNotInGCG || 0}`);
  } else {
    // Fallback to basic statistics
    sheet.getRange(7, 1).setValue(`Total Active Members: ${exportData.summary?.totalActiveMembers || 0}`);
    sheet.getRange(8, 1).setValue(`Total GCG Groups: ${exportData.summary?.totalGroups || 0}`);
    sheet.getRange(9, 1).setValue(`Active Members in GCGs: ${exportData.summary?.activeMembersInGCG || 0}`);
  }
}

/**
 * Build "Not in GCG" Report (Section 4)
 */
function buildNotInGCGReport(sheet, exportData, changes) {
  // Section 4 header at fixed location F10
  sheet.getRange(10, 6, 1, 5).setValue('Section 4: Proposed Updates to Active Members Not in a GCG');
  sheet.getRange(10, 6, 1, 5).setFontWeight('bold').setBackground('#e1c6e7'); // Light purple
  
  let currentRow = 11;
  
  // Use enhanced changes with inactive filtering
  const notInGCGChanges = changes.notInGCGChanges || { additions: [], deletions: [] };
  
  // Additions section
  sheet.getRange(currentRow, 6).setValue('Additions to "Not in GCG" Tab:');
  sheet.getRange(currentRow, 6).setFontWeight('bold');
  currentRow++;
  
  if (notInGCGChanges.additions.length > 0) {
    // Headers
    sheet.getRange(currentRow, 6).setValue('Person ID');
    sheet.getRange(currentRow, 7).setValue('Family Name');
    sheet.getRange(currentRow, 8).setValue('First Name');
    sheet.getRange(currentRow, 9).setValue('Last Name');
    sheet.getRange(currentRow, 6, 1, 4).setFontWeight('bold');
    currentRow++;
    
    notInGCGChanges.additions.forEach(person => {
      sheet.getRange(currentRow, 6).setValue(person.personId);
      sheet.getRange(currentRow, 7).setValue(person.familyName || person.fullName);
      sheet.getRange(currentRow, 8).setValue(person.firstName);
      sheet.getRange(currentRow, 9).setValue(person.lastName);
      currentRow++;
    });
  } else {
    sheet.getRange(currentRow, 6).setValue('None');
    currentRow++;
  }
  
  currentRow++; // Buffer row
  
  // Deletions section
  sheet.getRange(currentRow, 6).setValue('Deletions from "Not in GCG" Tab:');
  sheet.getRange(currentRow, 6).setFontWeight('bold');
  currentRow++;
  
  if (notInGCGChanges.deletions.length > 0) {
    // Headers
    sheet.getRange(currentRow, 6).setValue('Person ID');
    sheet.getRange(currentRow, 7).setValue('First Name');
    sheet.getRange(currentRow, 8).setValue('Last Name');
    sheet.getRange(currentRow, 9).setValue('Reason');
    sheet.getRange(currentRow, 6, 1, 4).setFontWeight('bold');
    currentRow++;
    
    notInGCGChanges.deletions.forEach(person => {
      sheet.getRange(currentRow, 6).setValue(person.personId);
      sheet.getRange(currentRow, 7).setValue(person.firstName);
      sheet.getRange(currentRow, 8).setValue(person.lastName);
      sheet.getRange(currentRow, 9).setValue(person.reason || 'Unknown');
      currentRow++;
    });
  } else {
    sheet.getRange(currentRow, 6).setValue('None');
    currentRow++;
  }
}

/**
 * NEW: Build Inactive Members in GCGs Report (Section 5)
 */
function buildInactiveMembersInGCGsReport(sheet, exportData, changes) {
  // Calculate where Section 4 ends to position Section 5
  const notInGCGChanges = changes.notInGCGChanges || { additions: [], deletions: [] };
  const section4Rows = 6 + notInGCGChanges.additions.length + notInGCGChanges.deletions.length;
  const section5StartRow = 10 + section4Rows + 2; // +2 for buffer
  
  let currentRow = section5StartRow;
  
  // Section 5 Header - Light orange background
  sheet.getRange(currentRow, 6, 1, 5).setValue('Section 5: Newly Discovered Inactive Members in GCGs (Recommended for Removal)');
  sheet.getRange(currentRow, 6, 1, 5).setFontWeight('bold').setBackground('#ffe0b3'); // Light orange
  currentRow++;
  
  // Check if there are newly discovered inactive members in GCGs
  if (changes.newInactiveInGCGs && changes.newInactiveInGCGs.length > 0) {
    // Column headers
    sheet.getRange(currentRow, 6).setValue('Person ID');
    sheet.getRange(currentRow, 7).setValue('First Name');
    sheet.getRange(currentRow, 8).setValue('Last Name');
    sheet.getRange(currentRow, 9).setValue('GCG Group');
    sheet.getRange(currentRow, 10).setValue('Inactive Reason');
    sheet.getRange(currentRow, 6, 1, 5).setFontWeight('bold');
    currentRow++;
    
    // Data rows
    changes.newInactiveInGCGs.forEach(member => {
      sheet.getRange(currentRow, 6).setValue(member.personId);
      sheet.getRange(currentRow, 7).setValue(member.firstName);
      sheet.getRange(currentRow, 8).setValue(member.lastName);
      sheet.getRange(currentRow, 9).setValue(member.gcgAssignment.groupName);
      sheet.getRange(currentRow, 10).setValue(member.inactiveReason || 'Unknown');
      currentRow++;
    });
    
    // Add helpful note
    currentRow++;
    sheet.getRange(currentRow, 6).setValue('Note: People already flagged as inactive in comments are not shown here.');
    sheet.getRange(currentRow, 6, 1, 5).setFontStyle('italic').setFontColor('#666666');
    currentRow++;
    
  } else {
    // No newly discovered inactive members in GCGs
    sheet.getRange(currentRow, 6).setValue('✅ No newly discovered inactive members in GCGs - all known cases already flagged');
    currentRow++;
  }
  
  return currentRow + 1; // Return next available row for Section 6
}

/**
 * Build Data Inconsistencies Report (NOW Section 6)
 */
function buildDataInconsistenciesReport(sheet, exportData) {
  // Calculate where Section 5 ends to position Section 6
  const changes = { newInactiveInGCGs: exportData.newInactiveInGCGs || [] };
  const section5EndRow = buildInactiveMembersInGCGsReport(sheet, exportData, changes);
  const section6StartRow = section5EndRow + 1; // +1 for buffer
  
  let currentRow = section6StartRow;
  
  sheet.getRange(currentRow, 6, 1, 5).setValue('Section 6: Data Inconsistencies (Review Required)');
  sheet.getRange(currentRow, 6, 1, 5).setFontWeight('bold').setBackground('#f8d7da'); // Light red
  currentRow++;
  
  // Check for synthetic members (people in GCGs but not in active members export)
  if (exportData.missingFromActive && exportData.missingFromActive.length > 0) {
    sheet.getRange(currentRow, 6).setValue('People in GCGs but not in Active Members export:');
    sheet.getRange(currentRow, 6).setFontWeight('bold');
    currentRow++;
    
    // Headers
    sheet.getRange(currentRow, 6).setValue('Person ID');
    sheet.getRange(currentRow, 7).setValue('First Name');
    sheet.getRange(currentRow, 8).setValue('Last Name');
    sheet.getRange(currentRow, 9).setValue('GCG Group');
    sheet.getRange(currentRow, 6, 1, 4).setFontWeight('bold');
    currentRow++;
    
    exportData.missingFromActive.forEach(person => {
      sheet.getRange(currentRow, 6).setValue(person.personId);
      sheet.getRange(currentRow, 7).setValue(person.firstName);
      sheet.getRange(currentRow, 8).setValue(person.lastName);
      sheet.getRange(currentRow, 9).setValue(person.gcgGroup);
      currentRow++;
    });
  } else {
    sheet.getRange(currentRow, 6).setValue('✅ No data inconsistencies found');
    currentRow++;
  }
}

/**
 * Helper function to get changes for a specific group
 */
function getGroupChanges(group, changes) {
  const groupChanges = {
    additions: [],
    updates: [],
    removals: [],
    inactive: []
  };
  
  // Filter changes for this specific group
  if (changes.additions) {
    groupChanges.additions = changes.additions.filter(person => 
      person.gcgStatus && person.gcgStatus.groupName === group.displayName
    );
  }
  
  if (changes.updates) {
    groupChanges.updates = changes.updates.filter(person => 
      person.gcgStatus && person.gcgStatus.groupName === group.displayName
    );
  }
  
  if (changes.removals) {
    groupChanges.removals = changes.removals.filter(person => 
      person.gcgGroup === group.displayName
    );
  }
  
  // Add inactive members for this group
  if (changes.newInactiveInGCGs) {
    groupChanges.inactive = changes.newInactiveInGCGs.filter(member =>
      member.gcgAssignment && member.gcgAssignment.groupName === group.displayName
    );
  }
  
  return groupChanges;
}

/**
 * Format the preview sheet for better readability
 */
function formatPreviewSheet(sheet) {
  // Set column widths
  sheet.setColumnWidth(1, 200); // Person ID / Group Name
  sheet.setColumnWidth(2, 120); // First Name
  sheet.setColumnWidth(3, 120); // Last Name
  sheet.setColumnWidth(4, 100); // Count/Reason
  sheet.setColumnWidth(5, 100); // Extra space
  sheet.setColumnWidth(6, 200); // Section headers
  sheet.setColumnWidth(7, 120); // First Name
  sheet.setColumnWidth(8, 120); // Last Name
  sheet.setColumnWidth(9, 150); // Group/Reason
  sheet.setColumnWidth(10, 120); // Inactive Reason
  
  // Freeze header rows
  sheet.setFrozenRows(3);
  
  // Set borders for better visual separation
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 0 && lastCol > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).setBorder(true, true, true, true, true, true);
  }
}
