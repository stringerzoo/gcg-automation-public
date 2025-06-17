/**
 * Preview Report System for GCG Automation
 * Generates comprehensive preview before applying updates
 */

/**
 * Generate comprehensive preview report
 * Main function that creates the preview sheet with all sections
 */
function generatePreviewReport() {
  console.log('📊 Generating comprehensive preview report...');
  
  try {
    // Get the latest data and changes
    const exportData = parseRealGCGDataWithGCGMembers();
    const changes = fixedCompareWithInactiveFiltering(exportData);
    
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
    buildStatsReport(previewSheet, exportData);
    buildNotInGCGReport(previewSheet, exportData);
    buildDataInconsistenciesReport(previewSheet, exportData);
    
    // Format the sheet for readability
    formatPreviewSheet(previewSheet);
    
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
  
  sheet.getRange('A1').setValue(`Breeze Update Preview Report - ${timestamp}`);
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A1:I1').merge();
  sheet.getRange('A1').setBackground('#1976d2').setFontColor('white');
}

/**
 * Build Section 1: GCG Summary (A2:D40) - Added inactive count column
 * Shows truth data count vs current sheet count vs inactive count for each group
 */
function buildGCGSummaryReport(sheet, exportData, changes) {
  console.log('📝 Building GCG summary section...');
  
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.SHEET_ID);
  const allCurrentMembers = getGCGMembersWithPersonId(ss); // Include inactive
  
  // Headers (row 2 now) - Added column D
  sheet.getRange('A2').setValue('GCG Name');
  sheet.getRange('B2').setValue('Breeze Export Count');
  sheet.getRange('C2').setValue('GCG Members Tab Count');
  sheet.getRange('D2').setValue('Inactive Count');
  
  // Make headers bold
  sheet.getRange('A2:D2').setFontWeight('bold');
  
  // Sort groups alphabetically by leader name
  const groups = exportData.groups.sort((a, b) => a.leader.localeCompare(b.leader));
  
  groups.forEach((group, index) => {
    const row = index + 3; // Starting at row 3 now
    
    // Count current sheet members for this group (using normalized names)
    const allGroupMembers = allCurrentMembers.filter(m => 
      normalizeGroupName(m.group) === normalizeGroupName(group.displayName)
    );
    
    // Count active vs inactive members
    const activeGroupMembers = allGroupMembers.filter(m => !isMarkedInactive(m));
    const inactiveGroupMembers = allGroupMembers.filter(m => isMarkedInactive(m));
    
    // Add hyperlink to section 2 for group name
    const groupSectionRow = findGroupInSection2(group.displayName, groups);
    if (groupSectionRow > 0) {
      sheet.getRange(row, 1).setFormula(`=HYPERLINK("#A${groupSectionRow}", "${group.displayName}")`);
    } else {
      sheet.getRange(row, 1).setValue(group.displayName);
    }
    
    sheet.getRange(row, 2).setValue(group.memberCount); // Breeze export count
    sheet.getRange(row, 3).setValue(activeGroupMembers.length); // Active in sheet
    sheet.getRange(row, 4).setValue(inactiveGroupMembers.length); // Inactive in sheet
    
    // Highlight rows where Breeze count doesn't match active count
    if (group.memberCount !== activeGroupMembers.length) {
      sheet.getRange(row, 1, 1, 4).setBackground('#fff2cc'); // Light yellow
    }
    
    // Special highlighting for groups with inactive members
    if (inactiveGroupMembers.length > 0) {
      sheet.getRange(row, 4).setBackground('#ffcccc'); // Light red for inactive count
    }
  });
  
  console.log(`✅ Built summary for ${groups.length} groups with active/inactive breakdown`);
}

/**
 * Helper function to find group position in section 2
 */
function findGroupInSection2(groupName, groups) {
  // Section 2 starts at row 41, find this group's position
  let currentRow = 41;
  
  for (let i = 0; i < groups.length; i++) {
    if (normalizeGroupName(groups[i].displayName) === normalizeGroupName(groupName)) {
      return currentRow;
    }
    // Each group takes approximately 6 rows (header + additions + deletions + spacing)
    currentRow += 6;
  }
  
  return 0; // Not found
}

/**
 * Build Section 2: Group-by-Group Report (A41+) - Added inactive section
 * Shows additions, deletions, and inactive members for each group with changes
 */
function buildGroupByGroupReport(sheet, exportData, changes) {
  console.log('📝 Building group-by-group section...');
  
  let currentRow = 41; // Starting at row 41 now
  
  // Sort groups alphabetically
  const groups = exportData.groups.sort((a, b) => a.leader.localeCompare(b.leader));
  
  // Get all current members (including inactive) for inactive analysis
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.SHEET_ID);
  const allCurrentMembers = getGCGMembersWithPersonId(ss);
  
  groups.forEach(group => {
    // Find group-specific changes
    const groupChanges = findGroupSpecificChangesWithInactive(group, changes, exportData, allCurrentMembers);
    
    // Only include groups that have changes OR inactive members
    if (groupChanges.additions.length === 0 && 
        groupChanges.deletions.length === 0 && 
        groupChanges.inactive.length === 0) {
      return; // Skip groups with no changes or inactive members
    }
    
    // Group header
    sheet.getRange(currentRow, 1).setValue(group.displayName);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setBackground('#e1f5fe'); // Light blue
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
    
    // NEW: Inactive section
    sheet.getRange(currentRow, 1).setValue('Currently Listed as Inactive');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setBackground('#ffcccc'); // Light red
    currentRow++;
    
    if (groupChanges.inactive.length > 0) {
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
  
  const groupsWithActivity = groups.filter(group => {
    const groupChanges = findGroupSpecificChangesWithInactive(group, changes, exportData, allCurrentMembers);
    return groupChanges.additions.length > 0 || 
           groupChanges.deletions.length > 0 || 
           groupChanges.inactive.length > 0;
  });
  
  console.log(`✅ Built group-by-group report for ${groupsWithActivity.length} groups with changes or inactive members`);
}

/**
 * Enhanced helper function to find group-specific changes including inactive members
 * @param {Object} group - Group object from export data
 * @param {Object} changes - Overall changes object
 * @param {Object} exportData - Full export data
 * @param {Array} allCurrentMembers - All current members including inactive
 * @returns {Object} Group-specific additions, deletions, and inactive members
 */
function findGroupSpecificChangesWithInactive(group, changes, exportData, allCurrentMembers) {
  const groupAdditions = [];
  const groupDeletions = [];
  const groupInactive = [];
  
  // Find additions for this group
  changes.additions.forEach(change => {
    if (normalizeGroupName(change.member.gcgStatus.groupName) === normalizeGroupName(group.displayName)) {
      groupAdditions.push({
        personId: change.member.personId,
        firstName: change.member.firstName,
        lastName: change.member.lastName
      });
    }
  });
  
  // Find deletions for this group
  changes.removals.forEach(change => {
    if (normalizeGroupName(change.member.group) === normalizeGroupName(group.displayName)) {
      groupDeletions.push({
        personId: change.member.personId,
        firstName: change.member.firstName,
        lastName: change.member.lastName
      });
    }
  });
  
  // Find inactive members for this group
  const inactiveInGroup = allCurrentMembers.filter(member => 
    normalizeGroupName(member.group) === normalizeGroupName(group.displayName) && 
    isMarkedInactive(member)
  );
  
  inactiveInGroup.forEach(member => {
    groupInactive.push({
      personId: member.personId,
      firstName: member.firstName,
      lastName: member.lastName,
      inactiveReason: member.actionSteps || 'Marked inactive'
    });
  });
  
  return {
    additions: groupAdditions,
    deletions: groupDeletions,
    inactive: groupInactive
  };
}

/**
 * Build Section 3: Statistics (E2:G5) - shifted down by 1 row and moved data to column G
 * Shows overall statistics about membership with clean counts
 */
function buildStatsReport(sheet, exportData) {
  console.log('📝 Building statistics section...');
  
  const originalActiveCount = exportData.summary.originalActiveMembers;
  const activeInGCGCount = exportData.membersWithGCGStatus.filter(m => 
    m.gcgStatus.inGroup && !m.isSynthetic
  ).length;
  const activeNotInGCGCount = originalActiveCount - activeInGCGCount;
  const gcgMembersNotInActive = exportData.summary.syntheticMembers || 0;
  
  // Clean statistics - labels in E, data in G (F is buffer)
  sheet.getRange('E2').setValue('Active Members:');
  sheet.getRange('G2').setValue(originalActiveCount);
  
  sheet.getRange('E3').setValue('Active Members in GCGs:');
  sheet.getRange('G3').setValue(activeInGCGCount);
  
  sheet.getRange('E4').setValue('Active Members not in GCGs:');
  sheet.getRange('G4').setValue(activeNotInGCGCount);
  
  sheet.getRange('E5').setValue('GCG Members not in Active list:');
  sheet.getRange('G5').setValue(gcgMembersNotInActive);
  
  // Make labels bold
  sheet.getRange('E2:E5').setFontWeight('bold');
  
  // Highlight the data inconsistency if present
  if (gcgMembersNotInActive > 0) {
    sheet.getRange('E5:G5').setBackground('#fff2cc'); // Light yellow warning
  }
  
  console.log('✅ Built statistics section with clean counts');
}

/**
 * Build Section 4: Not in GCG Updates (E10+)
 * Shows proposed updates to the "Not in a GCG" tab
 */
function buildNotInGCGReport(sheet, exportData) {
  console.log('📝 Building Not in GCG section...');
  
  // Header
  sheet.getRange('E10').setValue('Proposed Updates to Active Members Not in a GCG');
  sheet.getRange('E10').setFontWeight('bold');
  sheet.getRange('E10').setBackground('#f3e5f5'); // Light purple
  
  // Column headers
  sheet.getRange('E11').setValue('Person ID');
  sheet.getRange('F11').setValue('First Name');
  sheet.getRange('G11').setValue('Last Name');
  sheet.getRange('H11').setValue('Family ID');
  sheet.getRange('I11').setValue('Family Role');
  sheet.getRange('E11:I11').setFontWeight('bold');
  
  let currentRow = 12;
  
  // Calculate Not in GCG changes
  const notInGCGChanges = calculateNotInGCGChanges(exportData);
  
  // Additions section
  sheet.getRange(currentRow, 5).setValue('Additions');
  sheet.getRange(currentRow, 5).setFontWeight('bold');
  currentRow++;
  
  if (notInGCGChanges.additions.length > 0) {
    notInGCGChanges.additions.forEach(person => {
      sheet.getRange(currentRow, 5).setValue(person.personId);
      sheet.getRange(currentRow, 6).setValue(person.firstName);
      sheet.getRange(currentRow, 7).setValue(person.lastName);
      sheet.getRange(currentRow, 8).setValue(person.familyId || 'null');
      sheet.getRange(currentRow, 9).setValue(person.familyRole || 'null');
      currentRow++;
    });
  } else {
    sheet.getRange(currentRow, 5).setValue('None');
    currentRow++;
  }
  
  // Deletions section
  sheet.getRange(currentRow, 5).setValue('Deletions');
  sheet.getRange(currentRow, 5).setFontWeight('bold');
  currentRow++;
  
  if (notInGCGChanges.deletions.length > 0) {
    notInGCGChanges.deletions.forEach(person => {
      sheet.getRange(currentRow, 5).setValue(person.personId);
      sheet.getRange(currentRow, 6).setValue(person.firstName);
      sheet.getRange(currentRow, 7).setValue(person.lastName);
      sheet.getRange(currentRow, 8).setValue(person.familyId || 'null');
      sheet.getRange(currentRow, 9).setValue(person.familyRole || 'null');
      currentRow++;
    });
  } else {
    sheet.getRange(currentRow, 5).setValue('None');
    currentRow++;
  }
  
  console.log(`✅ Built Not in GCG section with ${notInGCGChanges.additions.length} additions and ${notInGCGChanges.deletions.length} deletions`);
}

/**
 * Helper function to find group-specific changes
 * @param {Object} group - Group object from export data
 * @param {Object} changes - Overall changes object
 * @param {Object} exportData - Full export data
 * @returns {Object} Group-specific additions and deletions
 */
function findGroupSpecificChanges(group, changes, exportData) {
  const groupAdditions = [];
  const groupDeletions = [];
  
  // Find additions for this group
  changes.additions.forEach(change => {
    if (normalizeGroupName(change.member.gcgStatus.groupName) === normalizeGroupName(group.displayName)) {
      groupAdditions.push({
        personId: change.member.personId,
        firstName: change.member.firstName,
        lastName: change.member.lastName
      });
    }
  });
  
  // Find deletions for this group
  changes.removals.forEach(change => {
    if (normalizeGroupName(change.member.group) === normalizeGroupName(group.displayName)) {
      groupDeletions.push({
        personId: change.member.personId,
        firstName: change.member.firstName,
        lastName: change.member.lastName
      });
    }
  });
  
  return {
    additions: groupAdditions,
    deletions: groupDeletions
  };
}

/**
 * Helper function to calculate Not in GCG changes
 * @param {Object} exportData - Full export data
 * @returns {Object} People to add/remove from "Not in GCG" tab
 */
function calculateNotInGCGChanges(exportData) {
  // Get people not in GCGs from export data
  const notInGCGFromExport = exportData.membersWithGCGStatus.filter(m => 
    !m.gcgStatus.inGroup && m.isActiveMember && !m.isSynthetic
  );
  
  // For now, return simplified structure
  // In full implementation, this would compare with current "Not in GCG" tab
  const additions = notInGCGFromExport.slice(0, 10).map(person => ({
    personId: person.personId,
    firstName: person.firstName,
    lastName: person.lastName,
    familyId: 'TBD', // Would come from family data
    familyRole: 'TBD' // Would come from family data
  }));
  
  return {
    additions: additions,
    deletions: [] // Would be populated by comparing with current sheet
  };
}

/**
 * Build Section 5: Data Inconsistencies (E26+) - shifted down by 1 row with buffer and improved formatting
 * Shows GCG members who aren't in the Active Members list
 */
function buildDataInconsistenciesReport(sheet, exportData) {
  console.log('📝 Building data inconsistencies section...');
  
  if (!exportData.missingFromActive || exportData.missingFromActive.length === 0) {
    console.log('✅ No data inconsistencies found');
    return;
  }
  
  let currentRow = 26; // Added 1-row buffer from Section 4
  
  // Section header with improved highlighting
  sheet.getRange(`E${currentRow}:I${currentRow}`).merge();
  sheet.getRange(`E${currentRow}`).setValue('Data Inconsistencies - Action Required');
  sheet.getRange(`E${currentRow}`).setFontWeight('bold');
  sheet.getRange(`E${currentRow}`).setBackground('#ffcdd2'); // Light red
  currentRow++;
  
  // Subsection header
  sheet.getRange(`E${currentRow}`).setValue('GCG Members Not in Active Members List');
  sheet.getRange(`E${currentRow}`).setFontWeight('bold');
  currentRow++;
  
  // Column headers
  sheet.getRange(`E${currentRow}`).setValue('Person ID');
  sheet.getRange(`F${currentRow}`).setValue('First Name');
  sheet.getRange(`G${currentRow}`).setValue('Last Name');
  sheet.getRange(`H${currentRow}`).setValue('Current GCG');
  sheet.getRange(`I${currentRow}`).setValue('Action Needed');
  sheet.getRange(`E${currentRow}:I${currentRow}`).setFontWeight('bold');
  currentRow++;
  
  // List each inconsistent member
  exportData.missingFromActive.forEach(person => {
    sheet.getRange(`E${currentRow}`).setValue(person.personId);
    sheet.getRange(`F${currentRow}`).setValue(person.firstName);
    sheet.getRange(`G${currentRow}`).setValue(person.lastName);
    sheet.getRange(`H${currentRow}`).setValue(person.groupName);
    sheet.getRange(`I${currentRow}`).setValue('Add to Active Members OR Remove from GCG');
    currentRow++;
  });
  
  // Add instructions with improved step 4
  currentRow++;
  sheet.getRange(`E${currentRow}`).setValue('Instructions:');
  sheet.getRange(`E${currentRow}`).setFontWeight('bold');
  currentRow++;
  
  sheet.getRange(`E${currentRow}`).setValue('1. Review each person listed above');
  currentRow++;
  sheet.getRange(`E${currentRow}`).setValue('2. If they should be active: Add them to the "Members - Active" tag in Breeze');
  currentRow++;
  sheet.getRange(`E${currentRow}`).setValue('3. If they are inactive: Remove them from their GCG tag in Breeze');
  currentRow++;
  sheet.getRange(`E${currentRow}`).setValue('4. If they are members-in-process: Verify their membership status in Breeze and update as needed');
  currentRow++;
  sheet.getRange(`E${currentRow}`).setValue('5. Re-export and run this report again to verify fixes');
  
  console.log(`✅ Built data inconsistencies section with ${exportData.missingFromActive.length} issues`);
}

/**
 * Format the preview sheet for readability with improvements
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Preview sheet to format
 */
function formatPreviewSheet(sheet) {
  console.log('🎨 Formatting preview sheet...');
  
  try {
    // Auto-resize columns
    sheet.autoResizeColumns(1, 10);
    
    // Format main headers (shifted down by 1 row and updated for new column)
    sheet.getRange('A2:D2').setBackground('#4CAF50').setFontColor('white');
    sheet.getRange('E2:G5').setBorder(true, true, true, true, true, true);
    sheet.getRange('E11:I12').setBackground('#2196F3').setFontColor('white');
    
    // Set column widths for better readability (updated for new column D)
    sheet.setColumnWidth(1, 200); // Group names
    sheet.setColumnWidth(2, 140); // Breeze export count
    sheet.setColumnWidth(3, 160); // GCG members tab count (active)
    sheet.setColumnWidth(4, 120); // Inactive count
    sheet.setColumnWidth(5, 120); // Person ID
    sheet.setColumnWidth(6, 100); // First name
    sheet.setColumnWidth(7, 100); // Last name
    sheet.setColumnWidth(8, 180); // Group name / Family ID / Reason
    sheet.setColumnWidth(9, 200); // Action needed / Family Role
    
    // Don't freeze rows per feedback
    // sheet.setFrozenRows(1); - REMOVED
    
    console.log('✅ Preview sheet formatted successfully');
    
  } catch (error) {
    console.warn('⚠️ Some formatting may not have applied:', error.message);
  }
}
