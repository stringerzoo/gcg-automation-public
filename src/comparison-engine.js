/**
 * Comparison Engine for GCG Automation (Production Version)
 * Core logic for comparing export data with current Google Sheet
 */

/**
 * Get GCG Members with Person ID from the sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet object
 * @returns {Array} Current GCG members with Person IDs
 */
function getGCGMembersWithPersonId(ss) {
  console.log('📋 Reading GCG Members with Person ID...');
  
  const sheet = ss.getSheetByName('GCG Members');
  if (!sheet) {
    throw new Error('GCG Members sheet not found');
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }
  
  const headerRowIndex = findHeaderRow(data);
  const headers = data[headerRowIndex >= 0 ? headerRowIndex : 1];
  
  // Column mapping with Person ID
  const cols = {
    personId: findColumnIndex(headers, 'Person ID'),
    firstName: findColumnIndex(headers, 'First'),
    lastName: findColumnIndex(headers, 'Last'),
    group: findColumnIndex(headers, 'Group'),
    deacon: findColumnIndex(headers, 'Deacon'),
    pastor: findColumnIndex(headers, 'Pastor'),
    team: findColumnIndex(headers, 'Team'),
    actionSteps: findColumnIndex(headers, 'Action Steps'),
    assignedTo: findColumnIndex(headers, 'Assigned to')
  };
  
  if (cols.personId === -1) {
    throw new Error('Person ID column not found. Please ensure column A has Person ID header.');
  }
  
  const members = [];
  const dataStartRow = (headerRowIndex >= 0 ? headerRowIndex : 1) + 1;
  
  for (let i = dataStartRow; i < data.length; i++) {
    const row = data[i];
    
    if (!row[cols.firstName] || !row[cols.lastName]) continue;
    
    const member = {
      personId: row[cols.personId] ? String(row[cols.personId]) : null,
      firstName: row[cols.firstName],
      lastName: row[cols.lastName],
      fullName: `${row[cols.firstName]} ${row[cols.lastName]}`.trim(),
      group: row[cols.group] || '',
      deacon: row[cols.deacon] || '',
      pastor: row[cols.pastor] || '',
      team: row[cols.team] || '',
      actionSteps: row[cols.actionSteps] || '',
      assignedTo: row[cols.assignedTo] || '',
      rowIndex: i + 1
    };
    
    // Skip 'x' entries and include all others
    if (member.group && member.group.toString().toLowerCase() !== 'x') {
      members.push(member);
    }
  }
  
  console.log(`✅ Found ${members.length} GCG members with Person IDs`);
  
  // Report ID coverage
  const withIds = members.filter(m => m.personId).length;
  console.log(`📊 ${withIds}/${members.length} members have Person IDs (${(withIds/members.length*100).toFixed(1)}%)`);
  
  return members;
}

/**
 * Check if a member should be considered inactive based on Action Steps column
 * @param {Object} member - Member object from current sheet
 * @returns {boolean} True if member is marked as inactive
 */
function isMarkedInactive(member) {
  if (!member.actionSteps) return false;
  
  const actionSteps = member.actionSteps.toString().toLowerCase();
  
  // Check for various inactive indicators
  const inactiveIndicators = [
    'inactive',
    'moved away',
    'no longer active',
    'left church',
    'transferred'
  ];
  
  return inactiveIndicators.some(indicator => actionSteps.includes(indicator));
}

/**
 * Normalize group names for accurate comparison
 * Handles format differences like "Gene Cone" vs "Gene Cone & Scott Stringer"
 * @param {string} groupName - Group name to normalize
 * @returns {string} Normalized group name (primary leader only)
 */
function normalizeGroupName(groupName) {
  if (!groupName) return '';
  
  const normalized = groupName.toString().trim();
  
  // Extract primary leader name (everything before &)
  const parts = normalized.split('&');
  const primaryLeader = parts[0].trim();
  
  return primaryLeader;
}

/**
 * Enhanced data parsing that includes people from GCG tags even if missing from Active Members
 * This provides a complete picture while flagging data inconsistencies
 * @returns {Object} Enhanced GCG data with synthetic members
 */
function parseRealGCGDataWithGCGMembers() {
  console.log('🎯 Parsing REAL GCG data (including GCG-only members)...');
  
  try {
    // Get the standard data using smart file detection
    const standardData = parseRealGCGDataSmart();
    
    // Find people in GCG assignments but missing from Active Members
    const gcgPersonIds = new Set(Object.keys(standardData.assignments));
    const activePersonIds = new Set(standardData.activeMembers.map(m => m.personId));
    
    const missingFromActive = [];
    gcgPersonIds.forEach(personId => {
      if (!activePersonIds.has(personId)) {
        const assignment = standardData.assignments[personId];
        const group = standardData.groups.find(g => g.displayName === assignment.groupName);
        const member = group ? group.members.find(m => m.personId === personId) : null;
        
        if (member) {
          missingFromActive.push({
            personId: personId,
            firstName: member.firstName,
            lastName: member.lastName,
            fullName: `${member.firstName} ${member.lastName}`,
            groupName: assignment.groupName,
            sheetName: assignment.sheetName
          });
        }
      }
    });
    
    // Create synthetic active member records for people in GCGs but missing from Active Members
    const syntheticActiveMembers = missingFromActive.map(person => ({
      personId: person.personId,
      firstName: person.firstName,
      lastName: person.lastName,
      fullName: person.fullName,
      address: { street: '', city: '', state: '', zip: '' },
      isActiveMember: false, // Flag as potentially inactive
      isSynthetic: true // Flag as created from GCG data
    }));
    
    // Combine active members with synthetic ones
    const enhancedActiveMembers = [
      ...standardData.activeMembers,
      ...syntheticActiveMembers
    ];
    
    // Recreate membersWithGCGStatus with the enhanced list
    const enhancedMembersWithGCGStatus = enhancedActiveMembers.map(member => {
      const gcgAssignment = standardData.assignments[member.personId];
      return {
        ...member,
        gcgStatus: {
          inGroup: !!gcgAssignment,
          groupName: gcgAssignment?.groupName || null,
          leader: gcgAssignment?.leader || null,
          coLeader: gcgAssignment?.coLeader || null
        }
      };
    });
    
    const inGCG = enhancedMembersWithGCGStatus.filter(m => m.gcgStatus.inGroup).length;
    const notInGCG = enhancedMembersWithGCGStatus.filter(m => !m.gcgStatus.inGroup).length;
    const participationRate = (inGCG / enhancedActiveMembers.length * 100).toFixed(1);
    
    console.log('\n📊 ENHANCED RESULTS:');
    console.log(`👥 Active Members: ${standardData.activeMembers.length} (original) + ${syntheticActiveMembers.length} (from GCG tags) = ${enhancedActiveMembers.length}`);
    console.log(`🏘️ GCG Groups: ${standardData.groups.length}`);
    console.log(`✅ In GCGs: ${inGCG} (${participationRate}%)`);
    console.log(`❌ Not in GCGs: ${notInGCG}`);
    
    return {
      activeMembers: enhancedActiveMembers,
      membersWithGCGStatus: enhancedMembersWithGCGStatus,
      groups: standardData.groups,
      assignments: standardData.assignments,
      missingFromActive: missingFromActive,
      summary: {
        totalActiveMembers: enhancedActiveMembers.length,
        originalActiveMembers: standardData.activeMembers.length,
        syntheticMembers: syntheticActiveMembers.length,
        totalGroups: standardData.groups.length,
        inGCG: inGCG,
        notInGCG: notInGCG,
        participationRate: participationRate
      }
    };
    
  } catch (error) {
    console.error('❌ Enhanced parsing failed:', error.message);
    throw error;
  }
}

/**
 * Enhanced comparison with inactive filtering that handles the complete workflow
 * @param {Object} exportData - Enhanced export data with synthetic members
 * @returns {Object} Changes needed with proper filtering
 */
function fixedCompareWithInactiveFiltering(exportData) {
  console.log('🔍 Enhanced comparison with inactive filtering...');
  
  try {
    const config = getConfig();
    const ss = SpreadsheetApp.openById(config.SHEET_ID);
    
    // Get current data with Person IDs
    const allCurrentMembers = getGCGMembersWithPersonId(ss);
    
    // Filter out members marked as inactive
    const activeCurrentMembers = allCurrentMembers.filter(member => {
      const isInactive = isMarkedInactive(member);
      if (isInactive) {
        console.log(`⚠️ Filtering out ${member.fullName} - marked as inactive: "${member.actionSteps}"`);
      }
      return !isInactive;
    });
    
    console.log(`📊 Filtered: ${allCurrentMembers.length} → ${activeCurrentMembers.length} (removed ${allCurrentMembers.length - activeCurrentMembers.length} inactive)`);
    
    // Get export data - filter out people marked as inactive in current sheet
    const exportMembers = exportData.membersWithGCGStatus.filter(m => {
      if (!m.gcgStatus.inGroup) return false;
      
      // Check if this person is marked inactive in our current sheet
      const currentMember = allCurrentMembers.find(cm => cm.personId === m.personId);
      if (currentMember) {
        const isInactive = isMarkedInactive(currentMember);
        if (isInactive) {
          console.log(`⚠️ Filtering export member ${m.fullName} - marked inactive in current sheet`);
          return false;
        }
      }
      
      return true;
    });
    
    console.log(`📊 Comparing: ${activeCurrentMembers.length} active current vs ${exportMembers.length} active export`);
    
    // Create lookup maps
    const currentByPersonId = new Map();
    const currentByName = new Map();
    
    activeCurrentMembers.forEach(member => {
      if (member.personId) {
        currentByPersonId.set(member.personId, member);
      } else {
        currentByName.set(member.fullName.toLowerCase(), member);
      }
    });
    
    const exportByPersonId = new Map();
    exportMembers.forEach(member => {
      exportByPersonId.set(member.personId, member);
    });
    
    const changes = {
      additions: [],
      updates: [],
      removals: [],
      inactiveFiltered: allCurrentMembers.length - activeCurrentMembers.length,
      exportMissingButActive: []
    };
    
    // Find additions and updates
    exportMembers.forEach(exportMember => {
      const currentMember = currentByPersonId.get(exportMember.personId);
      
      if (!currentMember) {
        const nameMatch = currentByName.get(exportMember.fullName.toLowerCase());
        
        if (nameMatch) {
          changes.updates.push({
            type: 'missingPersonId',
            member: nameMatch,
            exportMember: exportMember,
            reason: 'Person exists by name but missing Person ID'
          });
        } else {
          changes.additions.push({
            type: 'addition',
            member: exportMember,
            reason: 'New member in GCG exports'
          });
        }
      } else {
        // Check for group updates with normalization
        const currentGroupNormalized = normalizeGroupName(currentMember.group);
        const exportGroupNormalized = normalizeGroupName(exportMember.gcgStatus.groupName);
        
        if (currentGroupNormalized !== exportGroupNormalized) {
          changes.updates.push({
            type: 'update',
            currentMember: currentMember,
            exportMember: exportMember,
            updates: [{
              field: 'group',
              oldValue: currentMember.group,
              newValue: exportMember.gcgStatus.groupName,
              normalizedOld: currentGroupNormalized,
              normalizedNew: exportGroupNormalized
            }]
          });
        }
      }
    });
    
    // Find removals
    activeCurrentMembers.forEach(currentMember => {
      if (currentMember.personId && !exportByPersonId.has(currentMember.personId)) {
        const stillActiveInExport = exportData.activeMembers.find(m => m.personId === currentMember.personId);
        
        if (stillActiveInExport) {
          changes.exportMissingButActive.push({
            type: 'exportDataIncomplete',
            member: currentMember,
            reason: 'Person is active but missing from GCG tags in export'
          });
        } else {
          changes.removals.push({
            type: 'removal',
            member: currentMember,
            reason: 'No longer in active members'
          });
        }
      }
    });
    
    const totalChanges = changes.additions.length + changes.updates.length + changes.removals.length;
    
    console.log('\n📊 COMPARISON RESULTS:');
    console.log(`🔄 Total real changes: ${totalChanges}`);
    console.log(`➕ Additions: ${changes.additions.length}`);
    console.log(`🔄 Updates: ${changes.updates.length}`);
    console.log(`➖ Removals: ${changes.removals.length}`);
    console.log(`⚠️ Export missing but active: ${changes.exportMissingButActive.length}`);
    console.log(`🚫 Filtered inactive: ${changes.inactiveFiltered}`);
    
    return changes;
    
  } catch (error) {
    console.error('❌ Enhanced comparison failed:', error.message);
    throw error;
  }
}

/**
 * Find the actual header row in a sheet (skips instruction rows)
 * @param {Array} data - Sheet data
 * @returns {number} Row index of actual headers, or -1 if not found
 */
function findHeaderRow(data) {
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    
    // Look for common header patterns
    const hasFirst = row.some(cell => cell && cell.toString().toLowerCase().includes('first'));
    const hasLast = row.some(cell => cell && cell.toString().toLowerCase().includes('last'));
    
    if (hasFirst && hasLast) {
      return i;
    }
  }
  
  return -1;
}

/**
 * Find column index by header name (case-insensitive)
 * @param {Array} headers - Array of header names
 * @param {string} searchName - Column name to find
 * @returns {number} Column index or -1 if not found
 */
function findColumnIndex(headers, searchName) {
  const searchLower = searchName.toLowerCase();
  
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && headers[i].toString().toLowerCase().includes(searchLower)) {
      return i;
    }
  }
  
  return -1;
}
