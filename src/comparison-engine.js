/**
 * Enhanced Comparison Engine for GCG Automation - Family Logic Implementation
 * Core logic for comparing export data with current Google Sheet including family grouping
 */

/**
 * FAMILY ROLE PRIORITY CONSTANTS
 * Based on Breeze family role hierarchy
 */
const FAMILY_ROLE_PRIORITY = {
  'Head of Household': 1,
  'Spouse': 2,
  'Adult': 3,
  'Child': 4,
  'Unassigned': 5
};

/**
 * Get priority score for family role (lower number = higher priority)
 * @param {string} familyRole - Role like 'Head of Household', 'Spouse', etc.
 * @returns {number} Priority score
 */
function getFamilyRolePriority(familyRole) {
  if (!familyRole) return 99; // No role = lowest priority
  return FAMILY_ROLE_PRIORITY[familyRole] || 99;
}

/**
 * Enhanced function to calculate Not in GCG changes with proper family logic
 * @param {Object} exportData - Full export data with active members
 * @returns {Object} Additions and deletions for "Not in GCG" tab
 */
function calculateNotInGCGChangesWithFamilyLogic(exportData) {
  console.log('👨‍👩‍👧‍👦 Calculating Not in GCG changes with family logic...');
  
  try {
    // Get current "Not in GCG" tab data (with improved family-grouped reading)
    const config = getConfig();
    const ss = SpreadsheetApp.openById(config.SHEET_ID);
    const currentNotInGCG = getCurrentNotInGCGMembers(ss);
    
    // Find people who SHOULD be in "Not in GCG" (with family representatives)
    const shouldBeInNotInGCG = calculateFamilyRepresentatives(exportData);
    
    // Create lookup maps using Person ID
    const currentlyListed = new Set(currentNotInGCG.map(p => p.personId).filter(id => id));
    const shouldBeListed = new Set(shouldBeInNotInGCG.map(p => p.personId));
    
    console.log(`📊 Currently listed: ${currentlyListed.size} people`);
    console.log(`📊 Should be listed: ${shouldBeListed.size} people`);
    
    // Calculate additions (should be listed but aren't)
    const additions = shouldBeInNotInGCG.filter(person => 
      !currentlyListed.has(person.personId)
    );
    
    // Calculate deletions (currently listed but shouldn't be)
    // ENHANCED: Look up proper Family ID and Family Role from exportData
    const deletions = [];
    currentNotInGCG.forEach(currentPerson => {
      if (!currentPerson.personId) return; // Skip if no Person ID
      
      const shouldBeIncluded = shouldBeListed.has(currentPerson.personId);
      
      if (!shouldBeIncluded) {
        // Check WHY they shouldn't be included
        const personInExport = exportData.membersWithGCGStatus.find(m => m.personId === currentPerson.personId);
        
        let reason = 'Unknown';
        if (!personInExport) {
          reason = 'No longer in active members';
        } else if (personInExport.gcgStatus.inGroup) {
          reason = `Now in GCG: ${personInExport.gcgStatus.groupName}`;
        } else if (personInExport.isSynthetic) {
          reason = 'Data inconsistency - not in active members export';
        }
        
        // FIXED: Use proper Family ID and Family Role from exportData lookup
        const properFamilyId = personInExport ? personInExport.familyId : null;
        const properFamilyRole = personInExport ? personInExport.familyRole : null;
        
        deletions.push({
          personId: currentPerson.personId,
          firstName: currentPerson.firstName,
          lastName: currentPerson.lastName,
          // FIXED: Use looked-up values instead of family-grouped data
          familyId: properFamilyId,
          familyRole: properFamilyRole,
          reason: reason,
          rowIndex: currentPerson.rowIndex
        });
        
        console.log(`➖ Deletion: ${currentPerson.firstName} ${currentPerson.lastName} - ${reason} (Family ID: ${properFamilyId || 'none'})`);
      }
    });
    
    console.log(`✅ Family logic results: ${additions.length} additions, ${deletions.length} deletions`);
    
    return {
      additions: additions,
      deletions: deletions,
      familyGroupsProcessed: shouldBeInNotInGCG.length
    };
    
  } catch (error) {
    console.error('❌ Family logic calculation failed:', error.message);
    throw error;
  }
}

/**
 * Get current members from "Not in GCG" tab with Person IDs
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet object
 * @returns {Array} Current "Not in GCG" members
 * updated to handle family-grouped format
 */
function getCurrentNotInGCGMembers(ss) {
  console.log('📋 Reading current "Not in GCG" tab (family-grouped format)...');
  
  const sheet = ss.getSheetByName('Not in a GCG');
  if (!sheet) {
    console.warn('⚠️ "Not in a GCG" sheet not found');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 3) { // Headers in row 3, data starts row 4
    return [];
  }
  
  // Find headers (should be in row 3)
  const headers = data[2]; // Row 3 (index 2)
  const personIdCol = findColumnIndex(headers, 'Person ID');
  const familyNameCol = findColumnIndex(headers, 'Family Name'); // Column A - modified family names
  const firstNameCol = findColumnIndex(headers, 'First Name'); // Column B - comma-separated first names
  const lastNameCol = findColumnIndex(headers, 'Last Name'); // Column C - family last name
  const familyIdCol = findColumnIndex(headers, 'Family');
  const familyRoleCol = findColumnIndex(headers, 'Family Role');
  
  console.log(`🔍 Column mapping: PersonID=${personIdCol}, FamilyName=${familyNameCol}, FirstName=${firstNameCol}, LastName=${lastNameCol}`);
  
  if (personIdCol === -1) {
    console.warn('⚠️ Person ID column not found in "Not in GCG" tab');
    return [];
  }
  
  const members = [];
  for (let i = 3; i < data.length; i++) { // Data starts row 4 (index 3)
    const row = data[i];
    
    if (row[personIdCol]) {
      // For family-grouped data, we use the Family Name (Column A) as the primary identifier
      // This has been manually edited to show the specific person's name (like "Rachel King")
      const familyName = row[familyNameCol] || '';
      const commaSeparatedFirstNames = row[firstNameCol] || '';
      const lastName = row[lastNameCol] || '';
      
      // Extract the actual person's name from Family Name column
      // Family Name should be like "Rachel King" (manually edited) or "John & Sally Smith" (default)
      let actualFirstName = '';
      if (familyName.includes('&')) {
        // Default family format - use first name from comma-separated list
        actualFirstName = commaSeparatedFirstNames.split(',')[0].trim();
      } else {
        // Manually edited to show specific person - extract first name
        const nameParts = familyName.trim().split(' ');
        actualFirstName = nameParts[0];
      }
      
      members.push({
        personId: String(row[personIdCol]),
        firstName: actualFirstName,
        lastName: lastName,
        familyName: familyName, // Keep the full family name for reference
        familyId: row[familyIdCol] || null,
        familyRole: row[familyRoleCol] || null,
        rowIndex: i + 1
      });
      
      console.log(`📋 Current member: ${actualFirstName} ${lastName} (Family Name: "${familyName}", ID: ${row[personIdCol]})`);
    }
  }
  
  console.log(`✅ Found ${members.length} current "Not in GCG" members`);
  return members;
}

/**
 * Calculate family representatives for "Not in GCG" using priority logic
 * @param {Object} exportData - Full export data
 * @returns {Array} Family representatives who should be in "Not in GCG"
 * Enhanced Calculate family representatives with Elders and Tuesday School exclusions
 */
function calculateFamilyRepresentatives(exportData) {
  console.log('👨‍👩‍👧‍👦 Calculating family representatives with exclusions...');
  
  // Get people who should be excluded from "Not in GCG" analysis
  const excludedPersonIds = getExcludedPersonIds(exportData);
  console.log(`🚫 Found ${excludedPersonIds.size} people to exclude (Elders + Tuesday School)`);
  
  // ALSO get people marked as inactive in the current GCG Members sheet
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.SHEET_ID);
  const allCurrentMembers = getGCGMembersWithPersonId(ss);
  const inactivePersonIds = new Set(
    allCurrentMembers
      .filter(member => isMarkedInactive(member))
      .map(member => member.personId)
      .filter(id => id) // Only include those with Person IDs
  );
  
  console.log(`⚠️ Found ${inactivePersonIds.size} people marked inactive in current sheet`);
  
  // Combine both exclusion sets
  const allExcludedIds = new Set([...excludedPersonIds, ...inactivePersonIds]);
  console.log(`🚫 Total excluded: ${allExcludedIds.size} people`);
  
  // Get active members not in GCGs, excluding Elders, Tuesday School, AND inactive members
  const notInGCGCandidates = exportData.membersWithGCGStatus.filter(m => 
    !m.gcgStatus.inGroup && 
    m.isActiveMember && 
    !m.isSynthetic &&
    !allExcludedIds.has(m.personId) // ENHANCED: Exclude both admin tags AND inactive members
  );
  
  console.log(`📋 Found ${notInGCGCandidates.length} active members not in GCGs (after all exclusions)`);
  
  // Group by Family ID
  const familyGroups = new Map();
  const individualsWithoutFamily = [];
  
  notInGCGCandidates.forEach(person => {
    const familyId = person.familyId || person.family;
    const familyRole = person.familyRole || person.family_role;
    
    if (familyId && familyId !== 'null' && familyId !== '') {
      if (!familyGroups.has(familyId)) {
        familyGroups.set(familyId, []);
      }
      familyGroups.get(familyId).push({
        ...person,
        familyId: familyId,
        familyRole: familyRole
      });
    } else {
      individualsWithoutFamily.push({
        ...person,
        familyId: null,
        familyRole: null
      });
    }
  });
  
  console.log(`👨‍👩‍👧‍👦 Processing ${familyGroups.size} families and ${individualsWithoutFamily.length} individuals`);
  
  const representatives = [];
  
  // Process families - pick one representative per family
  familyGroups.forEach((familyMembers, familyId) => {
    const representative = selectFamilyRepresentative(familyMembers);
    if (representative) {
      representatives.push(representative);
    }
  });
  
  // Add individuals without families
  representatives.push(...individualsWithoutFamily);
  
  console.log(`✅ Selected ${representatives.length} family representatives (after all exclusions)`);
  
  return representatives;
}
/**
 * NEW FUNCTION: Get Person IDs of people who should be excluded from "Not in GCG" analysis
 * @param {Object} exportData - Full export data with GCG assignments
 * @returns {Set} Set of Person IDs to exclude
 */
function getExcludedPersonIds(exportData) {
  console.log('🚫 Finding excluded Person IDs from raw tags data...');
  
  const excludedIds = new Set();
  
  try {
    // Get the raw tags file to read Elders and Tuesday School directly
    const tagsFile = findLatestFile('TAGS_EXPORT');
    const spreadsheet = SpreadsheetApp.openById(tagsFile.getId());
    const allSheets = spreadsheet.getSheets();
    
    console.log(`📊 Scanning ${allSheets.length} total tabs for exclusions...`);
    
    // FIXED: More specific exclusion tab names (exact matches only)
    const exclusionTabs = [
      { name: 'Elders', exact: true },
      { name: 'Tuesday School W Pastor Doug', exact: true },
      { name: 'Elders And Elders Wives', exact: true },
      { name: 'Elders Children Under 18', exact: true }
      // NOTE: Removed "Elders Children" (the broader category) but kept "Elders Children Under 18"
      // This should exclude children under 18 but allow adult children of elders to remain in the list
    ];
    
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      
      // Check if this is an exclusion tab using exact matching
      const matchingTab = exclusionTabs.find(exclusionTab => {
        if (exclusionTab.exact) {
          return sheetName.toLowerCase() === exclusionTab.name.toLowerCase();
        } else {
          return sheetName.toLowerCase().includes(exclusionTab.name.toLowerCase());
        }
      });
      
      if (matchingTab) {
        console.log(`🚫 Processing exclusion tab: "${sheetName}" (exact match for "${matchingTab.name}")`);
        
        try {
          const data = sheet.getDataRange().getValues();
          if (data.length <= 1) return; // Skip if no data
          
          // Find Person ID column (should be first column typically)
          const headers = data[0];
          const personIdCol = findColumnIndex(headers, 'Person ID');
          
          if (personIdCol === -1) {
            console.warn(`⚠️ No Person ID column found in ${sheetName}`);
            return;
          }
          
          // Add all Person IDs from this tab to exclusions
          for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const personId = row[personIdCol];
            
            if (personId) {
              excludedIds.add(String(personId));
              console.log(`🚫 Excluding: ${personId} from ${sheetName}`);
            }
          }
          
        } catch (error) {
          console.warn(`⚠️ Failed to process ${sheetName}: ${error.message}`);
        }
      }
    });
    
    console.log(`✅ Found ${excludedIds.size} people to exclude from administrative tabs`);
    
  } catch (error) {
    console.error('❌ Failed to read exclusion tabs:', error.message);
    console.warn('⚠️ Continuing without exclusions');
  }
  
  return excludedIds;
}
/**
 * Select the best family representative based on priority rules
 * @param {Array} familyMembers - Array of family members not in GCGs
 * @returns {Object} Selected family representative
 */
function selectFamilyRepresentative(familyMembers) {
  if (familyMembers.length === 0) return null;
  if (familyMembers.length === 1) return familyMembers[0];
  
  // Sort by family role priority (lower number = higher priority)
  const sortedMembers = familyMembers.sort((a, b) => {
    const priorityA = getFamilyRolePriority(a.familyRole);
    const priorityB = getFamilyRolePriority(b.familyRole);
    
    if (priorityA !== priorityB) {
      return priorityA - priorityB; // Lower priority number = higher actual priority
    }
    
    // If same priority, sort by age (older first) - use birthdate if available
    // For now, just use first name alphabetically as tie-breaker
    return (a.firstName || '').localeCompare(b.firstName || '');
  });
  
  const selected = sortedMembers[0];
  
  console.log(`👨‍👩‍👧‍👦 Family ${selected.familyId}: Selected ${selected.firstName} ${selected.lastName} (${selected.familyRole || 'No Role'}) from ${familyMembers.length} members`);
  
  return selected;
}

/**
 * Enhanced data parsing that includes family information from Active Members export
 * @returns {Object} Enhanced GCG data with family information
 */
function parseRealGCGDataWithFamilyInfo() {
  console.log('🎯 Parsing REAL GCG data with family information...');
  
  try {
    // Get the standard data using smart file detection
    const standardData = parseRealGCGDataWithGCGMembers();
    
    // Enhance active members with family information from the export
    const enhancedActiveMembers = enhanceActiveMembersWithFamilyData(standardData.activeMembers);
    
    // Recreate membersWithGCGStatus with enhanced family data
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
    
    console.log('\n📊 ENHANCED RESULTS WITH FAMILY DATA:');
    console.log(`👥 Active Members: ${enhancedActiveMembers.length}`);
    console.log(`🏘️ GCG Groups: ${standardData.groups.length}`);
    console.log(`✅ In GCGs: ${inGCG} (${participationRate}%)`);
    console.log(`❌ Not in GCGs: ${notInGCG}`);
    
    return {
      ...standardData,
      activeMembers: enhancedActiveMembers,
      membersWithGCGStatus: enhancedMembersWithGCGStatus,
      summary: {
        ...standardData.summary,
        totalActiveMembers: enhancedActiveMembers.length,
        inGCG: inGCG,
        notInGCG: notInGCG,
        participationRate: participationRate
      }
    };
    
  } catch (error) {
    console.error('❌ Enhanced family data parsing failed:', error.message);
    throw error;
  }
}

/**
 * Enhance active members with family data from the original export
 * @param {Array} activeMembers - Original active members array
 * @returns {Array} Enhanced active members with family information
 */
function enhanceActiveMembersWithFamilyData(activeMembers) {
  console.log('👨‍👩‍👧‍👦 Enhancing active members with family data...');
  
  try {
    // Re-read the Active Members file to get family information
    const activeMembersFile = findLatestFile('ACTIVE_MEMBERS');
    const spreadsheet = SpreadsheetApp.openById(activeMembersFile.getId());
    const dataSheet = spreadsheet.getSheets()[0];
    const data = dataSheet.getDataRange().getValues();
    
    if (data.length === 0) {
      console.warn('⚠️ No data in Active Members sheet for family enhancement');
      return activeMembers;
    }
    
    // Find family-related columns
    const headers = data[0];
    const columnMap = {
      personId: findColumnIndex(headers, 'Breeze ID'),
      familyId: findColumnIndex(headers, 'Family'),
      familyRole: findColumnIndex(headers, 'Family Role'),
      membershipStartDate: findColumnIndex(headers, 'Membership Start Date'),
      yearsSinceMembership: findColumnIndex(headers, 'Years Since Membership Start Date')
    };
    
    // Create lookup map for family data
    const familyDataMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const personId = String(row[columnMap.personId]);
      
      if (personId) {
        familyDataMap.set(personId, {
          familyId: row[columnMap.familyId] || null,
          familyRole: row[columnMap.familyRole] || null,
          membershipStartDate: row[columnMap.membershipStartDate] || null,
          yearsSinceMembership: row[columnMap.yearsSinceMembership] || null
        });
      }
    }
    
    // Enhance active members with family data
    const enhancedMembers = activeMembers.map(member => {
      const familyData = familyDataMap.get(member.personId) || {};
      return {
        ...member,
        familyId: familyData.familyId,
        familyRole: familyData.familyRole,
        membershipStartDate: familyData.membershipStartDate,
        yearsSinceMembership: familyData.yearsSinceMembership
      };
    });
    
    const withFamilyData = enhancedMembers.filter(m => m.familyId).length;
    console.log(`✅ Enhanced ${enhancedMembers.length} members: ${withFamilyData} have family data`);
    
    return enhancedMembers;
    
  } catch (error) {
    console.error('❌ Family data enhancement failed:', error.message);
    // Return original members if enhancement fails
    return activeMembers;
  }
}

/**
 * Main enhanced comparison function that uses family logic
 * @param {Object} exportData - Enhanced export data with family information
 * @returns {Object} Changes needed with family-aware "Not in GCG" processing
 */
function enhancedCompareWithFamilyLogic(exportData) {
  console.log('🔍 Enhanced comparison with family logic...');
  
  try {
    // Get the standard comparison results for GCG members
    const standardChanges = fixedCompareWithInactiveFiltering(exportData);
    
    // Calculate family-aware "Not in GCG" changes
    const notInGCGChanges = calculateNotInGCGChangesWithFamilyLogic(exportData);
    
    console.log('\n📊 ENHANCED COMPARISON RESULTS:');
    console.log(`🔄 GCG Member changes: ${standardChanges.additions.length + standardChanges.updates.length + standardChanges.removals.length}`);
    console.log(`👨‍👩‍👧‍👦 Not in GCG changes: ${notInGCGChanges.additions.length + notInGCGChanges.deletions.length}`);
    console.log(`📋 Family groups processed: ${notInGCGChanges.familyGroupsProcessed}`);
    
    return {
      ...standardChanges,
      notInGCGChanges: notInGCGChanges,
      familyProcessing: {
        familyGroupsProcessed: notInGCGChanges.familyGroupsProcessed,
        additionsCount: notInGCGChanges.additions.length,
        deletionsCount: notInGCGChanges.deletions.length
      }
    };
    
  } catch (error) {
    console.error('❌ Enhanced family comparison failed:', error.message);
    throw error;
  }
}

/**
 * Updated helper function to calculate Not in GCG changes (used by preview report)
 * This replaces the previous calculateNotInGCGChanges function
 * @param {Object} exportData - Full export data
 * @returns {Object} People to add/remove from "Not in GCG" tab with real family data
 */
function calculateNotInGCGChanges(exportData) {
  console.log('📋 Calculating Not in GCG changes with family logic...');
  
  try {
    // Use the enhanced family logic
    return calculateNotInGCGChangesWithFamilyLogic(exportData);
  } catch (error) {
    console.error('❌ FAMILY LOGIC ERROR (details):', error.message);
    console.error('❌ Error stack:', error.stack);
    console.warn('⚠️ Family logic failed, falling back to simple logic');    
    
    // Fallback to simple logic if family processing fails
    const notInGCGFromExport = exportData.membersWithGCGStatus.filter(m => 
      !m.gcgStatus.inGroup && m.isActiveMember && !m.isSynthetic
    );
    
    const additions = notInGCGFromExport.slice(0, 10).map(person => ({
      personId: person.personId,
      firstName: person.firstName,
      lastName: person.lastName,
      familyId: person.familyId || 'null',
      familyRole: person.familyRole || 'null'
    }));
    
    return {
      additions: additions,
      deletions: [],
      familyGroupsProcessed: 0
    };
  }
}

/**
 * Find column index by header name (case-insensitive) - reused utility
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

// Keep all existing functions from the original comparison-engine.js
// (getGCGMembersWithPersonId, isMarkedInactive, normalizeGroupName, 
//  fixedCompareWithInactiveFiltering, parseRealGCGDataWithGCGMembers, etc.)

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
    const exportMembers = exportData.activeMembers.filter(m => {
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
 * Enhanced fuzzy name matching that includes nicknames
 * Add these functions to comparison-engine.js
 */

/**
 * Create enhanced name lookup map that includes nicknames
 * @param {Array} activeMembers - Array of active members with nickname data
 * @returns {Map} Enhanced lookup map with multiple name variations
 */
function createEnhancedNameLookup(activeMembers) {
  console.log('🔍 Creating enhanced name lookup with nickname support...');
  
  const nameLookup = new Map();
  let nicknameCount = 0;
  
  activeMembers.forEach(member => {
    const personId = member.personId;
    const firstName = member.firstName || '';
    const lastName = member.lastName || '';
    const nickname = member.nickname || ''; // This should come from the Active Members export
    
    // Standard full name
    const fullName = `${firstName} ${lastName}`.trim().toLowerCase();
    if (fullName) {
      nameLookup.set(fullName, { personId, member, matchType: 'fullName' });
    }
    
    // If nickname exists, also add nickname + last name
    if (nickname && nickname.trim()) {
      const nicknameFullName = `${nickname.trim()} ${lastName}`.trim().toLowerCase();
      nameLookup.set(nicknameFullName, { personId, member, matchType: 'nickname' });
      nicknameCount++;
      
      console.log(`📝 Added nickname mapping: "${nicknameFullName}" for ${firstName} ${lastName} (ID: ${personId})`);
    }
    
    // Also add first name only variations for better fuzzy matching
    if (firstName) {
      const firstOnly = firstName.toLowerCase();
      if (!nameLookup.has(firstOnly)) {
        nameLookup.set(firstOnly, { personId, member, matchType: 'firstOnly' });
      }
    }
    
    if (nickname && nickname.trim()) {
      const nicknameOnly = nickname.trim().toLowerCase();
      if (!nameLookup.has(nicknameOnly)) {
        nameLookup.set(nicknameOnly, { personId, member, matchType: 'nicknameOnly' });
      }
    }
  });
  
  console.log(`✅ Enhanced lookup created: ${nameLookup.size} name variations, ${nicknameCount} nicknames`);
  return nameLookup;
}

/**
 * Enhanced name matching function that tries multiple strategies
 * @param {string} searchName - Name to search for (like "Moses Babatola")
 * @param {Map} nameLookup - Enhanced name lookup map
 * @returns {Object|null} Matching member data or null
 */
function findPersonByEnhancedName(searchName, nameLookup) {
  if (!searchName) return null;
  
  const searchLower = searchName.toLowerCase().trim();
  
  // Strategy 1: Exact match
  if (nameLookup.has(searchLower)) {
    const match = nameLookup.get(searchLower);
    console.log(`✅ Found exact match for "${searchName}": ${match.member.firstName} ${match.member.lastName} (${match.matchType})`);
    return match;
  }
  
  // Strategy 2: Try first name + last name separately
  const nameParts = searchLower.split(' ').filter(part => part.length > 0);
  if (nameParts.length >= 2) {
    const firstName = nameParts[0];
    const lastName = nameParts[nameParts.length - 1];
    
    // Look for matches where first name OR nickname matches, AND last name matches
    for (const [lookupName, data] of nameLookup.entries()) {
      const lookupParts = lookupName.split(' ');
      if (lookupParts.length >= 2) {
        const lookupFirst = lookupParts[0];
        const lookupLast = lookupParts[lookupParts.length - 1];
        
        if (lookupLast === lastName && lookupFirst === firstName) {
          console.log(`✅ Found component match for "${searchName}": ${data.member.firstName} ${data.member.lastName} (${data.matchType})`);
          return data;
        }
      }
    }
  }
  
  // Strategy 3: Fuzzy matching for common variations
  const commonVariations = [
    searchLower.replace(/\s+/g, ''), // Remove spaces
    searchLower.replace(/[^a-z\s]/g, ''), // Remove special characters
  ];
  
  for (const variation of commonVariations) {
    if (nameLookup.has(variation)) {
      const match = nameLookup.get(variation);
      console.log(`✅ Found fuzzy match for "${searchName}": ${match.member.firstName} ${match.member.lastName} (${match.matchType})`);
      return match;
    }
  }
  
  console.log(`❌ No match found for "${searchName}"`);
  return null;
}

/**
 * Enhanced populatePersonIds function that uses nickname matching
 * Replace the existing populatePersonIds function in comparison-engine.js
 */
function populatePersonIdsWithNicknames() {
  console.log('🔧 Populating Person IDs with nickname support...');
  
  try {
    // Get export data and current sheet data
    const exportData = parseRealGCGDataWithFamilyInfo();
    const config = getConfig();
    const ss = SpreadsheetApp.openById(config.SHEET_ID);
    const sheet = ss.getSheetByName('GCG Members');
    
    if (!sheet) {
      throw new Error('GCG Members sheet not found');
    }
    
    // Create enhanced name lookup with nicknames
    const nameLookup = createEnhancedNameLookup(exportData.activeMembers);
    
    // Get current sheet data
    const data = sheet.getDataRange().getValues();
    const headerRowIndex = findHeaderRow(data);
    const headers = data[headerRowIndex >= 0 ? headerRowIndex : 1];
    
    const cols = {
      personId: findColumnIndex(headers, 'Person ID'),
      firstName: findColumnIndex(headers, 'First'),
      lastName: findColumnIndex(headers, 'Last')
    };
    
    if (cols.personId === -1) {
      throw new Error('Person ID column not found');
    }
    
    let populated = 0;
    let nicknameMatches = 0;
    
    // Process each row
    const dataStartRow = (headerRowIndex >= 0 ? headerRowIndex : 1) + 1;
    for (let i = dataStartRow; i < data.length; i++) {
      const row = data[i];
      
      // Skip if Person ID already exists or no name data
      if (row[cols.personId] || !row[cols.firstName] || !row[cols.lastName]) {
        continue;
      }
      
      const searchName = `${row[cols.firstName]} ${row[cols.lastName]}`.trim();
      const match = findPersonByEnhancedName(searchName, nameLookup);
      
      if (match) {
        // Update the Person ID in the sheet
        sheet.getRange(i + 1, cols.personId + 1).setValue(match.personId);
        populated++;
        
        if (match.matchType === 'nickname' || match.matchType === 'nicknameOnly') {
          nicknameMatches++;
          console.log(`🎯 Nickname match: "${searchName}" → ${match.member.firstName} ${match.member.lastName} (nickname: ${match.member.nickname})`);
        }
      }
    }
    
    console.log(`✅ Populated ${populated} Person IDs (${nicknameMatches} via nickname matching)`);
    
    return {
      totalPopulated: populated,
      nicknameMatches: nicknameMatches,
      nameLookupSize: nameLookup.size
    };
    
  } catch (error) {
    console.error('❌ Enhanced Person ID population failed:', error.message);
    throw error;
  }
}

/**
 * Test function to check nickname matching
 */
function testNicknameMatching() {
  console.log('🧪 Testing nickname matching system...');
  
  try {
    const exportData = parseRealGCGDataWithFamilyInfo();
    const nameLookup = createEnhancedNameLookup(exportData.activeMembers);
    
    // Test cases
    const testCases = [
      'Moses Babatola', // Should find Olutimehin Babatola
      'Olutimehin Babatola', // Should find by real name
      'Moses', // Should find by nickname only
    ];
    
    console.log('🔍 Testing nickname matching:');
    testCases.forEach(testName => {
      const match = findPersonByEnhancedName(testName, nameLookup);
      if (match) {
        console.log(`✅ "${testName}" → ${match.member.firstName} ${match.member.lastName} (nickname: ${match.member.nickname || 'none'}, match type: ${match.matchType})`);
      } else {
        console.log(`❌ "${testName}" → No match found`);
      }
    });
    
    // Show all people with nicknames
    console.log('\n📝 All people with nicknames in export:');
    exportData.activeMembers
      .filter(m => m.nickname && m.nickname.trim())
      .forEach(member => {
        console.log(`   ${member.firstName} ${member.lastName} → "${member.nickname}"`);
      });
    
  } catch (error) {
    console.error('❌ Nickname test failed:', error.message);
  }
}

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

/**
 * Debug functions to test family logic step by step
 */

/**
 * Test the family logic processing step by step
 */
function testFamilyLogicProcessing() {
  console.log('🧪 Testing family logic processing step by step...');
  
  try {
    // Step 1: Get enhanced data with family info
    console.log('📊 Step 1: Getting enhanced data with family info...');
    const exportData = parseRealGCGDataWithFamilyInfo();
    console.log(`✅ Enhanced data loaded: ${exportData.activeMembers.length} active members`);
    
    // Step 2: Check if we have family data
    const withFamilyData = exportData.activeMembers.filter(m => m.familyId).length;
    console.log(`👨‍👩‍👧‍👦 Members with family data: ${withFamilyData}/${exportData.activeMembers.length}`);
    
    // Step 3: Test family representative calculation
    console.log('📊 Step 2: Testing family representative calculation...');
    const familyReps = calculateFamilyRepresentatives(exportData);
    console.log(`✅ Family representatives calculated: ${familyReps.length} representatives`);
    
    // Step 4: Test current "Not in GCG" reading
    console.log('📊 Step 3: Testing current "Not in GCG" tab reading...');
    const config = getConfig();
    const ss = SpreadsheetApp.openById(config.SHEET_ID);
    const currentNotInGCG = getCurrentNotInGCGMembers(ss);
    console.log(`✅ Current "Not in GCG" members: ${currentNotInGCG.length}`);
    
    // Step 5: Test full family logic
    console.log('📊 Step 4: Testing full family logic...');
    const familyChanges = calculateNotInGCGChangesWithFamilyLogic(exportData);
    console.log(`✅ Family logic results:`);
    console.log(`   ➕ Additions: ${familyChanges.additions.length}`);
    console.log(`   ➖ Deletions: ${familyChanges.deletions.length}`);
    console.log(`   👨‍👩‍👧‍👦 Family groups processed: ${familyChanges.familyGroupsProcessed}`);
    
    // Step 6: Show sample results
    if (familyChanges.additions.length > 0) {
      console.log('📋 Sample additions:');
      familyChanges.additions.slice(0, 5).forEach(person => {
        console.log(`   ${person.firstName} ${person.lastName} (Family: ${person.familyId}, Role: ${person.familyRole})`);
      });
    }
    
    if (familyChanges.deletions.length > 0) {
      console.log('📋 Sample deletions:');
      familyChanges.deletions.slice(0, 5).forEach(person => {
        console.log(`   ${person.firstName} ${person.lastName} (Family: ${person.familyId}, Role: ${person.familyRole})`);
      });
    }
    
    return familyChanges;
    
  } catch (error) {
    console.error('❌ Family logic test failed:', error.message);
    console.error('Stack trace:', error.stack);
    throw error;
  }
}

/**
 * Debug the Active Members family data parsing
 */
function debugActiveMembersFamilyData() {
  console.log('🔍 Debugging Active Members family data parsing...');
  
  try {
    // Re-read the Active Members file to inspect family data
    const activeMembersFile = findLatestFile('ACTIVE_MEMBERS');
    const spreadsheet = SpreadsheetApp.openById(activeMembersFile.getId());
    const dataSheet = spreadsheet.getSheets()[0];
    const data = dataSheet.getDataRange().getValues();
    
    if (data.length === 0) {
      console.error('❌ No data in Active Members sheet');
      return;
    }
    
    // Show headers to verify family columns exist
    const headers = data[0];
    console.log('📋 Active Members headers:');
    headers.forEach((header, index) => {
      if (header && header.toString().toLowerCase().includes('family')) {
        console.log(`   Column ${index}: ${header}`);
      }
    });
    
    // Find family-related columns
    const columnMap = {
      personId: findColumnIndex(headers, 'Breeze ID'),
      firstName: findColumnIndex(headers, 'First Name'),
      lastName: findColumnIndex(headers, 'Last Name'),
      familyId: findColumnIndex(headers, 'Family'),
      familyRole: findColumnIndex(headers, 'Family Role'),
      membershipStartDate: findColumnIndex(headers, 'Membership Start Date')
    };
    
    console.log('🔍 Column mapping:');
    Object.entries(columnMap).forEach(([field, index]) => {
      console.log(`   ${field}: ${index >= 0 ? `Column ${index} (${headers[index]})` : 'NOT FOUND'}`);
    });
    
    // Show sample family data
    console.log('📋 Sample family data (first 10 rows):');
    for (let i = 1; i <= Math.min(10, data.length - 1); i++) {
      const row = data[i];
      const personId = row[columnMap.personId];
      const firstName = row[columnMap.firstName];
      const lastName = row[columnMap.lastName];
      const familyId = row[columnMap.familyId];
      const familyRole = row[columnMap.familyRole];
      
      if (personId && firstName && lastName) {
        console.log(`   ${firstName} ${lastName} (ID: ${personId}) - Family: ${familyId || 'null'}, Role: ${familyRole || 'null'}`);
      }
    }
    
    // Count family data coverage
    let withFamilyId = 0;
    let withFamilyRole = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[columnMap.familyId]) withFamilyId++;
      if (row[columnMap.familyRole]) withFamilyRole++;
    }
    
    console.log(`📊 Family data coverage:`);
    console.log(`   With Family ID: ${withFamilyId}/${data.length - 1} (${(withFamilyId/(data.length-1)*100).toFixed(1)}%)`);
    console.log(`   With Family Role: ${withFamilyRole}/${data.length - 1} (${(withFamilyRole/(data.length-1)*100).toFixed(1)}%)`);
    
    return {
      totalRows: data.length - 1,
      withFamilyId: withFamilyId,
      withFamilyRole: withFamilyRole,
      columnMap: columnMap
    };
    
  } catch (error) {
    console.error('❌ Active Members family data debug failed:', error.message);
    throw error;
  }
}

/**
 * Test the current Not in GCG tab reading
 */
function debugCurrentNotInGCGTab() {
  console.log('🔍 Debugging current "Not in GCG" tab reading...');
  
  try {
    const config = getConfig();
    const ss = SpreadsheetApp.openById(config.SHEET_ID);
    const sheet = ss.getSheetByName('Not in a GCG');
    
    if (!sheet) {
      console.error('❌ "Not in a GCG" sheet not found');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    console.log(`📊 Sheet has ${data.length} rows`);
    
    // Show headers (should be in row 3)
    if (data.length >= 3) {
      const headers = data[2]; // Row 3 (index 2)
      console.log('📋 Headers in row 3:');
      headers.forEach((header, index) => {
        if (header) {
          console.log(`   Column ${index}: ${header}`);
        }
      });
      
      // Find Person ID column
      const personIdCol = findColumnIndex(headers, 'Person ID');
      console.log(`🔍 Person ID column: ${personIdCol >= 0 ? personIdCol : 'NOT FOUND'}`);
      
      // Show sample data
      if (data.length > 3) {
        console.log('📋 Sample data (first 5 rows):');
        for (let i = 3; i < Math.min(8, data.length); i++) {
          const row = data[i];
          if (row[personIdCol]) {
            console.log(`   Row ${i+1}: ${row[personIdCol]} - ${row[1] || ''} ${row[2] || ''}`);
          }
        }
      }
    }
    
    const currentMembers = getCurrentNotInGCGMembers(ss);
    console.log(`✅ Successfully read ${currentMembers.length} current "Not in GCG" members`);
    
    return currentMembers;
    
  } catch (error) {
    console.error('❌ Current "Not in GCG" tab debug failed:', error.message);
    throw error;
  }
}

/**
 * Simple test to see what the preview report is actually calling
 */
function debugPreviewReportDataFlow() {
  console.log('🔍 Debugging preview report data flow...');
  
  try {
    // Test what parseRealGCGDataWithGCGMembers returns
    console.log('📊 Testing standard data parsing...');
    const standardData = parseRealGCGDataWithGCGMembers();
    console.log(`✅ Standard data: ${standardData.activeMembers.length} active members`);
    
    // Test if family enhancement is working
    console.log('📊 Testing family data enhancement...');
    const familyData = parseRealGCGDataWithFamilyInfo();
    console.log(`✅ Family data: ${familyData.activeMembers.length} active members`);
    
    const withFamilyInfo = familyData.activeMembers.filter(m => m.familyId).length;
    console.log(`👨‍👩‍👧‍👦 Members with family info: ${withFamilyInfo}/${familyData.activeMembers.length}`);
    
    // Test what calculateNotInGCGChanges returns (this is what preview calls)
    console.log('📊 Testing calculateNotInGCGChanges (preview report function)...');
    const previewChanges = calculateNotInGCGChanges(standardData);
    console.log(`✅ Preview changes: ${previewChanges.additions.length} additions, ${previewChanges.deletions.length} deletions`);
    
    // Test the enhanced family version
    console.log('📊 Testing enhanced family logic...');
    const familyChanges = calculateNotInGCGChangesWithFamilyLogic(familyData);
    console.log(`✅ Family changes: ${familyChanges.additions.length} additions, ${familyChanges.deletions.length} deletions`);
    
    console.log('\n🔍 COMPARISON:');
    console.log(`Standard approach: ${previewChanges.additions.length} additions`);
    console.log(`Family approach: ${familyChanges.additions.length} additions`);
    console.log(`Difference: ${Math.abs(previewChanges.additions.length - familyChanges.additions.length)}`);
    
    return {
      standard: previewChanges,
      family: familyChanges
    };
    
  } catch (error) {
    console.error('❌ Preview report data flow debug failed:', error.message);
    throw error;
  }
}

/**
 * Debug function to see exactly what error is happening in the preview report
 */
function debugPreviewFamilyLogicError() {
  console.log('🔍 Debugging preview family logic error...');
  
  try {
    // Test the exact same flow as the preview report
    console.log('📊 Step 1: Get family-enhanced data (same as preview)...');
    const exportData = parseRealGCGDataWithFamilyInfo();
    console.log(`✅ Family-enhanced data loaded: ${exportData.activeMembers.length} members`);
    
    // Test calling calculateNotInGCGChanges exactly like the preview does
    console.log('📊 Step 2: Call calculateNotInGCGChanges (preview function)...');
    const previewChanges = calculateNotInGCGChanges(exportData);
    console.log(`✅ Preview changes: ${previewChanges.additions.length} additions, ${previewChanges.deletions.length} deletions`);
    
    // Check if we're getting TBD values
    if (previewChanges.additions.length > 0) {
      const firstAddition = previewChanges.additions[0];
      console.log('📋 First addition details:');
      console.log(`   Name: ${firstAddition.firstName} ${firstAddition.lastName}`);
      console.log(`   Family ID: ${firstAddition.familyId}`);
      console.log(`   Family Role: ${firstAddition.familyRole}`);
      
      if (firstAddition.familyId === 'TBD') {
        console.log('❌ PROBLEM: Getting TBD values - fallback logic is running!');
      } else {
        console.log('✅ Getting real family data - family logic is working!');
      }
    }
    
    // Test calling the family logic directly
    console.log('📊 Step 3: Call family logic directly...');
    const familyChanges = calculateNotInGCGChangesWithFamilyLogic(exportData);
    console.log(`✅ Direct family logic: ${familyChanges.additions.length} additions, ${familyChanges.deletions.length} deletions`);
    
    // Compare results
    console.log('\n🔍 COMPARISON:');
    console.log(`Preview function: ${previewChanges.additions.length} additions`);
    console.log(`Direct family function: ${familyChanges.additions.length} additions`);
    
    if (previewChanges.additions.length !== familyChanges.additions.length) {
      console.log('❌ MISMATCH: Preview is using fallback logic!');
      
      // Check if Babatola family is properly reduced
      const babatolaInPreview = previewChanges.additions.filter(p => p.lastName === 'Babatola').length;
      const babatolaInFamily = familyChanges.additions.filter(p => p.lastName === 'Babatola').length;
      
      console.log(`Babatola family in preview: ${babatolaInPreview} members`);
      console.log(`Babatola family in family logic: ${babatolaInFamily} members`);
    } else {
      console.log('✅ MATCH: Preview is using family logic correctly!');
    }
    
    return {
      preview: previewChanges,
      family: familyChanges
    };
    
  } catch (error) {
    console.error('❌ Debug failed:', error.message);
    console.error('Stack trace:', error.stack);
    throw error;
  }
}

/**
 * Direct test of the calculateNotInGCGChanges function to see what's happening
 */
function testCalculateNotInGCGChangesDirectly() {
  console.log('🧪 Testing calculateNotInGCGChanges function directly...');
  
  try {
    // Get family-enhanced data
    const exportData = parseRealGCGDataWithFamilyInfo();
    console.log('✅ Got family-enhanced data');
    
    // Add debug logging inside the try block
    console.log('📋 About to call calculateNotInGCGChanges...');
    console.log('📋 Calculating Not in GCG changes with family logic...');
    
    let result;
    try {
      console.log('🔄 Trying family logic...');
      result = calculateNotInGCGChangesWithFamilyLogic(exportData);
      console.log('✅ Family logic succeeded!');
    } catch (error) {
      console.error('❌ CAUGHT ERROR in family logic:', error.message);
      console.error('❌ Error type:', error.name);
      console.error('❌ Full error:', error);
      console.error('❌ Stack trace:', error.stack);
      
      // Manual fallback
      console.log('🔄 Using manual fallback...');
      const notInGCGFromExport = exportData.membersWithGCGStatus.filter(m => 
        !m.gcgStatus.inGroup && m.isActiveMember && !m.isSynthetic
      );
      
      result = {
        additions: notInGCGFromExport.slice(0, 10).map(person => ({
          personId: person.personId,
          firstName: person.firstName,
          lastName: person.lastName,
          familyId: person.familyId || 'TBD',
          familyRole: person.familyRole || 'TBD'
        })),
        deletions: [],
        familyGroupsProcessed: 0
      };
    }
    
    console.log(`✅ Final result: ${result.additions.length} additions, ${result.deletions.length} deletions`);
    
    // Check first addition
    if (result.additions.length > 0) {
      const first = result.additions[0];
      console.log(`📋 First addition: ${first.firstName} ${first.lastName}`);
      console.log(`   Family ID: ${first.familyId}`);
      console.log(`   Family Role: ${first.familyRole}`);
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ Outer test failed:', error.message);
    throw error;
  }
}

/**
 * Test if the calculateNotInGCGChanges function exists and what it returns
 */
function inspectCalculateNotInGCGChangesFunction() {
  console.log('🔍 Inspecting calculateNotInGCGChanges function...');
  
  // Check if function exists
  if (typeof calculateNotInGCGChanges === 'function') {
    console.log('✅ calculateNotInGCGChanges function exists');
    
    // Check function source (first few lines)
    const functionSource = calculateNotInGCGChanges.toString();
    const lines = functionSource.split('\n').slice(0, 10);
    console.log('📋 Function start:');
    lines.forEach((line, index) => {
      console.log(`   ${index + 1}: ${line}`);
    });
    
  } else {
    console.log('❌ calculateNotInGCGChanges function NOT found');
  }
  
  // Check if family function exists
  if (typeof calculateNotInGCGChangesWithFamilyLogic === 'function') {
    console.log('✅ calculateNotInGCGChangesWithFamilyLogic function exists');
  } else {
    console.log('❌ calculateNotInGCGChangesWithFamilyLogic function NOT found');
  }
}

/**
 * Debug function to see what's happening with exclusions
 * Add this to comparison-engine.js and run it
 */
function debugExclusions() {
  console.log('🔍 Debugging exclusions for Elders and Tuesday School...');
  
  try {
    // Get the family-enhanced data
    const exportData = parseRealGCGDataWithFamilyInfo();
    
    // Test the NEW exclusion function
    console.log('📊 Testing NEW getExcludedPersonIds (reads raw tags)...');
    const excludedIds = getExcludedPersonIds(exportData);
    console.log(`🚫 Found ${excludedIds.size} excluded people`);
    
    // Show first 10 excluded people
    console.log('📋 First 10 excluded people:');
    Array.from(excludedIds).slice(0, 10).forEach(personId => {
      // Try to find their name from active members
      const person = exportData.activeMembers.find(m => m.personId === personId);
      const name = person ? `${person.firstName} ${person.lastName}` : 'Unknown';
      console.log(`   🚫 ${personId} - ${name}`);
    });
    
    // Check specific people from your preview results who SHOULD be excluded
    const testPeople = [
      { name: 'Jacob Oldham', id: '35252936', reason: 'Should be excluded - Tuesday School member' },
      { name: 'Douglas Sturgeon', id: '29760286', reason: 'Should be excluded if Elder' }
    ];
    
    // Also check people who should NOT be excluded
    const shouldNotBeExcluded = [
      { name: 'Dan Gurtner', id: '29767590', reason: 'Should NOT be excluded - not an Elder' }
    ];
    
    console.log('\n🔍 Checking people who SHOULD be excluded:');
    testPeople.forEach(person => {
      const isExcluded = excludedIds.has(person.id);
      console.log(`   ${person.name} (${person.id}): ${isExcluded ? '🚫 EXCLUDED ✅' : '✅ INCLUDED ❌'} - ${person.reason}`);
    });
    
    console.log('\n🔍 Checking people who should NOT be excluded:');
    shouldNotBeExcluded.forEach(person => {
      const isExcluded = excludedIds.has(person.id);
      console.log(`   ${person.name} (${person.id}): ${isExcluded ? '🚫 EXCLUDED ❌' : '✅ INCLUDED ✅'} - ${person.reason}`);
    });
    
    return {
      excludedCount: excludedIds.size,
      excludedIds: Array.from(excludedIds)
    };
    
  } catch (error) {
    console.error('❌ Debug failed:', error.message);
    console.error('Stack trace:', error.stack);
    throw error;
  }
}

/**
 * Test inactive member enhancement Phase 3
 */

function testPhase3() {
  // Get the data from Phase 2
  const exportData = parseRealGCGDataWithInactiveMembers();
  
  // Test the Phase 3 enhanced comparison
  const comparisonData = enhancedCompareWithInactiveAwareness(exportData);
  
  console.log('📊 PHASE 3 TEST RESULTS:');
  console.log(`🔄 GCG Member changes: ${comparisonData.additions?.length || 0} additions, ${comparisonData.updates?.length || 0} updates, ${comparisonData.removals?.length || 0} removals`);
  console.log(`👥 Not in GCG changes: ${comparisonData.notInGCGChanges?.additions?.length || 0} additions, ${comparisonData.notInGCGChanges?.deletions?.length || 0} deletions`);
  console.log(`⚠️ NEW inactive members in GCGs: ${comparisonData.newInactiveInGCGs?.length || 0}`);
  console.log('📋 Inactive processing stats:', comparisonData.inactiveProcessing);
  
  return comparisonData;
}

/**
 * Debug active members to see what properties the active members have
 */

function debugActiveMemberProperties() {
  const exportData = parseRealGCGDataWithInactiveMembers();
  
  console.log('🔍 Checking active member properties...');
  
  if (exportData.activeMembers && exportData.activeMembers.length > 0) {
    const firstMember = exportData.activeMembers[0];
    console.log('📋 First active member properties:', Object.keys(firstMember));
    console.log('🎯 Sample member:', JSON.stringify(firstMember, null, 2));
    
    // Check if gcgStatus exists
    if (firstMember.gcgStatus) {
      console.log('✅ gcgStatus property exists!');
    } else {
      console.log('❌ gcgStatus property missing');
    }
  }
}
