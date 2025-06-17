/**
 * Google Sheets Parser Module for GCG Automation (Production Version)
 * Parses Google Sheets data for Active Members and GCG Tags
 */

/**
 * Validates if a tab name represents a real GCG group
 * @param {string} tabName - The tab name from Excel (e.g., "Gcg Aaron White")
 * @returns {boolean} - True if this is a real GCG group
 */
function isValidGCGGroup(tabName) {
  // Must start with "Gcg " (Excel removes colons and ampersands)
  if (!tabName.startsWith('Gcg ')) {
    return false;
  }
  
  // Extract the part after "Gcg "
  const namePart = tabName.substring(4); // Remove "Gcg "
  
  // Administrative tags to exclude (case-insensitive)
  const excludedTags = [
    'survey',
    'leaders',
    'leadership',
    'training',
    'coordinator',
    'admin',
    'all leaders',
    'leader training',
    'development',
    'wives',
    'active',
    'not',
    'resources',
    'materials',
    'planning'
  ];
  
  // Check if this matches any excluded patterns
  const lowerNamePart = namePart.toLowerCase();
  for (const excluded of excludedTags) {
    if (lowerNamePart.includes(excluded)) {
      return false;
    }
  }
  
  // Split by double space first (handles co-leaders)
  const parts = namePart.split('  '); // Double space where & was
  
  // Should have 1 or 2 parts (leader and optional co-leader)
  if (parts.length > 2) {
    return false;
  }
  
  // Validate each part has at least first and last name
  for (const part of parts) {
    const words = part.trim().split(' ').filter(word => word.length > 0);
    
    // Each part should have at least 2 words (first and last name)
    if (words.length < 2) {
      return false;
    }
    
    // Each word should be at least 2 characters and start with capital letter
    for (const word of words) {
      if (word.length < 2 || !/^[A-Z]/.test(word)) {
        return false;
      }
    }
  }
  
  return true;
}

/**
 * Parse Active Members Google Sheet
 * @param {GoogleAppsScript.Drive.File} file - Google Sheets file
 * @returns {Object} Parsed active members data
 */
function parseActiveMembersSheet(file) {
  console.log('📊 Parsing Active Members Google Sheet...');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const sheets = spreadsheet.getSheets();
    
    console.log(`📄 Sheet: ${spreadsheet.getName()}`);
    console.log(`📝 Found ${sheets.length} tabs: ${sheets.map(s => s.getName()).join(', ')}`);
    
    // Get the first sheet (should be the main data)
    const dataSheet = sheets[0];
    const data = dataSheet.getDataRange().getValues();
    
    if (data.length === 0) {
      throw new Error('No data found in Active Members sheet');
    }
    
    // First row should be headers
    const headers = data[0];
    console.log(`📋 Headers: ${headers.slice(0, 8).join(', ')}... (showing first 8)`);
    
    // Find important column indices
    const columnMap = {
      personId: findColumnIndex(headers, 'Breeze ID'),
      firstName: findColumnIndex(headers, 'First Name'),
      lastName: findColumnIndex(headers, 'Last Name'),
      streetAddress: findColumnIndex(headers, 'Street Address'),
      city: findColumnIndex(headers, 'City'),
      state: findColumnIndex(headers, 'State'),
      zip: findColumnIndex(headers, 'Zip')
    };
    
    // Validate required columns
    if (columnMap.personId === -1 || columnMap.firstName === -1 || columnMap.lastName === -1) {
      throw new Error('Required columns (Breeze ID, First Name, Last Name) not found');
    }
    
    // Process member data
    const members = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row[columnMap.personId]) continue;
      
      const member = {
        personId: String(row[columnMap.personId]),
        firstName: row[columnMap.firstName] || '',
        lastName: row[columnMap.lastName] || '',
        fullName: `${row[columnMap.firstName] || ''} ${row[columnMap.lastName] || ''}`.trim(),
        address: {
          street: row[columnMap.streetAddress] || '',
          city: row[columnMap.city] || '',
          state: row[columnMap.state] || '',
          zip: row[columnMap.zip] || ''
        },
        isActiveMember: true,
        sourceRow: i + 1
      };
      
      members.push(member);
    }
    
    console.log(`✅ Parsed ${members.length} active members`);
    
    return {
      members: members,
      totalCount: members.length,
      headers: headers,
      sheetName: dataSheet.getName(),
      lastUpdated: file.getLastUpdated()
    };
    
  } catch (error) {
    console.error('❌ Error parsing Active Members sheet:', error.message);
    throw error;
  }
}

/**
 * Parse Tags Google Sheet with multiple tabs
 * @param {GoogleAppsScript.Drive.File} file - Google Sheets file
 * @returns {Object} Parsed GCG assignments data
 */
function parseTagsSheet(file) {
  console.log('🏷️  Parsing Tags Google Sheet...');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const allSheets = spreadsheet.getSheets();
    
    console.log(`📄 Sheet: ${spreadsheet.getName()}`);
    console.log(`📝 Found ${allSheets.length} total tabs`);
    
    // Find GCG sheets using robust validation
    const gcgSheets = allSheets.filter(sheet => isValidGCGGroup(sheet.getName()));
    
    console.log(`🏘️  Found ${gcgSheets.length} individual GCG sheets`);
    
    // Parse each GCG sheet
    const groups = [];
    const assignments = {}; // personId -> group info
    let totalMembers = 0;
    
    gcgSheets.forEach(sheet => {
      try {
        const groupData = parseGCGSheet(sheet);
        groups.push(groupData);
        
        // Add to assignments lookup
        groupData.members.forEach(member => {
          if (member.personId) {
            assignments[member.personId] = {
              groupName: groupData.displayName,
              sheetName: groupData.sheetName,
              leader: groupData.leader,
              coLeader: groupData.coLeader,
              memberCount: groupData.memberCount
            };
          }
        });
        
        totalMembers += groupData.memberCount;
        
      } catch (error) {
        console.warn(`⚠️  Failed to parse sheet ${sheet.getName()}: ${error.message}`);
      }
    });
    
    console.log(`✅ Parsed ${groups.length} GCG groups with ${totalMembers} total members`);
    
    return {
      assignments: assignments,
      groups: groups,
      totalGroups: groups.length,
      totalMembers: totalMembers,
      lastUpdated: file.getLastUpdated(),
      sourceFile: file.getName()
    };
    
  } catch (error) {
    console.error('❌ Error parsing Tags sheet:', error.message);
    throw error;
  }
}

/**
 * Parse individual GCG sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Individual GCG sheet
 * @returns {Object} Group data
 */
function parseGCGSheet(sheet) {
  const sheetName = sheet.getName();
  const data = sheet.getDataRange().getValues();
  
  if (data.length === 0) {
    return {
      sheetName: sheetName,
      displayName: sheetName,
      leader: 'Unknown',
      coLeader: null,
      members: [],
      memberCount: 0
    };
  }
  
  // Extract leader info from sheet name
  const leaderInfo = extractLeaderFromSheetName(sheetName);
  
  // Parse members (assuming Person ID, First Name, Last Name format)
  const headers = data[0];
  const personIdCol = findColumnIndex(headers, 'Person ID');
  const firstNameCol = findColumnIndex(headers, 'First Name');
  const lastNameCol = findColumnIndex(headers, 'Last Name');
  
  const members = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (personIdCol >= 0 && row[personIdCol]) {
      const personId = String(row[personIdCol]).trim();
      const firstName = firstNameCol >= 0 ? String(row[firstNameCol] || '').trim() : '';
      const lastName = lastNameCol >= 0 ? String(row[lastNameCol] || '').trim() : '';
      
      if (personId && firstName && lastName) {
        members.push({
          personId: personId,
          firstName: firstName,
          lastName: lastName,
          sourceSheet: sheetName
        });
      }
    }
  }
  
  return {
    sheetName: sheetName,
    displayName: leaderInfo.displayName,
    leader: leaderInfo.leader,
    coLeader: leaderInfo.coLeader,
    members: members,
    memberCount: members.length
  };
}

/**
 * Extract leader name(s) from GCG sheet name
 * @param {string} sheetName - Name like "Gcg Gene Cone  Scott Stringer"
 * @returns {Object} Leader information
 */
function extractLeaderFromSheetName(sheetName) {
  // Remove "Gcg " prefix and trim
  const namesPart = sheetName.replace(/^Gcg\s+/i, '').trim();
  
  // Check for double space (indicates co-leaders)
  if (namesPart.includes('  ')) {
    const parts = namesPart.split('  ');
    const leader = parts[0].trim();
    const coLeader = parts[1].trim();
    
    return {
      leader: leader,
      coLeader: coLeader,
      displayName: `${leader} & ${coLeader}`
    };
  } else {
    return {
      leader: namesPart,
      coLeader: null,
      displayName: namesPart
    };
  }
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
