/**
 * Smart file detection system for GCG Automation (Production Version)
 * Handles dated filenames automatically
 */

/**
 * Initialize smart file detection configuration
 */
function updateConfigForSmartDetection() {
  const smartConfig = {
    DRIVE_FOLDER_ID: '1P7JKNiYgFcQh6TtzHS1q9DofASu6pfYh',
    SHEET_ID: '1H_bKbWbSTCBJWffd4bbGiRpbxjcxIhq2hb7ICUt1-M0',
    
    FILE_PATTERNS: {
      ACTIVE_MEMBERS: {
        contains: 'immanuelky-people',
        excludes: [],
        description: 'Active Members Export'
      },
      TAGS_EXPORT: {
        contains: 'immanuelky-tags',
        excludes: [],
        description: 'Tags Export'
      }
    },
    
    FILE_SELECTION: {
      strategy: 'latest',
      dateFormats: [
        'MM-dd-yyyy',
        'MM-dd-yy', 
        'yyyy-MM-dd',
        'dd-MM-yyyy'
      ]
    },
    
    NOTIFICATIONS: {
      ADMIN_EMAIL: 'sstringer@immanuelky.org',
      SEND_CHANGE_NOTIFICATIONS: true,
      SEND_ERROR_NOTIFICATIONS: true
    }
  };
  
  // Save configuration
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('GCG_FILE_CONFIG', JSON.stringify(smartConfig));
  
  console.log('✅ Smart file detection configuration saved');
  return smartConfig;
}

/**
 * Smart file finder that automatically picks the latest file matching patterns
 * @param {string} fileType - Type of file to find ('ACTIVE_MEMBERS' or 'TAGS_EXPORT')
 * @returns {GoogleAppsScript.Drive.File} Latest matching file
 */
function findLatestFile(fileType) {
  const config = getConfig();
  const folder = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);
  const pattern = config.FILE_PATTERNS[fileType];
  
  if (!pattern) {
    throw new Error(`Unknown file pattern: ${fileType}`);
  }
  
  console.log(`🔍 Looking for ${pattern.description} files containing: "${pattern.contains}"`);
  
  const matchingFiles = [];
  const files = folder.getFiles();
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // Check if file matches our pattern
    if (fileName.toLowerCase().includes(pattern.contains.toLowerCase())) {
      // Check exclusions
      const excluded = pattern.excludes.some(exclude => 
        fileName.toLowerCase().includes(exclude.toLowerCase())
      );
      
      if (!excluded) {
        matchingFiles.push({
          file: file,
          name: fileName,
          lastModified: file.getLastUpdated(),
          size: file.getSize()
        });
      }
    }
  }
  
  if (matchingFiles.length === 0) {
    throw new Error(`No files found matching pattern "${pattern.contains}" for ${pattern.description}`);
  }
  
  // Sort by last modified date (newest first)
  matchingFiles.sort((a, b) => b.lastModified.getTime() - a.lastModified.getTime());
  
  const selectedFile = matchingFiles[0];
  
  console.log(`📄 Found ${matchingFiles.length} matching files for ${pattern.description}`);
  console.log(`✅ Selected: ${selectedFile.name} (modified ${selectedFile.lastModified})`);
  
  return selectedFile.file;
}

/**
 * Parse real GCG data using smart file detection
 * This is the main function that combines everything
 */
function parseRealGCGDataSmart() {
  console.log('🎯 Parsing REAL GCG data (with smart file detection)...');
  
  try {
    // Parse Active Members using smart detection
    console.log('📊 Parsing Active Members...');
    const activeMembersFile = findLatestFile('ACTIVE_MEMBERS');
    const activeMembersResult = parseActiveMembersSheet(activeMembersFile);
    const activeMembers = activeMembersResult.members;
    
    console.log(`✅ Parsed ${activeMembers.length} active members`);
    
    // Parse GCG Tags using smart detection
    console.log('🏷️ Parsing GCG assignments...');
    const tagsFile = findLatestFile('TAGS_EXPORT');
    const tagsResult = parseTagsSheet(tagsFile);
    
    console.log(`🏘️ Found ${tagsResult.totalGroups} GCG groups with ${tagsResult.totalMembers} total members`);
    
    // Cross-reference data
    const membersWithGCGStatus = activeMembers.map(member => {
      const gcgAssignment = tagsResult.assignments[member.personId];
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
    
    const inGCG = membersWithGCGStatus.filter(m => m.gcgStatus.inGroup).length;
    const notInGCG = membersWithGCGStatus.filter(m => !m.gcgStatus.inGroup).length;
    const participationRate = (inGCG / activeMembers.length * 100).toFixed(1);
    
    console.log('\n📊 FINAL RESULTS:');
    console.log(`👥 Active Members: ${activeMembers.length}`);
    console.log(`🏘️ GCG Groups: ${tagsResult.totalGroups}`);
    console.log(`✅ In GCGs: ${inGCG} (${participationRate}%)`);
    console.log(`❌ Not in GCGs: ${notInGCG}`);
    
    return {
      activeMembers: activeMembers,
      membersWithGCGStatus: membersWithGCGStatus,
      groups: tagsResult.groups,
      assignments: tagsResult.assignments,
      summary: {
        totalActiveMembers: activeMembers.length,
        totalGroups: tagsResult.totalGroups,
        inGCG: inGCG,
        notInGCG: notInGCG,
        participationRate: participationRate
      }
    };
    
  } catch (error) {
    console.error('❌ Smart data parsing failed:', error.message);
    throw error;
  }
}
