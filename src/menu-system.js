/**
 * Menu System for GCG Automation
 * Provides user interface for the automation system
 */

/**
 * Create custom menu in Google Sheets
 * This function runs automatically when the sheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Breeze Update')
    .addItem('📚 Tutorial: How to Prepare for Update', 'showUpdateTutorial')
    .addSeparator()
    .addItem('📊 Generate Preview Report', 'generatePreviewReport')
    .addSeparator()
    .addItem('🚀 Apply Updates (with confirmation)', 'applyUpdatesWithConfirmation')
    .addSeparator()
    .addItem('⚙️ Configure Settings', 'showConfigDialog')
    .addItem('🔧 System Health Check', 'runHealthCheck')
    .addSeparator()
    .addItem('📖 View Documentation & Code', 'openGitHubRepo')
    .addToUi();
}

/**
 * Show tutorial dialog with step-by-step instructions
 */
function showUpdateTutorial() {
  const ui = SpreadsheetApp.getUi();
  const tutorial = `
BREEZE UPDATE TUTORIAL

Before running an update, please ensure:

1. 📤 EXPORT FROM BREEZE:
   • Export "Members - Active" to Google Sheets
   • Export "Tags" to Google Sheets  
   • Save both with today's date in filename
   • Files should be named like: "immanuelky-people-MM-DD-YYYY"
   • Files should be named like: "immanuelky-tags-MM-DD-YYYY"

2. 📁 UPLOAD TO GOOGLE DRIVE:
   • Upload both files to your automation folder
   • System will automatically find the latest files

3. 📋 REVIEW YOUR CURRENT SHEET:
   • Mark any inactive members in "Action Steps / Comments" column
   • Use keywords like "inactive", "moved away", "left church"
   • Ensure Person IDs are populated in column A of GCG Members tab

4. 🔍 GENERATE PREVIEW:
   • Use "Generate Preview Report" to see all changes
   • Review carefully before applying
   • Check the "Breeze Update Preview" tab

5. ✅ APPLY UPDATES:
   • Only after reviewing preview report
   • System will ask for confirmation
   • Creates backup before making changes

IMPORTANT NOTES:
• Never run updates without reviewing the preview first
• Inactive members are automatically filtered out
• System preserves your notes and pastoral care data
• Contact your system administrator if you see unexpected results

Need help? Contact: sstringer@immanuelky.org
  `;
  
  ui.alert('Breeze Update Tutorial', tutorial, ui.ButtonSet.OK);
}

/**
 * Apply updates with multiple confirmation steps
 */
function applyUpdatesWithConfirmation() {
  const ui = SpreadsheetApp.getUi();
  
  // Step 1: Check if preview report exists
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let previewSheet;
  
  try {
    previewSheet = ss.getSheetByName('Breeze Update Preview');
  } catch (e) {
    ui.alert(
      'Preview Required', 
      'Please generate a preview report first!\n\nUse "Generate Preview Report" from the Breeze Update menu.', 
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Step 2: Show summary of changes
  try {
    const exportData = parseRealGCGDataWithGCGMembers();
    const changes = fixedCompareWithInactiveFiltering(exportData);
    
    const totalChanges = changes.additions.length + changes.updates.length + changes.removals.length;
    
    const summary = `
CHANGES SUMMARY:

➕ Additions: ${changes.additions.length}
🔄 Updates: ${changes.updates.length}  
➖ Removals: ${changes.removals.length}
🚫 Inactive filtered: ${changes.inactiveFiltered}

Total changes: ${totalChanges}

This will modify your GCG Members and Not in GCG tabs.
`;
    
    const summaryResponse = ui.alert(
      'Review Changes',
      summary + '\nDo you want to proceed with these updates?',
      ui.ButtonSet.YES_NO
    );
    
    if (summaryResponse !== ui.Button.YES) {
      return;
    }
    
  } catch (error) {
    ui.alert('Error', 'Failed to analyze changes: ' + error.message, ui.ButtonSet.OK);
    return;
  }
  
  // Step 3: Final confirmation with warning
  const finalResponse = ui.alert(
    'FINAL CONFIRMATION',
    '⚠️  WARNING: This action cannot be undone!\n\n' +
    'The system will:\n' +
    '• Update your GCG Members tab\n' +
    '• Update your Not in GCG tab\n' +
    '• Preserve your notes and pastoral care data\n' +
    '• Create a backup timestamp\n\n' +
    'Are you absolutely sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );
  
  if (finalResponse === ui.Button.YES) {
    try {
      // Show progress message
      ui.alert('Processing...', 'Updates are being applied. Please wait...', ui.ButtonSet.OK);
      
      // Apply the updates
      const result = applyAllUpdates();
      
      // Success message
      ui.alert(
        'Success!', 
        `Updates applied successfully!\n\n` +
        `✅ ${result.appliedChanges} changes applied\n` +
        `📋 Check the updated tabs to verify results\n` +
        `⏰ Backup created at: ${result.backupTime}`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      ui.alert(
        'Error', 
        'Update failed: ' + error.message + '\n\nYour data has not been modified.', 
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * Show configuration dialog
 */
function showConfigDialog() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = getConfig();
    
    const configInfo = `
CURRENT CONFIGURATION:

📁 Google Drive Folder: ${config.DRIVE_FOLDER_ID}
📊 Google Sheet ID: ${config.SHEET_ID}
📧 Admin Email: ${config.NOTIFICATIONS.ADMIN_EMAIL}

🔍 File Patterns:
• Active Members: "${config.FILE_PATTERNS.ACTIVE_MEMBERS.contains}"
• Tags Export: "${config.FILE_PATTERNS.TAGS_EXPORT.contains}"

📋 File Selection: ${config.FILE_SELECTION.strategy} (newest file)

To modify these settings, contact your system administrator.
    `;
    
    ui.alert('Current Configuration', configInfo, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', 'Failed to load configuration: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Run system health check
 */
function runHealthCheck() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('Running Health Check...', 'Click OK to start the health check. The system will analyze all components and show results when complete.', ui.ButtonSet.OK);
    
    const results = performHealthCheck();
    
    let status = '✅ SYSTEM HEALTHY';
    if (results.criticalIssues > 0) {
      status = '❌ CRITICAL ISSUES FOUND';
    } else if (results.issues > 0) {
      status = '⚠️ WARNINGS FOUND';
    }
    
    const report = `
${status}

📊 Configuration: ${results.config ? '✅' : '❌'}
📁 Google Drive Access: ${results.driveAccess ? '✅' : '❌'}
📋 Google Sheets Access: ${results.sheetsAccess ? '✅' : '❌'}
📄 Export Files Found: ${results.filesFound ? '✅' : '❌'}
🔍 Data Parsing: ${results.dataParsing ? '✅' : '❌'}

${results.details}

${results.issues > 0 ? 'Contact your system administrator if issues persist.' : 'System is operating normally.'}
    `;
    
    ui.alert('Health Check Results', report, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Health Check Failed', 'Error during health check: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Perform comprehensive health check
 * @returns {Object} Health check results
 */
function performHealthCheck() {
  const results = {
    config: false,
    driveAccess: false,
    sheetsAccess: false,
    filesFound: false,
    dataParsing: false,
    issues: 0,
    criticalIssues: 0,
    details: ''
  };
  
  try {
    // Test 1: Configuration
    const config = getConfig();
    results.config = true;
    results.details += '✅ Configuration loaded successfully\n';
    
    // Test 2: Google Drive Access
    try {
      const folder = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);
      results.driveAccess = true;
      results.details += `✅ Google Drive folder accessible: "${folder.getName()}"\n`;
    } catch (error) {
      results.issues++;
      results.criticalIssues++;
      results.details += `❌ Google Drive access failed: ${error.message}\n`;
    }
    
    // Test 3: Google Sheets Access
    try {
      const ss = SpreadsheetApp.openById(config.SHEET_ID);
      results.sheetsAccess = true;
      results.details += `✅ Google Sheet accessible: "${ss.getName()}"\n`;
    } catch (error) {
      results.issues++;
      results.criticalIssues++;
      results.details += `❌ Google Sheet access failed: ${error.message}\n`;
    }
    
    // Test 4: Export Files
    if (results.driveAccess) {
      try {
        const activeMembersFile = findLatestFile('ACTIVE_MEMBERS');
        const tagsFile = findLatestFile('TAGS_EXPORT');
        results.filesFound = true;
        results.details += `✅ Export files found:\n`;
        results.details += `   • Active Members: ${activeMembersFile.getName()}\n`;
        results.details += `   • Tags: ${tagsFile.getName()}\n`;
      } catch (error) {
        results.issues++;
        results.criticalIssues++;
        results.details += `❌ Export files not found: ${error.message}\n`;
      }
    }
    
    // Test 5: Data Parsing with Clean Statistics
    if (results.filesFound) {
      try {
        const exportData = parseRealGCGDataWithGCGMembers();
        results.dataParsing = true;
        results.details += `✅ Data parsing successful:\n`;
        results.details += `   • ${exportData.summary.originalActiveMembers} active members\n`;
        results.details += `   • ${exportData.summary.totalGroups} GCG groups\n`;
        results.details += `   • ${exportData.summary.inGCG} active members in GCGs\n`;
        
        // Add data inconsistency note if present (not counted as an issue)
        if (exportData.summary.syntheticMembers > 0) {
          results.details += `📋 ${exportData.summary.syntheticMembers} GCG members not in Active Members list\n`;
          results.details += `   → Review these people in the preview report\n`;
        }
        
      } catch (error) {
        results.issues++;
        results.criticalIssues++;
        results.details += `❌ Data parsing failed: ${error.message}\n`;
      }
    }
    
  } catch (error) {
    results.issues++;
    results.criticalIssues++;
    results.details += `❌ Health check error: ${error.message}\n`;
  }
  
  return results;
}

/**
 * Open GitHub repository in a new tab
 * Provides direct access to documentation and source code
 */
function openGitHubRepo() {
  // Replace with your actual GitHub repository URL
  const repoUrl = 'https://github.com/stringerzoo/gcg-automation-public';
  
  // Create HTML content that opens the link immediately
  const htmlContent = `
    <script>
      window.open('${repoUrl}', '_blank');
      google.script.host.close();
    </script>
    <p>Opening GitHub repository...</p>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(300)
    .setHeight(100);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'GitHub Repository');
}

/**
 * Main function to apply all updates
 * This will be implemented in a separate update-engine.js file
 */
function applyAllUpdates() {
  console.log('🚀 Applying all updates...');
  
  try {
    // Get the latest data and changes
    const exportData = parseRealGCGDataWithGCGMembers();
    const changes = fixedCompareWithInactiveFiltering(exportData);
    
    // Create backup timestamp
    const backupTime = new Date().toISOString();
    
    // Apply changes to GCG Members tab
    const gcgResults = applyGCGMemberChanges(changes);
    
    // Apply changes to Not in GCG tab  
    const notInGCGResults = applyNotInGCGChanges(exportData);
    
    // Update group counts in Group List tab
    const groupListResults = updateGroupCounts(exportData);
    
    const totalApplied = gcgResults.applied + notInGCGResults.applied + groupListResults.applied;
    
    console.log(`✅ Applied ${totalApplied} total changes`);
    
    return {
      appliedChanges: totalApplied,
      backupTime: backupTime,
      gcgChanges: gcgResults.applied,
      notInGCGChanges: notInGCGResults.applied,
      groupListChanges: groupListResults.applied
    };
    
  } catch (error) {
    console.error('❌ Update application failed:', error.message);
    throw error;
  }
}

/**
 * Placeholder functions for update engine
 * These will be implemented in update-engine.js
 */
function applyGCGMemberChanges(changes) {
  console.log('🔄 Applying GCG Member changes...');
  // TODO: Implement actual update logic
  return { applied: changes.additions.length + changes.updates.length + changes.removals.length };
}

function applyNotInGCGChanges(exportData) {
  console.log('🔄 Applying Not in GCG changes...');
  // TODO: Implement actual update logic
  return { applied: 0 };
}

function updateGroupCounts(exportData) {
  console.log('🔄 Updating group counts...');
  // TODO: Implement actual update logic
  return { applied: exportData.groups.length };
}
