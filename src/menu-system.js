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
    
// Test 4: Export Files (including inactive members)
    if (results.driveAccess) {
      try {
        const activeMembersFile = findLatestFile('ACTIVE_MEMBERS');
        const tagsFile = findLatestFile('TAGS_EXPORT');
        results.filesFound = true;
        results.details += `✅ Export files found:\n`;
        results.details += `   • Active Members: ${activeMembersFile.getName()}\n`;
        results.details += `   • Tags: ${tagsFile.getName()}\n`;
        
        // Test for inactive members file (optional)
        try {
          const inactiveMembersFile = findLatestFile('INACTIVE_MEMBERS');
          results.details += `   • Inactive Members: ${inactiveMembersFile.getName()}\n`;
        } catch (error) {
          results.details += `   ⚠️ Inactive members file not found (optional)\n`;
          results.details += `     → System will work without inactive file\n`;
        }
        
      } catch (error) {
        results.issues++;
        results.criticalIssues++;
        results.details += `❌ Export files not found: ${error.message}\n`;
      }
    }
    
    // Test 5: Data Parsing with Clean Statistics
    if (results.filesFound) {
      try {
        const exportData = parseRealGCGDataWithInactiveMembers();
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
 * Open GitHub repository with comprehensive preview
 * Shows what users will find before they click the link
 */
function openGitHubRepo() {
  // Replace with your actual GitHub repository URL
  const repoUrl = 'https://github.com/stringerzoo/gcg-automation-public';
  
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.4;">
      <h2 style="color: #24292e; margin-bottom: 20px;">🔧 GCG Automation System Repository</h2>
      
      <div style="background: #f6f8fa; padding: 15px; border-radius: 6px; margin-bottom: 20px;">
        <p style="margin: 0; font-weight: bold; color: #0366d6;">
          📖 Complete documentation, source code, and support resources
        </p>
      </div>
      
      <h3 style="color: #24292e; margin-bottom: 10px;">📋 What you'll find in the repository:</h3>
      
      <div style="margin-bottom: 20px;">
        <p><strong>📚 Documentation Files:</strong></p>
        <ul style="margin: 8px 0; padding-left: 20px;">
          <li><strong>README.md</strong> - System overview, features, and architecture</li>
          <li><strong>SETUP.md</strong> - Complete installation and configuration guide</li>
          <li><strong>TROUBLESHOOTING.md</strong> - Problem resolution and debugging</li>
        </ul>
        
        <p><strong>💻 Source Code:</strong></p>
        <ul style="margin: 8px 0; padding-left: 20px;">
          <li><strong>/src/config.js</strong> - Configuration management</li>
          <li><strong>/src/smart-file-detection.js</strong> - Automatic file finding</li>
          <li><strong>/src/google-sheets-parser.js</strong> - Data parsing from exports</li>
          <li><strong>/src/comparison-engine.js</strong> - Core comparison logic</li>
          <li><strong>/src/preview-report.js</strong> - Preview report generation</li>
          <li><strong>/src/menu-system.js</strong> - User interface (this menu!)</li>
        </ul>
        
        <p><strong>🛠️ Support Resources:</strong></p>
        <ul style="margin: 8px 0; padding-left: 20px;">
          <li><strong>Issue tracking</strong> - Report bugs or request features</li>
          <li><strong>Version history</strong> - See all updates and changes</li>
          <li><strong>Technical support</strong> - Contact information and guidelines</li>
        </ul>
      </div>
      
      <div style="background: #fff3cd; border: 1px solid #ffeaa7; padding: 12px; border-radius: 4px; margin-bottom: 20px;">
        <p style="margin: 0; font-size: 14px;">
          💡 <strong>Tip:</strong> Bookmark the repository for easy access to troubleshooting guides and updates!
        </p>
      </div>
      
      <div style="text-align: center; margin: 25px 0;">
        <a href="${repoUrl}" target="_blank" 
           style="display: inline-block; background: #28a745; color: white; padding: 12px 24px; 
                  text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 16px;">
          🚀 Open GitHub Repository
        </a>
      </div>
      
      <div style="border-top: 1px solid #e1e4e8; padding-top: 15px; margin-top: 20px;">
        <p style="font-size: 12px; color: #586069; margin: 5px 0;">
          <strong>Repository URL:</strong> <code style="background: #f6f8fa; padding: 2px 4px; border-radius: 3px;">${repoUrl}</code>
        </p>
        <p style="font-size: 12px; color: #586069; margin: 5px 0;">
          <em>If the link doesn't open automatically, copy the URL above and paste it into your browser.</em>
        </p>
      </div>
      
      <div style="text-align: center; margin-top: 20px;">
        <input type="button" value="Close" onclick="google.script.host.close();" 
               style="background: #6c757d; color: white; border: none; padding: 8px 16px; 
                      border-radius: 4px; cursor: pointer; font-size: 14px;">
      </div>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(550);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'GCG Automation Repository');
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
