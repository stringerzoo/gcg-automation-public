# GCG Automation System

A Google Apps Script-based automation system for managing Gospel Community Group (GCG) membership data at Immanuel Baptist Church. This system synchronizes data between Breeze Church Management System exports and Google Sheets.

## 🎯 Purpose

This automation system replaces manual monthly data synchronization by:
- Reading Breeze exports automatically from Google Drive
- Comparing truth data with current Google Sheets
- Generating preview reports before making changes
- Updating GCG membership tabs with proper safeguards
- Preserving pastoral notes and care data during updates

## 📊 What It Does

### Core Functions
- **Smart File Detection**: Automatically finds latest Breeze exports by date
- **Data Comparison**: Identifies additions, updates, and removals needed
- **Preview Reports**: Shows all changes before applying them
- **Safe Updates**: Preserves notes, handles inactive members, creates backups
- **Health Monitoring**: Validates system components and data integrity

### Data Sources
- **Active Members Export**: `immanuelky-people-[date].xlsx` from Breeze
- **Tags Export**: `immanuelky-tags-[date].xlsx` from Breeze  
- **Target Sheet**: "GCG Placement - Ministry Teams" Google Sheet

## 🎮 User Interface

The system adds a **"Breeze Update"** menu to your Google Sheet with:

- 📚 **Tutorial**: Step-by-step preparation guide
- 📊 **Generate Preview**: Creates comprehensive change report
- 🚀 **Apply Updates**: Safely applies changes with confirmations
- ⚙️ **Configure Settings**: View current system configuration
- 🔧 **Health Check**: Validates all system components

## 📋 Preview Report Sections

The preview report includes:

1. **GCG Summary** (A1:C39): Truth data vs current sheet counts per group
2. **Group-by-Group Changes** (A40+): Additions/deletions by group with Person IDs
3. **Statistics** (E1:F10): Active member counts and participation rates
4. **Not in GCG Updates** (E10+): Proposed changes to "Not in a GCG" tab
5. **Data Inconsistencies** (E25+): GCG members not in Active Members list

## 🔧 Technical Architecture

### File Structure
```
/src/
├── config.js                 # Configuration management
├── smart-file-detection.js   # Automatic file finding
├── google-sheets-parser.js   # Data parsing from exports
├── comparison-engine.js      # Core comparison logic
├── preview-report.js         # Preview report generation
└── menu-system.js           # User interface
```

### Key Features
- **Person ID Matching**: Uses Breeze Person ID for accurate comparisons
- **Inactive Filtering**: Automatically excludes members marked as inactive
- **Group Name Normalization**: Handles co-leader format differences
- **Family Grouping**: Manages "Not in GCG" tab by families
- **Data Validation**: Robust error handling and validation

## 🚀 Getting Started

See [SETUP.md](SETUP.md) for detailed installation and configuration instructions.

## 📖 Workflow Overview

### Monthly Process
1. **Export from Breeze**: Active Members and Tags to Google Sheets format
2. **Upload to Drive**: Place files in designated Google Drive folder
3. **Generate Preview**: Review all proposed changes carefully
4. **Apply Updates**: Execute changes with multiple confirmations
5. **Verify Results**: Check updated tabs and resolve any issues

### Data Flow
```
Breeze Exports → Google Drive → Smart Detection → Data Parsing → 
Comparison Engine → Preview Report → User Review → Safe Updates
```

## 🛡️ Safety Features

- **Preview First**: Never applies changes without user review
- **Multiple Confirmations**: Requires explicit approval before updates
- **Inactive Handling**: Respects "Action Steps" column for inactive members
- **Notes Preservation**: Maintains pastoral care data during updates
- **Backup Timestamps**: Creates restoration points
- **Health Monitoring**: Validates system integrity

## 📞 Support

- **System Administrator**: sstringer@immanuelky.org
- **Documentation**: See files in this repository
- **Troubleshooting**: Use the Health Check feature in the Breeze Update menu

## 📝 Version History

- **v2.0**: File-based automation system (production)
- **v1.0**: Breeze API approach (deprecated)

## 🔒 Security Notes

- Uses Google Apps Script's built-in authentication
- No external API keys required for file-based approach
- Operates within Google Workspace security boundaries
- Script Properties store only non-sensitive configuration

## 🤝 Contributing

This system is designed for church staff with minimal technical background. All changes should:
- Maintain the simple user interface
- Include comprehensive error handling
- Update documentation accordingly
- Test thoroughly before deployment

---

*Built with ❤️ for pastoral care and community management at Immanuel Baptist Church, Louisville, KY*
