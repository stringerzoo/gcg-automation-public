# GCG Automation System

A Google Apps Script-based automation system for managing Gospel Community Group (GCG) membership data at Immanuel Baptist Church. This system synchronizes data between Breeze Church Management System exports and Google Sheets, providing comprehensive preview reports and safe update processes.

## 🎯 Purpose

This automation system replaces manual monthly data synchronization by:
- **Reading Breeze exports automatically** from Google Drive using smart file detection
- **Comparing truth data** with current Google Sheets using Person ID matching
- **Generating comprehensive preview reports** before making any changes
- **Updating GCG membership tabs** with proper safeguards and data preservation
- **Preserving pastoral notes and care data** during updates
- **Handling family grouping logic** for "Not in a GCG" management

## 📊 What It Does

### Core Functions
- **Smart File Detection**: Automatically finds latest Breeze exports by date pattern matching
- **Data Comparison**: Identifies additions, updates, and removals with group name normalization
- **Comprehensive Preview Reports**: Shows all changes across 5 detailed sections
- **Safe Updates**: Preserves notes, handles inactive members, creates backups
- **Health Monitoring**: Validates system components and data integrity
- **Family Logic**: Groups families by head of household in "Not in a GCG" tab

### Data Sources
- **Active Members Export**: `immanuelky-people-[date].xlsx` from Breeze
- **Tags Export**: `immanuelky-tags-[date].xlsx` from Breeze  
- **Target Sheet**: "GCG Placement - Ministry Teams" Google Sheet

## 🎮 User Interface

The system adds a **"Breeze Update"** menu to your Google Sheet with:

- 📚 **Tutorial**: Step-by-step preparation guide for monthly updates
- 📊 **Generate Preview**: Creates comprehensive change report across 5 sections
- 🚀 **Apply Updates**: Safely applies changes with multiple confirmations
- ⚙️ **Configure Settings**: View and update system configuration
- 🔧 **Health Check**: Validates all system components and data files

## 📋 Preview Report Sections

The comprehensive preview report includes:

### 1. **GCG Summary** (A2:D39)
- Truth data vs current sheet counts per group
- **Breeze Export Count** vs **GCG Members Tab Count** vs **Inactive Count**
- Yellow highlighting for groups with discrepancies
- Hyperlinks to corresponding changes in Section 2

### 2. **Group-by-Group Changes** (A40+)
- Detailed additions and deletions by group with Person IDs
- Only shows groups that have actual changes (reduced clutter)
- Full name and Person ID for each change
- **Inactive** subsections for members marked as temporarily inactive

### 3. **Statistics** (F2:I5)
- Active member counts and participation rates
- Clear breakdown of GCG participation
- GCG members not in Active Members list count

### 4. **Not in GCG Updates** (F11+)
- Proposed changes to "Not in a GCG" tab
- **Enhanced family grouping logic**: One representative per family
- Prioritizes Head of Household > Spouse > Adult for family representatives
- Includes Family ID and Family Role columns

### 5. **Data Inconsistencies** (F26+)
- GCG members not found in Active Members list
- Clear instructions for resolving data issues
- Step-by-step process for cleaning up inconsistencies

## 🔧 Technical Architecture

### File Structure
```
/src/
├── config.js                 # Configuration management and script properties
├── smart-file-detection.js   # Automatic file finding with date patterns
├── google-sheets-parser.js   # Data parsing from Breeze exports
├── comparison-engine.js      # Core comparison logic with normalization
├── preview-report.js         # Comprehensive preview report generation
└── menu-system.js           # User interface and safety features
```

### Key Technical Features
- **Person ID Matching**: Uses Breeze Person ID for accurate cross-referencing
- **Inactive Filtering**: Automatically excludes members marked as inactive
- **Group Name Normalization**: Handles co-leader format differences (e.g., "Gene Cone" vs "Gene Cone & Scott Stringer")
- **Smart Family Grouping**: Manages "Not in GCG" tab by families using Family ID and Family Role
- **Data Validation**: Robust error handling and validation throughout
- **Safe Update Process**: Multiple confirmation steps and backup creation

## 🚀 Getting Started

See [SETUP.md](SETUP.md) for detailed installation and configuration instructions.

## 📖 Monthly Workflow

### Preparation Steps
1. **Export from Breeze**: Export both "Members - Active" and "Tags" to Google Sheets
2. **Upload to Drive**: Place exports in designated Google Drive folder
3. **Mark Inactive Members**: Use "Action Steps / Comments" column for temporarily inactive members
4. **Generate Preview**: Review all changes before applying

### Update Process
1. **Tutorial**: Review preparation steps in Breeze Update menu
2. **Health Check**: Verify system components are working properly
3. **Generate Preview**: Create comprehensive report of all proposed changes
4. **Review Changes**: Carefully examine all 5 sections of the preview report
5. **Apply Updates**: Execute changes with multiple confirmation prompts

## 🛡️ Safety Features

### Multiple Confirmation Steps
- Preview report generation required before updates
- "Are you sure?" prompts with change summaries
- Backup creation before any modifications

### Data Preservation
- Preserves all pastoral notes and care data
- Maintains inactive member tracking
- Handles co-leader format changes intelligently

### Error Handling
- Comprehensive health checks before operations
- Graceful handling of missing data
- Clear error messages with troubleshooting guidance

## 📈 Recent Improvements

### "Not in a GCG" Section Enhancement
- **Family grouping logic**: Only shows one person per family
- **Smart family representative selection**: Prioritizes head of household
- **Improved data accuracy**: Handles Family ID and Family Role properly
- **Reduced clutter**: Eliminates duplicate family member listings

### Preview Report Enhancements
- **Report header**: Added title and timestamp
- **Improved formatting**: Better column widths and visual separation
- **Enhanced hyperlinks**: Navigate directly from summary to details
- **Section buffering**: Clear visual separation between sections

## 🔍 Data Flow

```
Breeze Exports → Google Drive → Smart Detection → Data Parsing → 
Comparison Engine → Preview Report → User Review → Safe Updates → 
Updated GCG Placement Sheet
```

## 🤝 Support

For technical support or questions:
- **System Administrator**: sstringer@immanuelky.org
- **Health Check**: Run from Breeze Update menu for diagnostics
- **Tutorial**: Step-by-step guidance available in menu

## 📝 Version Notes

This system has been refined through extensive testing and feedback to provide:
- Accurate family logic for "Not in a GCG" management
- Intelligent group name normalization 
- Comprehensive preview reporting
- Safe, auditable update processes

The automation successfully handles the complexity of church membership data while preserving the pastoral care context that makes manual oversight valuable.
