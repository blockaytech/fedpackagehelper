# FedRAMP Package Helper

A Python toolkit for managing bulk updates to FedRAMP System Security Plan (SSP) documentation. Designed to help compliance teams efficiently update technology references, team names, positions, and other terms across large documentation packages while maintaining formatting and providing comprehensive audit trails.

## Features

- **Bulk Find & Replace**: Update technology names, team names, positions, and other terms across Word documents
- **Formatting Preservation**: Maintains bold, italic, and other text formatting during replacements
- **Preview Mode**: Review all changes with before/after context before applying
- **Term Discovery**: Automatically discovers potential new terms that may need tracking
- **Batch Processing**: Process entire documentation packages with a single command
- **Excel/CSV Export**: Export analysis results for stakeholder review
- **Audit Trail**: Complete logging of all changes made to documents
- **Delete Support**: Remove terms entirely while preserving surrounding context

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/fedpackagehelper.git
cd fedpackagehelper

# Create and activate virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Quick Start

### 1. Set Up Your Configuration

Copy the example configuration files and customize them:

```bash
cp examples/replacements.example.json replacements.json
cp examples/terms_dictionary.example.json terms_dictionary.json
```

Edit `replacements.json` with your actual replacement mappings:

```json
{
  "replacements": {
    "OldToolName": "NewToolName",
    "Old Team Name": "New Team Name",
    "Former CISO Title": "New CISO Title"
  }
}
```

### 2. Create Required Directories

```bash
mkdir -p originals output
```

Place your FedRAMP documentation (.docx files) in the `originals/` directory. These files will **never be modified** - changes are applied to copies in `drafts/`.

### 3. Analyze Your Package

```bash
python package_manager.py analyze
```

This scans all documents, identifies known terms, and discovers potential new terms.

### 4. Preview Changes

```bash
python package_manager.py preview
```

Review the changes that will be made with before/after context.

### 5. Apply Changes

```bash
python package_manager.py apply
```

This copies originals to `drafts/` and applies all replacements to the drafts. **Original files are preserved.**

### 6. Verify Drafts

```bash
python package_manager.py verify
```

Re-analyze the draft documents to verify all replacements were applied correctly and check for completeness.

## Usage

### Package Manager (Recommended)

The unified `package_manager.py` tool provides all functionality:

```bash
# Check package status and workflow state
python package_manager.py status

# Analyze documents for terms
python package_manager.py analyze

# Preview all changes
python package_manager.py preview

# Apply changes (originals -> drafts workflow)
python package_manager.py apply

# Verify drafts after applying changes
python package_manager.py verify

# Export analysis to Excel
python package_manager.py export --format excel

# Export analysis to CSV
python package_manager.py export --format csv
```

### Single Document Processing

For processing individual documents:

```bash
# Preview changes for one document
python ssp_bulk_update.py documents/your_ssp.docx --preview

# Apply changes to one document
python ssp_bulk_update.py documents/your_ssp.docx --output output/updated_ssp.docx
```

### Package Analyzer

For discovery and analysis only:

```bash
python package_analyzer.py --input documents/ --output analysis_report.txt
```

## Configuration Files

### replacements.json

Defines the find/replace mappings:

```json
{
  "_instructions": "Edit the 'new' values. Use 'DELETE' to remove terms entirely.",
  "replacements": {
    "Wiz": "YourSecurityTool",
    "Old Team Name": "New Team Name",
    "Deprecated Term": "DELETE"
  }
}
```

**Special Values:**
- `DELETE` or `REMOVE`: Removes the term entirely from the document

### terms_dictionary.json

Master dictionary of known terms to track:

```json
{
  "known_technologies": {
    "terms": {
      "Wiz": {"category": "security_scanning", "replacement": null},
      "Jira": {"category": "ticketing", "replacement": null}
    }
  },
  "known_teams": {
    "terms": {
      "Security Office": {"replacement": null},
      "Engineering Team": {"replacement": null}
    }
  },
  "known_positions": {
    "terms": {
      "Chief Information Security Officer": {"acronym": "CISO", "replacement": null}
    }
  }
}
```

## Directory Structure

```
fedpackagehelper/
├── package_manager.py      # Main unified tool
├── ssp_bulk_update.py      # Single document processor
├── package_analyzer.py     # Discovery and analysis tool
├── replacements.json       # Your replacement mappings (create from example)
├── terms_dictionary.json   # Your terms dictionary (create from example)
├── requirements.txt        # Python dependencies
├── originals/              # Source documents (NEVER modified)
├── drafts/                 # Working copies with changes applied
├── output/                 # Generated reports and exports
├── backups/                # Historical backups of drafts
└── examples/               # Example configuration files
    ├── replacements.example.json
    └── terms_dictionary.example.json
```

### Workflow Directory Usage

| Directory | Purpose | Modified? |
|-----------|---------|-----------|
| `originals/` | Source FedRAMP documents | Never |
| `drafts/` | Working copies with replacements applied | Yes |
| `output/` | Reports, exports, analysis results | Generated |
| `backups/` | Previous draft versions | Generated |

## Features in Detail

### Formatting Preservation

The tool replaces text at the run level within Word documents, preserving:
- Bold and italic formatting
- Font styles and sizes
- Colors and highlighting

### Replacement Ordering

Replacements are automatically ordered by length (longest first) to prevent partial replacement issues. For example, "FMSP Security Office" is replaced before "FMSP Security".

### Acronym Consistency

The tool warns if you're replacing a term but not its acronym (or vice versa):
- Replacing "Chief Information Security Officer" but not "CISO"
- Helps maintain document consistency

### Plural Form Detection

Automatically detects and warns about plural forms that may need separate replacements:
- "System Administrator" vs "System Administrators"
- "Engineer" vs "Engineers"

### Delete Mode

Set a replacement value to `DELETE` or `REMOVE` to completely remove the term:

```json
{
  "replacements": {
    "Deprecated Team Name": "DELETE"
  }
}
```

## Output Files

After running analysis or preview:

- `package_analysis_YYYYMMDD_HHMMSS.txt` - Human-readable analysis report
- `suggested_replacements_YYYYMMDD_HHMMSS.json` - Machine-readable term list
- `preview_YYYYMMDD_HHMMSS.txt` - Preview of all changes with context
- `changes_YYYYMMDD_HHMMSS.xlsx` - Excel export for review

## Best Practices

1. **Always Preview First**: Run `preview` before `apply` to verify changes
2. **Review Discovered Terms**: Check the analysis report for terms that may need tracking
3. **Keep Backups**: The tool creates backups, but consider version control for documents
4. **Test on Copies**: Test on document copies before processing originals
5. **Review Context**: Use the preview output to verify replacements make sense in context

## Troubleshooting

### "Module not found: docx"
```bash
source venv/bin/activate
pip install python-docx
```

### Partial Replacements
If "Security" is being replaced inside "Security Office", check that longer terms are in your replacements.json - the tool handles ordering automatically.

### Formatting Lost
Ensure you're using the latest version. Run-level replacement preserves formatting.

## Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## License

MIT License - see LICENSE file for details.

## Disclaimer

This tool is provided as-is. Always review changes before applying to official FedRAMP documentation. Maintain proper backups and version control of your compliance documentation.
