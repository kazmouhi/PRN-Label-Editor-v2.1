# APTIV PRN Label Editor

A professional Windows application for printing industrial labels using PRN template files and CSV data sources. Built with Python and customtkinter, this tool streamlines the label printing process for manufacturing and logistics operations.

![APTIV Logo](https://img.shields.io/badge/APTIV-Label%20Editor-FF6B00?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.8+-blue?style=for-the-badge&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

## Features

### 🖨️ Core Functionality
- **PRN Template Processing**: Load and customize PRN label templates with dynamic data
- **Batch Printing**: Process multiple label designs from CSV files with quantity control
- **Direct Printer Integration**: Send labels directly to network or local printers
- **Real-time Progress Tracking**: Monitor print jobs with detailed progress indicators
- **Auto-generated Date Codes**: Automatic date encoding using custom character mappings

### 📊 Data Management
- **CSV Import**: Load label data from CSV files with required columns (AFMPN, Country, QTY)
- **Country Code Mapping**: Automatic conversion of country names to ISO codes
- **Default Templates**: Auto-creates default CSV template on first launch
- **Data Validation**: Ensures all required fields are present before printing

### 🎨 User Interface
- **Modern Dark Theme**: Professional customtkinter interface with APTIV branding
- **Intuitive Layout**: Clean, organized interface with clear visual hierarchy
- **Status Messages**: In-app status display replacing popup dialogs
- **Progress Bars**: Dual progress tracking (per-label and total job)
- **Printer Selection**: Dropdown menu for available Windows printers

### 🔧 Advanced Features
- **Comprehensive Logging**: Rotating log files with detailed operation tracking
- **Error Handling**: Graceful error management with user-friendly messages
- **Thread-safe Printing**: Non-blocking print operations
- **Cancellable Jobs**: Stop print jobs mid-process
- **Template Variables**: Support for multiple dynamic placeholders

## Installation

### Prerequisites
- Windows OS (required for win32print)
- Python 3.8 or higher
- pip package manager

### Required Dependencies

Install all dependencies using pip:

```bash
pip install customtkinter pillow pandas pywin32
```

Or use the requirements file:

```bash
pip install -r requirements.txt
```

**requirements.txt:**
```
customtkinter>=5.2.0
pillow>=10.0.0
pandas>=2.0.0
pywin32>=306
```

## Usage

### Starting the Application

Run the application from command line:

```bash
python "PRN Label Editor v2.1.py"
```

### Workflow

1. **Select PRN Template**
   - Click "Select PRN" button
   - Choose your label template file (.prn format)
   - Template can contain placeholder variables (see Template Variables section)

2. **Prepare CSV Data**
   - Application auto-creates `default_labels.csv` on first launch
   - Required columns: `AFMPN`, `Country`, `QTY`
   - Optional column: `CTT` (if used in your template)
   - Click "Select CSV" to choose your data file

3. **Configure Printer**
   - Select target printer from dropdown menu
   - Verify computer name display (for logging purposes)

4. **Print Labels**
   - Click "Print Labels" button
   - Monitor progress bars for real-time status
   - Use "Cancel" button to stop job if needed

### CSV Format Example

```csv
AFMPN,Country,QTY,CTT
12345-67890,United States,10,ABC123
98765-43210,Germany,5,XYZ789
11111-22222,France,15,DEF456
```

### Template Variables

The application supports the following placeholder variables in PRN templates:

| Variable | Description | Example |
|----------|-------------|---------|
| `@AFMPN@` | Full part number | 12345-67890 |
| `@AFMPN5@` | Last 5 digits of part number | 67890 |
| `@Country@` | Full country name | United States |
| `@CountryCode@` | ISO 2-letter country code | US |
| `@CTT@` | Custom tracking code | ABC123 |
| `@Date@` | Encoded date (custom format) | 2aE |
| `@mon@` | Month (01-12) | 10 |
| `@day@` | Day (01-31) | 02 |
| `@yea@` | Year (2-digit) | 25 |
| `@hou@` | Hour (00-23) | 14 |
| `@min@` | Minute (00-59) | 30 |
| `@sec@` | Second (00-59) | 45 |

### Date Code Format

The `@Date@` variable uses a special encoding system:
- **Day**: 1-31 mapped to `123456789ABCDEFGHJKLMNPQRSTVWX`
- **Month**: 01-12 mapped to `abcdefghjkmn`
- **Year**: 2019-2025 mapped to `9ABCDEF`

Example: October 2nd, 2025 → `2aE`

## Logging

### Log Location
Logs are stored in `./log/logs.txt`

### Log Features
- Rotating file handler (5MB per file, 5 backups)
- Timestamps for all operations
- Computer name tracking
- Print job details
- Error tracking with stack traces

### Log Entry Examples
```
2025-10-02 14:30:15 - INFO - PRN Label Editor started
2025-10-02 14:30:15 - INFO - Computer name: WORKSTATION-01
2025-10-02 14:31:20 - INFO - Starting print job to HP LaserJet, total labels to print: 30
2025-10-02 14:32:45 - INFO - Completed: Printed 30 labels from 3 designs
```

## Supported Countries

The application includes ISO 3166-1 alpha-2 country codes for 249 countries and territories. Common examples:

- United States → US
- United Kingdom / Great Britain → GB
- Germany → DE
- France → FR
- China → CN
- Japan → JP
- Canada → CA
- Mexico → MX

Full list available in source code (`CountrieCode_map` dictionary).

## Troubleshooting

### Common Issues

**"No printers found"**
- Ensure printer drivers are installed
- Check Windows printer settings
- Verify printer is shared/accessible

**CSV file locked**
- Close Excel or other applications using the CSV
- Application will prompt for action if file is in use

**Print job fails**
- Verify PRN template format matches printer requirements
- Check printer queue for errors
- Review log files for detailed error messages

**Missing columns error**
- Ensure CSV has required headers: AFMPN, Country, QTY
- Check for typos in column names (case-sensitive)

### Debug Mode

For detailed debugging, check the log file at `./log/logs.txt`. The rotating log maintains up to 5 backup files (25MB total).

## Development

### Project Structure
```
PRN Label Editor v2.1.py    # Main application file
default_labels.csv           # Auto-generated template
Label.prn                    # Default PRN template
log/
  └── logs.txt              # Application logs
```

### Key Classes and Methods

**PRNEditor Class**
- `setup_logging()`: Configure rotating file handler
- `create_and_open_default_csv()`: Initialize default CSV template
- `process_and_print()`: Main print processing loop
- `print_to_printer()`: Direct printer communication
- `get_date_code()`: Generate encoded date strings

### Customization

To modify the APTIV branding:
1. Update `create_aptiv_logo()` method for logo changes
2. Modify color scheme (search for `#FF6B00` - APTIV orange)
3. Adjust layout in `setup_gui()` method

## Requirements

- **OS**: Windows (for win32print API)
- **Python**: 3.8+
- **Dependencies**: See Installation section
- **Printer**: Windows-compatible printer with RAW print support

## License

This project is provided as-is for use within APTIV manufacturing operations. Modify and distribute according to your organization's policies.

## Contributing

Contributions are welcome! Please follow these guidelines:
1. Fork the repository
2. Create a feature branch
3. Test thoroughly on Windows environment
4. Submit pull request with detailed description

## Support

For issues or questions:
- Check log files in `./log/` directory
- Review error messages in application status bar
- Verify CSV format matches requirements
- Ensure PRN template variables are correct

## Version History

**v2.1** (Current)
- Added comprehensive logging system
- Improved error handling
- Enhanced status messaging
- Default CSV auto-creation
- File lock detection

**v2.0**
- Full GUI redesign with customtkinter
- APTIV branding integration
- Progress tracking improvements
- Direct printer integration

---

**Built for APTIV Manufacturing Operations**

*Streamlining label printing with modern Python tools*
