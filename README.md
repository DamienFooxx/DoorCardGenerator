# RC4DoorcardGenerator

Generates Doorcards from a power point template using Microsoft Response forms responses.

Handed down from Paturi Karthik, Hsu Stanley and modified and maintained by Damien Foo.

## Files/Folders You Don't Need to Modify

1. **`font` folder**: Contains required fonts (don't modify unless CSC requires font changes)
2. **`templates` folder**: Contains PowerPoint templates (don't modify if using Noctua template)
3. **`blobs` folder**: Contains template design documentation
4. **`main.py`**: Core script for generating doorcards (automatically handles most edge cases)

## Files/Folders You Will Be Modifying

1. **`xxx_photos/`**, **`xxx_pptx/`**, **`xxx_doorcards_png/`**: Rename these folders for the new semester/year
2. **`xxx_data.xlsx`**: Download from Microsoft Forms once data is collated
3. **`config.json`**: Update column names and file paths (see Configuration section below)
4. **`url_to_jpg.py`**: Update paths and URLs for photo downloading
'url_to_jpg.py': change "excel_path", "images_dir" variable and "cookies_jar" (Firefox or Chrome etc) accordingly, and copy paste the url of the Microsoft Forms responses page into the "Referer" variable. The url should look something like "https://nusu-my.sharepoint.com/:x:/r/personal/exxxxxxx_u_nus_edu/_layouts/..............Microsoft.Office.Excel.FMsFormsMetadataInWorkbookMetadata%3Atrue". Find the corresponding "User Agent" required by going to your browser, going to the Developer Tools page, then click on the "Network" tab, then click on one of the listed processes (the process will have a listed Status, Method, Domain etc.). Then under the "Headers" tab, look for the "Request Headers" tab and then find your "User Agent" to allow for authentication to download the images via the scripts.
## Configuration

### config.json Structure
```json
{
    "column": {
        "actualName": "Name",
        "displayName": "Name to be displayed", 
        "year": "Year of Study",
        "major": "Major",
        "caption": "Quote to be displayed (should be preferably less than 15 words)"
    },
    "location": {
        "excel": "NOCTUA Doorcards.xlsx",
        "template": "./templates/door_card.pptx",
        "font": "./font/DIN-Condensed-Bold.ttf",
        "photo": "./2526_photos",
        "target": "./2526_pptx"
    }
}
```

### Important Configuration Notes
- **Column Names**: If scripts don't recognize columns, check for hidden spaces or special characters
- **File Paths**: Ensure all paths point to existing directories/files
- **Column Mapping**: The `displayName` column should match the names in your image filenames

## Installation, Setup and Usage

### 1. Prerequisites
```bash
# Clone repository
git clone <repository-url>
cd doorcardgen

# Setup virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Setup Required Folders
```bash
# Create necessary directories (adjust names for your semester/year)
mkdir -p 2526_photos  # For downloaded photos
mkdir -p 2526_pptx    # For generated PowerPoint files
mkdir -p 2526doorcards_png  # For final PNG output
```

### 3. Configure the System
1. **Update `config.json`** with your specific column names and file paths
2. **Place your Excel file** in the root directory (update the `excel` path in config.json)
3. **Update `url_to_jpg.py`** if needed:
   - Set `excel_path` to your Excel file name
   - Set `images_dir` to your photos directory
   - Update the `Referer` URL if downloading from a different SharePoint location

### 4. Download Photos
```bash
# Run the photo download script
python url_to_jpg.py
```
**Note**: This script requires you to be logged into SharePoint in Chrome. Make sure you're authenticated before running.

### 5. Generate Doorcards
```bash
# Run the main generation script
python main.py
```

The script will:
- ✅ Check for existing PPTX files and offer to skip or recreate them
- ✅ Create intelligent name-to-image mappings
- ✅ Handle edge cases like multiple people with the same first name
- ✅ Validate mappings and report any issues
- ✅ Generate doorcards with progress tracking
- ✅ Provide detailed logging of the process

### 6. Convert to PNG
```bash
# For macOS
python ppt_to_png_(mac).py

# For Windows  
python ppt_to_png_(windows).py
```

### 7. Final Review
- Check the generated PNGs for any manual adjustments needed
- Verify image orientations and formatting
- Ensure all doorcards are generated correctly

## Troubleshooting

### Common Issues and Solutions

#### 1. Name Matching Issues
**Problem**: Script can't find images for some people
**Solutions**:
- Check if image filenames use full names vs first names only
- Ensure Excel names match image filename conventions
- Review the validation warnings in the console output

#### 2. Column Recognition Issues
**Problem**: Script doesn't recognize certain columns
**Solutions**:
- Run `print(df.columns)` in Python to see exact column names
- Check for hidden spaces or special characters in column names
- Update `config.json` with the exact column names

#### 3. Photo Download Issues
**Problem**: Photos don't download properly
**Solutions**:
- Ensure you're logged into SharePoint in Chrome
- Check the `Referer` URL in `url_to_jpg.py`
- Verify the Excel file has the correct photo URL column

#### 4. Existing Files Handling
**Problem**: Script asks about existing files every time
**Solution**: The script automatically detects existing files and offers options:
- **Option 1**: Skip existing files (recommended)
- **Option 2**: Force recreate all files (overwrites existing)

### Validation and Quality Assurance

The script includes several validation features:

1. **Name Mapping Validation**: Checks for duplicate mappings and unmapped names
2. **Image Quality Checks**: Validates image files and removes corrupted ones
3. **Progress Tracking**: Shows real-time progress and statistics
4. **Error Reporting**: Detailed error messages for troubleshooting

## Template Designing (if you need to redesign the template)

1. Open a new PowerPoint in the templates folder and open View > Slide Master
2. Create a new layout
3. Design your doorcard template and exit slide master after you are done
4. Create a new slide with your newly made layout
5. Open selection pane to rename the textboxes
6. Rename the textboxes accordingly:
   - `Name` - for the person's name
   - `Year` - for the year of study
   - `Major` - for the major
   - `Caption` - for the quote/caption
   - `Picture` - for the photo placeholder
7. Update the `template` path in `config.json` accordingly

## File Structure Example

```
doorcardgen/
├── 2526_photos/           # Downloaded photos
│   ├── John Doe.jpg
│   ├── Jane Smith.jpg
│   └── ...
├── 2526_pptx/            # Generated PowerPoint files
│   ├── JohnDoe_Noctua.pptx
│   ├── JaneSmith_Noctua.pptx
│   └── ...
├── 2526doorcards_png/    # Final PNG output
├── templates/            # PowerPoint templates
├── font/                # Required fonts
├── config.json          # Configuration file
├── main.py             # Main generation script
├── url_to_jpg.py       # Photo download script
└── requirements.txt    # Python dependencies
```

## Dependencies

- pandas >= 2.0.0
- Pillow >= 10.0.0
- python-pptx >= 0.6.21
- openpyxl >= 3.1.0
- browser-cookie3 >= 0.19.0
- requests >= 2.31.0

