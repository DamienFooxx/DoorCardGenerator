import requests
from pathlib import Path
import pandas as pd
import browser_cookie3

# Load the Excel file
excel_path = 'NOCTUA Doorcards.xlsx'  # Update this path to your file
df = pd.read_excel(excel_path)

# Print column names to debug
print("Available columns in the Excel file:")
for i, col in enumerate(df.columns):
    print(f"{i}: '{col}'")

# Directory for downloaded images
images_dir = Path('2526_photos')
images_dir.mkdir(exist_ok=True)

# Extract cookies from Chrome session (try both SharePoint and Forms domains)
try:
    cookies_jar = browser_cookie3.chrome(domain_name='sharepoint.com')
except:
    try:
        cookies_jar = browser_cookie3.chrome(domain_name='forms.office.com')
    except:
        cookies_jar = browser_cookie3.chrome(domain_name='microsoft.com')
session = requests.Session()
session.cookies.update(cookies_jar)

# Add headers to mimic browser requests
headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
    'Referer': 'https://forms.office.com/Pages/DesignPageV2.aspx?subpage=design&id=Xu-lWwkxd06Fvc_rDTR-gqvqKJCBFFxBuqb0ig6mrCFUMUdNOU5FS1cwOVdJREVOQ1hLNFZVSExHMy4u&analysis=true'  # Adjust if needed
}

# Define column names - you may need to adjust these based on the actual column names
# Let's try to find the photo column automatically
photo_col = None
name_col = None

for col in df.columns:
    if 'photo' in col.lower() or 'image' in col.lower():
        photo_col = col
    if 'name' in col.lower() and 'photo' not in col.lower():
        name_col = col

if photo_col is None:
    print("Could not find photo column. Please check the column names above and update the script.")
    exit(1)

if name_col is None:
    print("Could not find name column. Please check the column names above and update the script.")
    exit(1)

print(f"Using photo column: '{photo_col}'")
print(f"Using name column: '{name_col}'")

# Iterate over DataFrame to download each image
for index, row in df.iterrows():
    try:
        image_url = row[photo_col]  # Ensure this column name matches your Excel
        image_name = row[name_col]  # Ensure this column name matches your Excel
        
        # Skip if image_url is empty or NaN
        if pd.isna(image_url) or image_url == '':
            print(f"Skipping {image_name}: No image URL provided")
            continue
            
        # Skip if image_name is empty or NaN
        if pd.isna(image_name) or image_name == '':
            print(f"Skipping row {index}: No name provided")
            continue
        
        file_path = images_dir / f"{image_name}.jpg"

        print(f"Attempting to download {image_name} from: {image_url}")
        response = session.get(image_url, headers=headers, timeout=30)
        response.raise_for_status()
        
        with open(file_path, 'wb') as file:
            file.write(response.content)
        print(f"Downloaded {image_name}.jpg successfully ({len(response.content)} bytes)")
        
    except KeyError as e:
        print(f"Column not found: {e}")
        break
    except requests.RequestException as e:
        print(f"Failed to download {image_name if 'image_name' in locals() else 'unknown'} from {image_url if 'image_url' in locals() else 'unknown URL'}. Reason: {e}")
    except Exception as e:
        print(f"Unexpected error processing row {index}: {e}")
