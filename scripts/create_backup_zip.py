import os
import zipfile
import shutil

zip_filename = "Team_CamGiang_Full_Project_And_Brain_Backup.zip"
temp_zip = "temp_backup.zip"

print(f"Packaging {zip_filename}...")

# Directories / files to exclude
EXCLUDE_DIRS = {
    'node_modules',
    '.git',
    'build',
    '.gradle',
    '.idea',
    '.cxx',
    'captures'
}

EXCLUDE_FILES = {
    zip_filename,
    temp_zip,
    '~$Team_CamGiang_Report.xlsx'
}

with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
    for root, dirs, files in os.walk('.'):
        # Modify dirs in-place to skip excluded directories
        dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS and not d.startswith('.system_generated')]
        
        for file in files:
            if file in EXCLUDE_FILES or file.startswith('~$') or file.endswith('.tmp'):
                continue
            
            # Skip large intermediate xlsb files if needed, or include them?
            # Let's include everything except build artifacts
            filepath = os.path.join(root, file)
            rel_path = os.path.relpath(filepath, '.')
            
            # Skip build folder inside android/app
            if 'app\\build' in rel_path or 'app/build' in rel_path:
                continue

            try:
                zf.write(filepath, rel_path)
            except Exception as e:
                print(f"Skipping {rel_path} due to error: {e}")

print(f"Temporary archive created: {os.path.getsize(temp_zip)} bytes")

# Replace old zip
if os.path.exists(zip_filename):
    os.remove(zip_filename)
os.rename(temp_zip, zip_filename)

print(f"Updated {zip_filename} successfully! Size: {os.path.getsize(zip_filename)} bytes")
