import os
import sys
import site
import pandas as pd

def get_package_size(package_dir):
    """Calculates the total size of the files in a package directory."""
    total_size = 0
    # Walk through the directory and sum up the sizes of the relevant files
    for dirpath, dirnames, filenames in os.walk(package_dir):
        # Exclude unnecessary directories
        if '__pycache__' in dirnames:
            dirnames.remove('__pycache__')  # Don't traverse into __pycache__

        # Calculate size of files in the package directory
        for filename in filenames:
            # Include Python files and any compiled extensions
            if filename.endswith('.py') or filename.endswith(('.so', '.pyd', '.dll')):
                filepath = os.path.join(dirpath, filename)
                total_size += os.path.getsize(filepath)
    
    return total_size

def list_all_packages_and_sizes():
    """List all installed packages and their sizes in MB and output to Excel."""
    # Get all directories where Python packages are installed
    site_packages_dirs = site.getsitepackages()
    user_site = site.getusersitepackages()
    
    # Ensure user site packages is included
    if user_site not in site_packages_dirs:
        site_packages_dirs.append(user_site)
    
    # List to hold package data for DataFrame
    packages_data = []

    # Loop through all site-packages directories
    for site_dir in site_packages_dirs:
        print(f"Scanning directory: {site_dir}")
        for root, dirs, files in os.walk(site_dir):
            # Skip certain directories we don't want to scan
            if 'dist-info' in root or 'egg-info' in root or 'easy-install.pth' in files:
                continue

            # Check if this is a package directory (i.e., it contains Python files)
            if any(f.endswith('.py') for f in files):
                package_name = os.path.basename(root)
                
                # Get the size of the package directory
                package_size = get_package_size(root)
                if package_size == 0:
                    continue
                
                package_size_mb = package_size / (1024 ** 2)  # Convert size to MB
                packages_data.append([package_name, package_size_mb])

    # Create a DataFrame with the collected data
    df = pd.DataFrame(packages_data, columns=['Package Name', 'Size (MB)'])
    
    # Write the DataFrame to an Excel file
    excel_filename = "installed_packages_sizes.xlsx"
    df.to_excel(excel_filename, index=False, engine='openpyxl')
    print(f"Excel file saved as: {excel_filename}")

if __name__ == "__main__":
    list_all_packages_and_sizes()
