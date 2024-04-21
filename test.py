import os
from datetime import datetime

# Get the current date
current_date = datetime.now()

# Define the filename
filename = "example.txt"  # Replace "example.txt" with your actual filename

# Define the directory structure
directory_year_month = current_date.strftime("%B %Y")
directory_day = current_date.strftime("%B %d")

# Create the directories if they don't exist
os.makedirs(os.path.join(directory_year_month, directory_day, "Word Files"), exist_ok=True)
os.makedirs(os.path.join(directory_year_month, directory_day, "TradeMark Removed File"), exist_ok=True)
os.makedirs(os.path.join(directory_year_month, directory_day, "PDF Files"), exist_ok=True)

# Combine the directory paths
folder_path = os.path.join(directory_year_month, directory_day)

# Path to the new file
file_path = os.path.join(folder_path, filename)

# Create the new file
with open(file_path, "w") as file:
    # You can write to the file here if needed
    pass

print(f"New file '{filename}' created in folder '{folder_path}'")
