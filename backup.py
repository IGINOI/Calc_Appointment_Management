import shutil
import os
import pandas as pd

# Define the source and destination paths
source_path = 'C:/Users/david/Documents/CENTRO_FIT/Appointment_Schedule_2024.ods'  # Replace with your file's source path
destination_path = 'C:/Users/david/Desktop/Backup_'+str(pd.Timestamp.today().date())+'_Appointment_Schedule_2024.ods'  # Replace with your desired destination path

# Ensure the source file exists
if os.path.exists(source_path):
    try:
        # Copy the file to the new destination
        shutil.copy(source_path, destination_path)
        print(f"File copied successfully to {destination_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
else:
    print(f"Source file does not exist: {source_path}")