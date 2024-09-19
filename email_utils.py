import json
from pathlib import Path

# Define the path to the "data" directory and the JSON file within it
project_directory = Path(__file__).parent  # Gets the current directory of the script
data_directory = project_directory / "data"
data_directory.mkdir(parents=True, exist_ok=True)  # Create the "data" directory if it doesn't exist
json_file_path = data_directory / "email_info.json"

def save_email_info(email_info):
    """Save the email information to a JSON file."""
    # Load existing data if the file exists
    if json_file_path.exists():
        with open(json_file_path, 'r') as f:
            data = json.load(f)
    else:
        data = []

    # Append new email info
    data.append(email_info)

    # Write updated data to the file
    with open(json_file_path, 'w') as f:
        json.dump(data, f, indent=4)
