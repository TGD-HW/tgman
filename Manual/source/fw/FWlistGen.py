import os
import json

# Define the parameters with simplified names
parameters = {
    "controlHW": ["4", "5"],
    "NoAxis": ["S", "D"],
    "Comm": ["EtherCAT", "Profinet", "CoEcanopen"],
    "FB": ["HiperfaceDSL", "Endat2_2", "BissC_SSI"]
}

# Initialize an empty list to hold the combinations
combinations = []

# Generate all combinations
for controlHW in parameters["controlHW"]:
    for NoAxis in parameters["NoAxis"]:
        for Comm in parameters["Comm"]:
            for FB in parameters["FB"]:
                # Create a dictionary for the current combination
                combination = {
                    "controlHW": controlHW,
                    "NoAxis": NoAxis,
                    "Comm": Comm,
                    "FB": FB,
                    "FW": None  # Placeholder for firmware folder path
                }
                
                # Construct the folder name based on the combination
                folder_name = f"{controlHW}_{NoAxis}_{Comm}_{FB}"
                
                # Create the folder if it doesn't exist
                os.makedirs(folder_name, exist_ok=True)
                
                # Set the FW field to the relative path of the folder
                combination["FW"] = folder_name  # Use relative path
                
                # Add the combination to the list
                combinations.append(combination)

# Convert the list of combinations to JSON
json_output = json.dumps(combinations, indent=2)

# Print the JSON output to the terminal
print(json_output)

# Write the JSON output to a file named FW.json
with open('FW.json', 'w') as json_file:
    json_file.write(json_output)
