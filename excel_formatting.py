import pandas as pd
import re


df = pd.read_csv("DMC - TDS SUM FY 2022-23.csv")

# Define a function to extract Name, PAN, and Address
# Define a function to extract Name, PAN, and Address
def extract_details(party_name):
    if not isinstance(party_name, str):  # Check if party_name is not a string
        return pd.Series([None, None, None])  # Return None values if it's not a string
    
    # Define a regex pattern for PAN (5 letters, 4 digits, and 1 letter)
    pan_pattern = r"([A-Z]{5}[0-9]{4}[A-Z])"
    match = re.search(pan_pattern, party_name)

    if match:
        # Extract Name, PAN, and Address
        pan = match.group(1)
        name = party_name[:match.start()].strip()
        address = party_name[match.end():].strip()
        return pd.Series([name, pan, address])
    else:
        return pd.Series([None, None, None])

# Apply the function to each row in the DataFrame's 'Party Name' column
df[['Name', 'PAN', 'Address']] = df['Party Name'].apply(extract_details)

# Drop rows where 'Name', 'PAN', or 'Address' could not be extracted (optional)
df = df.dropna(subset=['Name', 'PAN', 'Address'])

# Display the result
print(df[['Name', 'PAN', 'Address']])

# Save the modified DataFrame to a new CSV if needed
df.to_csv("extracted_details.csv", index=False)
