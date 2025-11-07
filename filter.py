import pandas as pd
import re

# Load file
df = pd.read_excel("1_CTF-S_All_Extracted_with_sector_text.xlsx",
                   sheet_name="Sheet1")

# Fields to search
sector_cols = [f"Sector_text{i}" for i in range(1, 6)]
subsector_cols = [f"Subsector_text{i}" for i in range(1, 6)]
title_col = "Title of the project programme, activity or other"

# Strict transport-related regex patterns
patterns = [
    r'transport',                        # transport, transportation
    r'\bpublic transport\b',
    r'\bmobility\b',
    r'\broad(s)?\b',                     # road / roads
    r'\brail(way|road|s)?\b',            # rail / railway / railroads
    r'\bport(s)?\b',                     # port (NOT support)
    r'\bseaport(s)?\b',
    r'\bairport(s)?\b',
    r'\bmaritime\b',
    r'\bshipping\b',
    r'\baviation\b',
    r'\btraffic\b',
    r'\bbus(es)?\b',
    r'\bvehicle(s)?\b',
    r'\bfreight\b',
    r'\blogistic(s)?\b'
]

# Compile regex objects
compiled = [re.compile(p, re.IGNORECASE) for p in patterns]

# Function to detect transport-related entries
def is_transport_row(row):
    for col in sector_cols + subsector_cols + [title_col]:
        text = str(row.get(col, "")).strip()
        if not text:
            continue
        for rgx in compiled:
            if rgx.search(text):
                return True
    return False

# Apply detection
df_transport = df[df.apply(is_transport_row, axis=1)].copy()

# Save output if needed
output_path = "CTF_FTC_transport_related_STRICT.xlsx"
df_transport.to_excel(output_path, index=False)

output_path
