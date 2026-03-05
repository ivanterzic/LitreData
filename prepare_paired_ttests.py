import pandas as pd
import os
from itertools import combinations

# Read the input file
df = pd.read_excel('input.xlsx')

# Create output directory
output_dir = 'paired_comparisons'
os.makedirs(output_dir, exist_ok=True)

# Get unique techniques
techniques = df['technique'].unique()
print(f"Found {len(techniques)} techniques: {techniques}")

# Generate all possible pairs of techniques
technique_pairs = list(combinations(techniques, 2))
print(f"\nCreating {len(technique_pairs)} paired comparison files...")

# For each pair, create a file with all rows from both techniques
for technique1, technique2 in technique_pairs:
    # Extract data for both techniques
    data1 = df[df['technique'] == technique1].copy()
    data2 = df[df['technique'] == technique2].copy()
    
    # Combine both datasets
    combined = pd.concat([data1, data2], ignore_index=True)
    
    filename = f"{technique1}_vs_{technique2}.xlsx"
    filepath = os.path.join(output_dir, filename)
    
    # Use xlsxwriter engine for cleaner file creation
    with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
        combined.to_excel(writer, sheet_name='Sheet1', index=False)
    
    print(f"Created: {filename} ({len(combined)} rows - {len(data1)} + {len(data2)})")

print(f"\nTotal files created: {len(technique_pairs)}")
print(f"Files saved in '{output_dir}' folder")
print("\nFiles are ready for JASP paired samples t-tests!")
