import os
import pandas as pd
import re

# Folder and output file
import os

# Folder where .log files are
input_folder = r"D:\STUDY\CMS\SM_ML\SF-Calculation\Python Calculation"

# Folder where you want to save the Excel fileesults
output_folder = r"D:\STUDY\CMS\SM_ML\SF-Calculation\Python Calculation"

# Full path of the Excel file
output_file = os.path.join(output_folder, "Excited_State_Energies.xlsx")

# Data list
data = []

for filename in os.listdir(input_folder):
    if filename.endswith(".log"):
        filepath = os.path.join(input_folder, filename)
    

        file_data = {"File": filename}
        found_triplet = False
        found_singlet = False
        triplet_energies = []
        homo_energy = None
        lumo_energy = None

        try:
            with open(filepath, 'r') as file:
                lines = file.readlines()
            
            for line in lines:
                if line.strip().startswith("Excited State"):
                    # Regex to extract state type and energy
                    match = re.search(r'Excited State\s+\d+:\s+(\w+)-\w+\s+([\d.]+)\s+eV', line)
                    if match:
                        state_type = match.group(1)  # Singlet or Triplet
                        energy = float(match.group(2))

                        if state_type == "Singlet" and not found_singlet:
                            file_data["Singlet"] = energy
                            found_singlet = True
                            print(f"  → Singlet: {energy} eV")
                        elif state_type == "Triplet" and not found_triplet:
                            file_data["Triplet"] = energy
                            found_triplet = True
                            print(f"  → Triplet: {energy} eV")

                if found_singlet and found_triplet:
                    break
         # Add second triplet if available
            if len(triplet_energies) > 1:
                file_data["Triplet2"] = triplet_energies[1]
                
            for i, line in enumerate(lines):
                    if "Alpha  occ. eigenvalues" in line:
                        parts = line.strip().split("--")
                        if len(parts) > 1:
                            occ_values = parts[1].strip().split()
                            if occ_values:
                                homo_energy = float(occ_values[-1])
                                file_data["HOMO"] = homo_energy

                    if "Alpha  virt. eigenvalues" in line:
                        parts = line.strip().split("--")
                        if len(parts) > 1:
                            virt_values = parts[1].strip().split()
                            if virt_values:
                                lumo_energy = float(virt_values[0])
                                file_data["LUMO"] = lumo_energy
            if "Singlet" in file_data and "Triplet" in file_data:
                            data.append(file_data)
        except Exception as e:
            print(f"⚠️ Error in {filename}: {e}")

       

# Create DataFrame and calculate energy gap
df = pd.DataFrame(data)

if "Singlet" in df.columns and "Triplet" in df.columns:
    df["2×T1 - S1 Gap (eV)"] =  2 * df["Triplet"] - df["Singlet"] 

# Save to Excel
df.to_excel(output_file, index=False)
print(f"\n✅ Data saved to {output_file}")

