import pandas as pd
import os

# Create output directory if missing
OUT_DIR = "Output"
os.makedirs(OUT_DIR, exist_ok=True)

# Load input CSV
df = pd.read_csv("Input_File/input_btp_mtp_allocation.csv")

# Extract faculty columns (preferences start after CGPA column)
faculty_cols = df.columns[4:]
num_students, num_faculty = len(df), len(faculty_cols)

# Compute each faculty's base capacity and distribute remainder fairly
base_capacity = num_students // num_faculty
extra_students = num_students % num_faculty
faculty_capacity_global = {fac: base_capacity for fac in faculty_cols}
for fac in faculty_cols[:extra_students]:
    faculty_capacity_global[fac] += 1

# Count how many students ranked each faculty at each preference position
preference_counts = {fac: [0]*num_faculty for fac in faculty_cols}
for _, row in df.iterrows():
    for fac in faculty_cols:
        preference_counts[fac][int(row[fac]) - 1] += 1

# Save preference count summary to CSV
pd.DataFrame(preference_counts).T.reset_index().rename(columns={"index":"Fac"}).to_csv(
    os.path.join(OUT_DIR, "faculty_preference_count.csv"), index=False)

# Sort students descending by CGPA for allocation priority
df_sorted = df.sort_values(by="CGPA", ascending=False).copy()
df_sorted["Allocated"] = ""

# Allocate students in groups of 'num_faculty' (1 student per faculty per group)
for start in range(0, num_students, num_faculty):
    group = df_sorted.iloc[start:start+num_faculty]
    group_available = {fac: (1 if faculty_capacity_global[fac] > 0 else 0) for fac in faculty_cols}

    for idx, student in group.iterrows():
        # Sort each student's faculty preferences by rank (1=top)
        student_prefs = sorted([(fac, int(student[fac])) for fac in faculty_cols], key=lambda x: x[1])
        for fac, _ in student_prefs:
            # Allocate if faculty unused in this group and capacity remains
            if group_available[fac] > 0 and faculty_capacity_global[fac] > 0:
                df_sorted.at[idx, "Allocated"] = fac
                faculty_capacity_global[fac] -= 1
                group_available[fac] = 0
                break

# Save CGPA-wise allocated data (intermediate)
df_sorted.to_csv(os.path.join(OUT_DIR, "cgpa_wise_student_allocation.csv"), index=False)

# Sort final output by roll code (prefix+number)
def roll_sort_key(roll):
    prefix = ''.join([c for c in roll if c.isalpha()]).upper()
    digits = ''.join([c for c in roll if c.isdigit()])
    return (prefix, int(digits) if digits else 0)

df_sorted_final = df_sorted.sort_values(by="Roll", key=lambda x: x.map(roll_sort_key))

# Save final allocation (roll-wise)
df_sorted_final[["Roll", "Name", "Email", "CGPA", "Allocated"]].to_csv(
    os.path.join(OUT_DIR, "roll_wise_student_allocation.csv"), index=False)

print("Allocation complete.")
