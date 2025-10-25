import streamlit as st
import pandas as pd

st.set_page_config(page_title="Student-Faculty Allocation", layout="wide")
st.title("Student-Faculty Allocation System")

# ----------------------------
# File upload
# ----------------------------
uploaded_file = st.file_uploader("Upload student CSV file", type=["csv"])

if uploaded_file is not None:
    # Read uploaded CSV
    df = pd.read_csv(uploaded_file)

    st.subheader("Uploaded Student Data")
    st.dataframe(df)

    # Extract faculty columns
    faculty_cols = df.columns[4:]
    num_students, num_faculty = len(df), len(faculty_cols)

    # ----------------------------
    # Compute faculty capacities
    # ----------------------------
    base_capacity = num_students // num_faculty
    extra_students = num_students % num_faculty
    faculty_capacity_global = {fac: base_capacity for fac in faculty_cols}
    for fac in faculty_cols[:extra_students]:
        faculty_capacity_global[fac] += 1

    # ----------------------------
    # Faculty Preference Count
    # ----------------------------
    preference_counts = {fac: [0]*num_faculty for fac in faculty_cols}
    for _, row in df.iterrows():
        for fac in faculty_cols:
            preference_counts[fac][int(row[fac]) - 1] += 1

    pref_count_df = pd.DataFrame(preference_counts).T
    pref_count_df.columns = [f"Count Pref {i+1}" for i in range(num_faculty)]
    pref_count_df.index.name = "Fac"
    pref_count_df.reset_index(inplace=True)

    st.subheader("Faculty Preference Count")
    st.dataframe(pref_count_df)
    st.download_button(
        label="Download Faculty Preference Count CSV",
        data=pref_count_df.to_csv(index=False).encode('utf-8'),
        file_name="faculty_preference_count.csv",
        mime="text/csv"
    )

    # ----------------------------
    # Student Allocation
    # ----------------------------
    df_sorted = df.sort_values(by="CGPA", ascending=False).copy()
    df_sorted["Allocated"] = ""

    for start in range(0, num_students, num_faculty):
        group = df_sorted.iloc[start:start+num_faculty]
        group_available = {fac: (1 if faculty_capacity_global[fac] > 0 else 0) for fac in faculty_cols}

        for idx, student in group.iterrows():
            student_prefs = sorted([(fac, int(student[fac])) for fac in faculty_cols], key=lambda x: x[1])
            for fac, _ in student_prefs:
                if group_available[fac] > 0 and faculty_capacity_global[fac] > 0:
                    df_sorted.at[idx, "Allocated"] = fac
                    faculty_capacity_global[fac] -= 1
                    group_available[fac] = 0
                    break

    # ----------------------------
    # Show CGPA-wise allocation
    # ----------------------------
    st.subheader("CGPA-wise Student Allocation")
    st.dataframe(df_sorted)
    st.download_button(
        label="Download CGPA-wise Allocation CSV",
        data=df_sorted.to_csv(index=False).encode('utf-8'),
        file_name="cgpa_wise_student_allocation.csv",
        mime="text/csv"
    )

    # ----------------------------
    # Roll-wise sorting
    # ----------------------------
    def roll_sort_key(roll):
        prefix = ''.join([c for c in roll if c.isalpha()]).upper()
        digits = ''.join([c for c in roll if c.isdigit()])
        return (prefix, int(digits) if digits else 0)

    df_sorted_final = df_sorted.sort_values(by="Roll", key=lambda x: x.map(roll_sort_key))

    st.subheader("Roll-wise Student Allocation")
    st.dataframe(df_sorted_final[["Roll","Name","Email","CGPA","Allocated"]])
    st.download_button(
        label="Download Roll-wise Allocation CSV",
        data=df_sorted_final[["Roll","Name","Email","CGPA","Allocated"]].to_csv(index=False).encode('utf-8'),
        file_name="roll_wise_student_allocation.csv",
        mime="text/csv"
    )
