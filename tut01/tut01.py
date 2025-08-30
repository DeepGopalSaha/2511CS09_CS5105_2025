#streamlit implementation of logic
#which is in logic.py

import pandas as pd
import csv
import math
import streamlit as st
import io

# ------------------ Streamlit UI ------------------
st.set_page_config(layout="wide")
st.title("Student Grouping App")

# Initialize session state
if 'generated' not in st.session_state:
    st.session_state['generated'] = False
if 'df_clean' not in st.session_state:
    st.session_state['df_clean'] = None
if 'branches' not in st.session_state:
    st.session_state['branches'] = None
if 'n' not in st.session_state:
    st.session_state['n'] = 0

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel file with student details", type=["xlsx"])

# Number of groups
n_input = st.number_input("Enter number of groups:", min_value=1, step=1)

# Generate button
if st.button("Generate"):
    if uploaded_file and n_input > 0:
        df = pd.read_excel(uploaded_file)
        df_clean = df[['Roll', 'Name', 'Email']].copy()
        df_clean['Branch'] = df_clean['Roll'].str[4:6]
        branches = sorted(df_clean['Branch'].unique())

        st.session_state['generated'] = True
        st.session_state['df_clean'] = df_clean
        st.session_state['branches'] = branches
        st.session_state['n'] = n_input
        st.success("Data Generated Successfully!")

# Only show buttons/tables if generated
if st.session_state['generated']:
    df_clean = st.session_state['df_clean']
    branches = st.session_state['branches']
    n = st.session_state['n']
    headers = ['Roll', 'Name', 'Email', 'Branch']

    # Buttons in one row
    col1, col2, col3, col4 = st.columns(4)
    btn_distribution = col1.button("Branchwise Distribution")
    btn_branch = col2.button("Group Branch-wise Mix")
    btn_uniform = col3.button("Group Uniform Mix")
    btn_summary = col4.button("Summary")

    # ------------------ Helper Functions ------------------
    def create_branchwise_groups(df_clean, branches, n):
        branch_files = {b: g.values.tolist() for b, g in df_clean.groupby('Branch')}
        total_students = len(df_clean)
        group_size = math.ceil(total_students / n)
        branch_counters = {b: 0 for b in branches}
        branch_totals = {b: len(students) for b, students in branch_files.items()}

        groups = [[] for _ in range(n)]
        group_index = 0
        while any(branch_counters[b] < branch_totals[b] for b in branches):
            while len(groups[group_index]) < group_size and any(branch_counters[b] < branch_totals[b] for b in branches):
                for b in branches:
                    if branch_counters[b] < branch_totals[b]:
                        student = branch_files[b][branch_counters[b]]
                        groups[group_index].append(student)
                        branch_counters[b] += 1
                        if len(groups[group_index]) >= group_size:
                            break
            group_index += 1
            if group_index >= n:
                break
        return groups

    def create_uniform_groups(df_clean, branches, n):
        branch_files = {b: g.values.tolist() for b, g in df_clean.groupby('Branch')}
        total_students = len(df_clean)
        group_size = math.ceil(total_students / n)
        branch_totals = {b: len(students) for b, students in branch_files.items()}
        branch_counters = {b: 0 for b in branches}

        sorted_branches = sorted(branch_totals.items(), key=lambda x: x[1], reverse=True)
        sorted_branches = [b for b, _ in sorted_branches]

        groups = [[] for _ in range(n)]
        group_index = 0
        for b in sorted_branches:
            while branch_counters[b] < branch_totals[b] and group_index < n:
                remaining_branch = branch_totals[b] - branch_counters[b]
                remaining_group = group_size - len(groups[group_index])
                take = min(remaining_branch, remaining_group)
                students = branch_files[b][branch_counters[b]: branch_counters[b]+take]
                groups[group_index].extend(students)
                branch_counters[b] += take
                if len(groups[group_index]) >= group_size:
                    group_index += 1
        return groups

    def to_csv_download(data, headers):
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(headers)
        writer.writerows(data)
        return output.getvalue()

    # ------------------ Branchwise Distribution ------------------
    if btn_distribution:
        st.subheader("Branchwise Distribution")
        branch_counts = df_clean['Branch'].value_counts().to_dict()
        for b in branches:
            with st.expander(f"{b} Branch ({branch_counts[b]} students)"):
                df_branch = df_clean[df_clean['Branch'] == b].sort_values('Roll').copy()
                df_branch['Name (Branch Total)'] = df_branch.apply(
                    lambda row: f"{row['Name']} ({row['Branch']}-{branch_counts[row['Branch']]})", axis=1
                )
                display_df = df_branch[['Roll', 'Name (Branch Total)', 'Email']].reset_index(drop=True)
                display_df.index += 1  # SL No as index
                display_df.index.name = "Sl No"

                st.dataframe(display_df, use_container_width=True, height=300)

                # --- CSV Download button ---
                csv_data = to_csv_download(display_df.reset_index().values.tolist(), display_df.reset_index().columns)
                st.download_button(
                    label=f"Download {b} Branch CSV",
                    data=csv_data,
                    file_name=f"{b}_branch.csv",
                    mime="text/csv"
                )

    # ------------------ Branch-wise Mix ------------------
    if btn_branch:
        st.subheader("Branch-wise Mix Groups")
        groups = create_branchwise_groups(df_clean, branches, n)
        for i, group in enumerate(groups, start=1):
            with st.expander(f"Group G{i}_mix"):
                df_group = pd.DataFrame(group, columns=headers)
                df_group.insert(0, "Sl No", range(1, len(df_group)+1))
                st.dataframe(df_group, use_container_width=True, height=300)
                st.download_button(
                    label=f"Download G{i}_mix CSV",
                    data=to_csv_download(group, headers),
                    file_name=f"g{i}_mix.csv",
                    mime="text/csv"
                )

    # ------------------ Uniform Mix ------------------
    if btn_uniform:
        st.subheader("Uniform Mix Groups")
        groups = create_uniform_groups(df_clean, branches, n)
        for i, group in enumerate(groups, start=1):
            with st.expander(f"Group G{i}_uniform"):
                df_group = pd.DataFrame(group, columns=headers)
                df_group.insert(0, "Sl No", range(1, len(df_group)+1))
                st.dataframe(df_group, use_container_width=True, height=300)
                st.download_button(
                    label=f"Download G{i}_uniform CSV",
                    data=to_csv_download(group, headers),
                    file_name=f"g{i}_uniform.csv",
                    mime="text/csv"
                )

    # ------------------ Summary ------------------
    if btn_summary:
        st.subheader("Summary")

        # Branch-wise summary
        groups_branch = create_branchwise_groups(df_clean, branches, n)
        summary_rows = []
        for i, group in enumerate(groups_branch, start=1):
            row = {"Mix": f"G{i}"}
            for b in branches:
                row[b] = sum(1 for st in group if st[3] == b)
            row["Total"] = len(group)
            summary_rows.append(row)
        summary_df = pd.DataFrame(summary_rows, columns=["Mix"] + branches + ["Total"])

        # Add total row for branchwise summary
        grand_totals = {"Mix": "Total"}
        for b in branches:
            grand_totals[b] = summary_df[b].sum()
        grand_totals["Total"] = summary_df["Total"].sum()
        summary_df = pd.concat([summary_df, pd.DataFrame([grand_totals])], ignore_index=True)

        # Uniform summary
        groups_uniform = create_uniform_groups(df_clean, branches, n)
        summary_rows_uniform = []
        for i, group in enumerate(groups_uniform, start=1):
            row = {"Uniform": f"G{i}"}
            for b in branches:
                row[b] = sum(1 for st in group if st[3] == b)
            row["Total"] = len(group)
            summary_rows_uniform.append(row)
        summary_uniform_df = pd.DataFrame(summary_rows_uniform, columns=["Uniform"] + branches + ["Total"])

        # ➕ Add total row for uniform summary
        grand_totals_uniform = {"Uniform": "Total"}
        for b in branches:
            grand_totals_uniform[b] = summary_uniform_df[b].sum()
        grand_totals_uniform["Total"] = summary_uniform_df["Total"].sum()
        summary_uniform_df = pd.concat([summary_uniform_df, pd.DataFrame([grand_totals_uniform])], ignore_index=True)

        # Show in UI
        with st.expander("Branch-wise Mix Summary", expanded=True):
            st.dataframe(summary_df, use_container_width=True)
        with st.expander("Uniform Mix Summary", expanded=True):
            st.dataframe(summary_uniform_df, use_container_width=True)


        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Write both tables in the same sheet one after another
            summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=0)
            summary_uniform_df.to_excel(writer, sheet_name="Summary", index=False, startrow=len(summary_df) + 3)

        st.download_button(
            label="Download Summary (XLSX)",
            data=output.getvalue(),
            file_name="summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
