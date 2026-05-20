import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from openpyxl import Workbook
import io

st.set_page_config(page_title="Student Data & Analysis App", page_icon="🎓")
st.title("Student Data Management & Analysis App")

# ── Column definitions matching the actual Excel ──────────────────────────────
COLUMNS = [
    "School Name", "Student Name", "Class", "Date of Birth",
    "Gender", "Tribe", "Village", "District", "State",
    "Donor 1", "Donor 2", "Donor 3", "Fees Category"
]

CLASS_OPTIONS = ["Nursery", "LKG", "UKG", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]
GENDER_OPTIONS = ["Female", "Male", "Other"]

file_path = 'student_data.xlsx'

# ── Create file if it doesn't exist ──────────────────────────────────────────
if not os.path.exists(file_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Unknown"
    ws.append(COLUMNS[1:]) # School name becomes the sheet name
    wb.save(file_path)

# ── Helper: load & save data ─────────────────────────────────────────────────
def load_data():
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        all_dfs = []
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            if df.empty and len(xls.sheet_names) > 1:
                continue
            
            # Strip whitespace from column names
            df.columns = [str(c).strip() for c in df.columns]
            
            # Rename Excel raw column names to standardized COLUMNS
            mapping = {
                'SCHOOL NAME': 'School Name',
                'NAME': 'Student Name',
                'CLASS': 'Class',
                'DOB': 'Date of Birth',
                'GENDER': 'Gender',
                'IF ST, NAME THE TRIBE': 'Tribe',
                'VILLAGE/TOWN': 'Village',
                'DISTRICT': 'District',
                'STATE': 'State',
                '1 DONOR MAPPED': 'Donor 1',
                '2 DONOR MAPPED': 'Donor 2',
                '3 DONOR MAPPED': 'Donor 3',
                'SPONSORSHIP': 'Fees Category'
            }
            df = df.rename(columns=mapping)
            
            # Use sheet name as school name if column doesn't exist
            if "School Name" not in df.columns:
                df.insert(0, "School Name", sheet)
                
            # Remove duplicate columns that cause concat to fail
            df = df.loc[:, ~df.columns.duplicated()]
            
            all_dfs.append(df)
            
        if all_dfs:
            df_combined = pd.concat(all_dfs, ignore_index=True)
        else:
            df_combined = pd.DataFrame(columns=COLUMNS)
            
    except Exception as e:
        import traceback
        traceback.print_exc()
        import streamlit as st
        st.error(f"Error loading data: {e}")
        df_combined = pd.DataFrame(columns=COLUMNS)
        
    # Ensure all expected columns exist
    for col in COLUMNS:
        if col not in df_combined.columns:
            df_combined[col] = ""
            
    # Force Class and Date of Birth to string to avoid mixed type Arrow serialization errors
    df_combined['Class'] = df_combined['Class'].astype(str).str.strip().replace('nan', '')
    df_combined['Date of Birth'] = df_combined['Date of Birth'].astype(str).replace('nan', '')
    
    # Normalize the Fees Category
    if 'Fees Category' in df_combined.columns:
        def normalize_fees(val):
            if pd.isna(val):
                return "Unspecified"
            val_str = str(val).strip().upper()
            if "TUITION" in val_str:
                return "Tuition Fees"
            elif "HOSTEL" in val_str:
                return "Hostel Fees"
            elif "NUTRITION" in val_str:
                return "Nutrition Fees"
            elif val_str in ["", "NAN", "NONE"]:
                return "Unspecified"
            else:
                return str(val).strip().title()
        df_combined['Fees Category'] = df_combined['Fees Category'].apply(normalize_fees)
    
    return df_combined[COLUMNS]

def save_data(df, path_or_buf=file_path):
    # Save back to multiple sheets based on 'School Name'
    with pd.ExcelWriter(path_or_buf, engine='openpyxl') as writer:
        if "School Name" in df.columns and not df.empty:
            for school, group in df.groupby("School Name"):
                sheet_name = str(school)[:31] if pd.notna(school) and str(school).strip() else "Unknown"
                # Remove invalid characters for Excel sheet names
                for char in r'[]:*?/\'':
                    sheet_name = sheet_name.replace(char, '')
                sheet_name = sheet_name or "Unknown"
                
                # Save without the School Name column to preserve original format
                group_to_save = group.drop(columns=["School Name"], errors='ignore')
                group_to_save.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        else:
            df.to_excel(writer, sheet_name="Sheet1", index=False)


# ── Upload Section ────────────────────────────────────────────────────────────
st.subheader("⬆️ Upload Student Excel File (Optional)")
uploaded_file = st.file_uploader("Choose an Excel file to upload", type=["xlsx"])

if uploaded_file is not None:
    try:
        uploaded_df = pd.read_excel(uploaded_file, engine='openpyxl')
        uploaded_df.columns = [str(c).strip() for c in uploaded_df.columns]
        if "School Name" not in uploaded_df.columns:
            uploaded_df["School Name"] = "Uploaded Data"
            
        main_df = load_data()
        updated_df = pd.concat([main_df, uploaded_df], ignore_index=True)
        save_data(updated_df)
        st.success("✅ Uploaded data added successfully!")
    except Exception as e:
        st.error(f"❌ Error processing uploaded file: {e}")

# ── Load Data ─────────────────────────────────────────────────────────────────
df = load_data()

# ── Executive Dashboard ───────────────────────────────────────────────────────
st.header("📊 Executive Dashboard")

# High-level KPI Metrics
kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5, kpi_col6 = st.columns(6)
total_schools = df['School Name'].nunique()
total_students = len(df)
male_students = int((df['Gender'].str.strip() == 'Male').sum())
female_students = int((df['Gender'].str.strip() == 'Female').sum())
unique_tribes = df['Tribe'].nunique()

# Count unique non-empty active donors
all_donors = pd.concat([df['Donor 1'], df['Donor 2'], df['Donor 3']]).dropna().astype(str).str.strip().replace(['nan', 'None', ''], pd.NA).dropna()
active_donors = all_donors.nunique()

with kpi_col1:
    st.metric(label="🏫 Total Schools", value=total_schools)
with kpi_col2:
    st.metric(label="👦 Total Students", value=f"{total_students:,}")
with kpi_col3:
    st.metric(label="👨 Male Students", value=f"{male_students:,}")
with kpi_col4:
    st.metric(label="👩 Female Students", value=f"{female_students:,}")
with kpi_col5:
    st.metric(label="🏕️ Unique Tribes", value=unique_tribes)
with kpi_col6:
    st.metric(label="🤝 Active Donors", value=active_donors)

st.markdown("---")
st.subheader("📍 Detailed Student Breakdown")

dash_col1, dash_col2, dash_col3, dash_col4 = st.columns(4)

with dash_col1:
    st.subheader("🗺️ District Wise")
    district_count = df['District'].value_counts()
    for dist, count in district_count.items():
        if str(dist).strip():
            st.write(f"**{dist}**: {count:,}")

with dash_col2:
    st.subheader("📍 State Wise")
    state_count = df['State'].value_counts()
    for state, count in state_count.items():
        if str(state).strip():
            st.write(f"**{state}**: {count:,}")

with dash_col3:
    st.subheader("🏕️ Tribe Wise")
    tribe_count = df['Tribe'].value_counts()
    for tribe, count in tribe_count.items():
        if str(tribe).strip():
            st.write(f"**{tribe}**: {count:,}")

with dash_col4:
    st.subheader("💰 Fees Category Wise")
    fees_count = df['Fees Category'].value_counts()
    for cat, count in fees_count.items():
        if str(cat).strip():
            st.write(f"**{cat}**: {count:,}")

st.markdown("---")

# ── Show / Download Data ──────────────────────────────────────────────────────
col_show, col_dl = st.columns(2)
with col_show:
    if st.button("📂 Show All Student Data"):
        st.subheader("📜 Student Records")
        st.dataframe(df, use_container_width=True)
with col_dl:
    buf = io.BytesIO()
    save_data(df, buf)
    st.download_button(
        label="⬇️ Download Excel",
        data=buf.getvalue(),
        file_name="student_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ── Filter Section ────────────────────────────────────────────────────────────
st.subheader("🔎 Filter Data")
col1, col2, col3, col4, col5, col6 = st.columns(6)
with col1:
    school_filter = st.multiselect("🏢 School", sorted(df['School Name'].dropna().unique().astype(str).tolist()))
with col2:
    class_filter = st.multiselect("🏫 Class", sorted(df['Class'].dropna().unique().astype(str).tolist()))
with col3:
    gender_filter = st.multiselect("⚧ Gender", sorted(df['Gender'].dropna().unique().astype(str).tolist()))
with col4:
    district_filter = st.multiselect("🗺️ District", sorted(df['District'].dropna().unique().astype(str).tolist()))
with col5:
    tribe_filter = st.multiselect("🏕️ Tribe", sorted(df['Tribe'].dropna().unique().astype(str).tolist()))
with col6:
    fees_filter = st.multiselect("💰 Fees Category", sorted(df['Fees Category'].dropna().unique().astype(str).tolist()))

filtered_data = df.copy()
if school_filter:
    filtered_data = filtered_data[filtered_data['School Name'].isin(school_filter)]
if class_filter:
    filtered_data = filtered_data[filtered_data['Class'].astype(str).isin([str(c) for c in class_filter])]
if gender_filter:
    filtered_data = filtered_data[filtered_data['Gender'].isin(gender_filter)]
if district_filter:
    filtered_data = filtered_data[filtered_data['District'].isin(district_filter)]
if tribe_filter:
    filtered_data = filtered_data[filtered_data['Tribe'].isin(tribe_filter)]
if fees_filter:
    filtered_data = filtered_data[filtered_data['Fees Category'].isin(fees_filter)]

st.write(f"**Showing {len(filtered_data)} of {len(df)} students**")
st.dataframe(filtered_data, use_container_width=True)

# ── Analysis & Visualisation ──────────────────────────────────────────────────
st.subheader("📈 Data Analysis & Visualization")

# Create tabs for the different analysis views
t1, t2, t3, t4, t5, t6, t7, t8, t9 = st.tabs([
    "🏫 Classes", "🏢 Schools", "⚧ Gender", "🏕️ Tribes", "🗺️ Districts", "📍 States", "🤝 Donors", "💰 Fees", "📊 Stats"
])

with t1:
    class_count = filtered_data['Class'].value_counts().reset_index()
    class_count.columns = ['Class', 'Count']
    class_count['Class'] = class_count['Class'].astype(str)
    class_count = class_count.sort_values('Class', key=lambda x: x.map(
        {c: i for i, c in enumerate(CLASS_OPTIONS)}
    ).fillna(99))
    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(data=class_count, x='Class', y='Count', palette='viridis', ax=ax)
    ax.set_title("📊 Number of Students per Class")
    ax.set_xlabel("Class")
    ax.set_ylabel("No. of Students")
    st.pyplot(fig)

with t2:
    school_count = filtered_data['School Name'].value_counts().reset_index()
    school_count.columns = ['School Name', 'Count']
    fig, ax = plt.subplots(figsize=(12, 6))
    sns.barplot(data=school_count, x='School Name', y='Count', palette='plasma', ax=ax)
    ax.set_title("📊 Students by School")
    plt.xticks(rotation=45, ha='right')
    st.pyplot(fig)

with t3:
    gender_count = filtered_data['Gender'].value_counts()
    fig, ax = plt.subplots()
    ax.pie(gender_count, labels=gender_count.index, autopct='%1.1f%%',
           startangle=90, colors=['#FF69B4', '#4169E1', '#90EE90'])
    ax.set_title("Gender Distribution")
    st.pyplot(fig)

with t4:
    tribe_count = filtered_data['Tribe'].value_counts().reset_index()
    tribe_count.columns = ['Tribe', 'Count']
    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(data=tribe_count, x='Tribe', y='Count', palette='magma', ax=ax)
    ax.set_title("📊 Students by Tribe")
    plt.xticks(rotation=45, ha='right')
    st.pyplot(fig)

with t5:
    dist_count = filtered_data['District'].value_counts().reset_index()
    dist_count.columns = ['District', 'Count']
    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(data=dist_count, x='District', y='Count', palette='coolwarm', ax=ax)
    ax.set_title("📊 Students by District")
    plt.xticks(rotation=45, ha='right')
    st.pyplot(fig)

with t6:
    state_count = filtered_data['State'].value_counts().reset_index()
    state_count.columns = ['State', 'Count']
    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(data=state_count, x='State', y='Count', palette='Set2', ax=ax)
    ax.set_title("📊 Students by State")
    plt.xticks(rotation=45, ha='right')
    st.pyplot(fig)

with t7:
    st.subheader("🤝 Donor Sponsorship Analysis")
    
    # Identify donor columns
    donor_cols = [c for c in ['Donor 1', 'Donor 2', 'Donor 3'] if c in filtered_data.columns]
    
    if donor_cols:
        # Robust check for active sponsorship (excluding NaN, None, empty string)
        sponsored_series = filtered_data[donor_cols].apply(lambda x: x.astype(str).str.strip().replace(['nan', 'None', ''], pd.NA)).notna().any(axis=1)
        sponsored_count = sponsored_series.sum()
        unsponsored_count = len(filtered_data) - sponsored_count
        
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Sponsored Students", f"{sponsored_count:,}", f"{(sponsored_count/len(filtered_data))*100:.1f}% of total")
        with c2:
            st.metric("Unsponsored Students", f"{unsponsored_count:,}", f"{(unsponsored_count/len(filtered_data))*100:.1f}% of total")
        
        st.markdown("---")
        
        # Gather all mappings of donors
        donor_list = filtered_data[donor_cols].apply(lambda x: x.astype(str).str.strip().replace(['nan', 'None', ''], pd.NA)).stack().dropna()
        
        if not donor_list.empty:
            donor_counts = donor_list.value_counts().reset_index()
            donor_counts.columns = ['Donor Name', 'Sponsored Students']
            
            st.markdown("#### 🏆 Top Donors by Sponsorship Count")
            fig, ax = plt.subplots(figsize=(10, 5))
            sns.barplot(data=donor_counts.head(15), y='Donor Name', x='Sponsored Students', palette='autumn', ax=ax)
            ax.set_title("Top 15 Donors (Count of Students Supported)")
            st.pyplot(fig)
            
            st.markdown("#### 📜 All Mapped Donors & Sponsored Student Counts")
            st.dataframe(donor_counts, use_container_width=True)
        else:
            st.info("No active donors mapped in this filtered subset.")
    else:
        st.warning("Donor columns not found in loaded data.")

with t8:
    st.subheader("💰 Fees Category Distribution")
    
    fees_dist = filtered_data['Fees Category'].value_counts().reset_index()
    fees_dist.columns = ['Fees Category', 'Count']
    
    c1, c2 = st.columns([2, 3])
    with c1:
        st.write("**Fees Category Counts:**")
        st.dataframe(fees_dist, use_container_width=True)
    
    with c2:
        fig, ax = plt.subplots(figsize=(8, 4))
        sns.barplot(data=fees_dist, x='Fees Category', y='Count', palette='cool', ax=ax)
        ax.set_title("Number of Students by Fees Category")
        ax.set_ylabel("Count")
        ax.set_xlabel("Fees Category")
        plt.xticks(rotation=15)
        st.pyplot(fig)

with t9:
    st.metric("Total Students", len(filtered_data))
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Female", int((filtered_data['Gender'].str.strip() == 'Female').sum()))
    with c2:
        st.metric("Male", int((filtered_data['Gender'].str.strip() == 'Male').sum()))
    with c3:
        st.metric("Unique Tribes", filtered_data['Tribe'].nunique())
    st.write("**Class-wise Breakdown:**")
    st.dataframe(
        filtered_data.groupby('Class').size().reset_index(name='Count').astype(str),
        use_container_width=True
    )

st.markdown("---")
