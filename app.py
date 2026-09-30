import streamlit as st
import pandas as pd
import os
import io


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="EduTrack",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# PAGE STYLE
# ============================================================

st.markdown("""
<style>

.main .block-container {
    max-width: 95%;
    padding-top: 2rem;
}

div.stButton > button {
    width: 100%;
    border-radius: 6px;
}

h1, h2, h3 {
    font-weight: 600;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# SESSION STATE
# ============================================================

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "active_view" not in st.session_state:
    st.session_state.active_view = "dashboard"

if "file_signature_1" not in st.session_state:
    st.session_state.file_signature_1 = None

if "file_signature_2" not in st.session_state:
    st.session_state.file_signature_2 = None


# ============================================================
# DATABASE FILE NAMES
# ============================================================

FILE_1 = "file1.xlsx"
FILE_2 = "file2.xlsx"


# ============================================================
# STANDARD COLUMNS
# ============================================================

DEFAULT_COLUMNS = [
    "StudentID",
    "Student Name",
    "Email Address",
    "Registered Course",
    "Current Status",
    "Internal Mentor Remarks"
]


# ============================================================
# CREATE EMPTY DATAFRAME
# ============================================================

def create_empty_dataframe():

    return pd.DataFrame(
        columns=DEFAULT_COLUMNS
    )


# ============================================================
# CLEAN DATAFRAME
# ============================================================

def clean_dataframe(df):

    if df is None:
        return create_empty_dataframe()

    if not isinstance(df, pd.DataFrame):
        return create_empty_dataframe()

    if df.empty:
        return create_empty_dataframe()

    df = df.copy()

    # Remove completely empty rows and columns
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    if df.empty:
        return create_empty_dataframe()

    # --------------------------------------------------------
    # Flexible column mapping
    # --------------------------------------------------------

    rename_dict = {}

    for col in df.columns:

        col_clean = (
            str(col)
            .strip()
            .lower()
            .replace("_", " ")
            .replace("-", " ")
        )

        # Student ID
        if any(term in col_clean for term in [
            "student id",
            "studentid",
            "student number",
            "student no",
            "prn",
            "roll no",
            "roll number",
            "registration no",
            "registration number",
            "enrollment",
            "enrolment"
        ]):

            if "StudentID" not in rename_dict.values():
                rename_dict[col] = "StudentID"

        # Student Name
        elif any(term in col_clean for term in [
            "student name",
            "studentname",
            "candidate name",
            "full name",
            "name of student",
            "name"
        ]):

            if "Student Name" not in rename_dict.values():
                rename_dict[col] = "Student Name"

        # Email
        elif any(term in col_clean for term in [
            "email",
            "email address",
            "mail id",
            "mail"
        ]):

            if "Email Address" not in rename_dict.values():
                rename_dict[col] = "Email Address"

        # Course / Branch
        elif any(term in col_clean for term in [
            "course",
            "branch",
            "department",
            "dept",
            "program",
            "programme",
            "stream"
        ]):

            if "Registered Course" not in rename_dict.values():
                rename_dict[col] = "Registered Course"

        # Status
        elif "status" in col_clean:

            if "Current Status" not in rename_dict.values():
                rename_dict[col] = "Current Status"

        # Remarks
        elif any(term in col_clean for term in [
            "remark",
            "remarks",
            "note",
            "notes",
            "comment",
            "comments",
            "feedback"
        ]):

            if "Internal Mentor Remarks" not in rename_dict.values():
                rename_dict[col] = "Internal Mentor Remarks"

    df = df.rename(
        columns=rename_dict
    )

    # --------------------------------------------------------
    # Fallback Student ID
    # --------------------------------------------------------

    if "StudentID" not in df.columns:

        if len(df.columns) >= 1:

            first_column = df.columns[0]

            df = df.rename(
                columns={
                    first_column: "StudentID"
                }
            )

    # --------------------------------------------------------
    # Fallback Student Name
    # --------------------------------------------------------

    if "Student Name" not in df.columns:

        possible_columns = [
            c
            for c in df.columns
            if c != "StudentID"
        ]

        if possible_columns:

            df = df.rename(
                columns={
                    possible_columns[0]: "Student Name"
                }
            )

    # --------------------------------------------------------
    # Add missing standard columns
    # --------------------------------------------------------

    for column in DEFAULT_COLUMNS:

        if column not in df.columns:

            df[column] = ""

    # --------------------------------------------------------
    # Put standard columns first
    # --------------------------------------------------------

    extra_columns = [
        column
        for column in df.columns
        if column not in DEFAULT_COLUMNS
    ]

    df = df[
        DEFAULT_COLUMNS + extra_columns
    ]

    # --------------------------------------------------------
    # Clean Student ID
    # --------------------------------------------------------

    df["StudentID"] = (
        df["StudentID"]
        .fillna("")
        .astype(str)
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
        .str.strip()
    )

    # --------------------------------------------------------
    # Clean other standard columns
    # --------------------------------------------------------

    for column in DEFAULT_COLUMNS:

        if column != "StudentID":

            df[column] = (
                df[column]
                .fillna("")
                .astype(str)
                .str.strip()
            )

    # --------------------------------------------------------
    # Remove rows where both ID and Name are empty
    # --------------------------------------------------------

    df = df[
        (df["StudentID"] != "")
        |
        (df["Student Name"] != "")
    ]

    return df.reset_index(
        drop=True
    )


# ============================================================
# READ ALL EXCEL SHEETS
# ============================================================

def read_excel_all_sheets(file_source):

    all_data = []

    try:

        excel_file = pd.ExcelFile(
            file_source
        )

        for sheet_name in excel_file.sheet_names:

            try:

                current_sheet = pd.read_excel(
                    excel_file,
                    sheet_name=sheet_name
                )

                if current_sheet is None:
                    continue

                if current_sheet.empty:
                    continue

                cleaned_sheet = clean_dataframe(
                    current_sheet
                )

                if not cleaned_sheet.empty:

                    all_data.append(
                        cleaned_sheet
                    )

            except Exception:

                continue

    except Exception:

        return create_empty_dataframe()

    if not all_data:

        return create_empty_dataframe()

    final_df = pd.concat(
        all_data,
        ignore_index=True
    )

    return clean_dataframe(
        final_df
    )


# ============================================================
# READ UPLOADED FILE
# ============================================================

def read_uploaded_file(uploaded_file):

    if uploaded_file is None:

        return create_empty_dataframe()

    try:

        file_name = (
            uploaded_file.name.lower()
        )

        file_bytes = (
            uploaded_file.getvalue()
        )

        # CSV
        if file_name.endswith(".csv"):

            df = pd.read_csv(
                io.BytesIO(file_bytes)
            )

            return clean_dataframe(
                df
            )

        # Excel
        if (
            file_name.endswith(".xlsx")
            or
            file_name.endswith(".xls")
        ):

            return read_excel_all_sheets(
                io.BytesIO(file_bytes)
            )

        return create_empty_dataframe()

    except Exception as error:

        st.error(
            f"❌ Error reading file: {error}"
        )

        return create_empty_dataframe()


# ============================================================
# READ SAVED DATABASE
# ============================================================

def read_file_safely(file_path):

    if not os.path.exists(file_path):

        return create_empty_dataframe()

    try:

        if file_path.lower().endswith(
            ".csv"
        ):

            df = pd.read_csv(
                file_path
            )

            return clean_dataframe(
                df
            )

        return read_excel_all_sheets(
            file_path
        )

    except Exception:

        return create_empty_dataframe()


# ============================================================
# SAVE DATABASE
# ============================================================

def save_excel(df, file_path):

    try:

        df = clean_dataframe(
            df
        )

        df.to_excel(
            file_path,
            index=False,
            engine="openpyxl"
        )

        return True

    except Exception as error:

        st.error(
            f"❌ Could not save Excel file: {error}"
        )

        return False


# ============================================================
# LOAD BOTH DATABASES
# ============================================================

def load_databases():

    database_1 = read_file_safely(
        FILE_1
    )

    database_2 = read_file_safely(
        FILE_2
    )

    return database_1, database_2


# ============================================================
# LOGIN PAGE
# ============================================================

if not st.session_state.authenticated:

    st.markdown(
        """
        <h1 style="
            text-align:center;
            margin-top:80px;
        ">
        🎓 EduTrack
        </h1>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <p style="text-align:center;">
        🔒 Secure Mentor Access Panel
        </p>
        """,
        unsafe_allow_html=True
    )

    left, middle, right = st.columns(
        [1, 1.5, 1]
    )

    with middle:

        with st.form(
            "login_form"
        ):

            username = st.text_input(
                "Mentor Username ID"
            )

            password = st.text_input(
                "Password",
                type="password"
            )

            login_button = st.form_submit_button(
                "🔐 Enter EduTrack"
            )

            if login_button:

                if (
                    username == "mentor"
                    and password == "1234"
                ):

                    st.session_state.authenticated = True

                    st.session_state.active_view = (
                        "dashboard"
                    )

                    st.success(
                        "✅ Login successful!"
                    )

                    st.rerun()

                else:

                    st.error(
                        "❌ Invalid username or password."
                    )

    st.stop()


# ============================================================
# LOAD DATABASES
# ============================================================

df1, df2 = load_databases()


# ============================================================
# COMBINED DATA FOR SEARCH ONLY
# ============================================================

search_df_1 = df1.copy()
search_df_1["Source"] = "Excel Sheet 1"

search_df_2 = df2.copy()
search_df_2["Source"] = "Excel Sheet 2"

combined_df = pd.concat(
    [
        search_df_1,
        search_df_2
    ],
    ignore_index=True
)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title(
    "🎓 EduTrack"
)

st.sidebar.caption(
    "Mentor Management System"
)

st.sidebar.divider()

st.sidebar.write(
    "**Current Operator:**"
)

st.sidebar.write(
    "Active Mentor Account"
)

st.sidebar.divider()


# ============================================================
# NAVIGATION
# ============================================================

if st.sidebar.button(
    "📋 Master Database View"
):

    st.session_state.active_view = "dashboard"


if st.sidebar.button(
    "🔍 Search & Edit Records"
):

    st.session_state.active_view = "search"


if st.sidebar.button(
    "➕ Register New Student"
):

    st.session_state.active_view = "register"


st.sidebar.divider()


# ============================================================
# LOGOUT
# ============================================================

if st.sidebar.button(
    "🚪 Logout"
):

    st.session_state.authenticated = False

    st.session_state.active_view = "dashboard"

    st.rerun()


# ============================================================
# DASHBOARD
# ============================================================

if st.session_state.active_view == "dashboard":

    st.title(
        "🎓 EduTrack Dashboard"
    )

    st.write(
        "Student monitoring and mentor management system."
    )

    st.divider()


    # ========================================================
    # UPLOAD SECTION
    # ========================================================

    st.subheader(
        "📁 Student Database Upload"
    )

    st.info(
        "Upload your two student Excel databases separately."
    )

    upload_col1, upload_col2 = st.columns(
        2
    )


    # ========================================================
    # DATABASE 1 UPLOAD
    # ========================================================

    with upload_col1:

        st.markdown(
            "### 📄 Student Database 1"
        )

        uploaded_file1 = st.file_uploader(
            "Upload Excel File 1",
            type=[
                "xlsx",
                "xls",
                "csv"
            ],
            key="student_file_1"
        )

        if uploaded_file1 is not None:

            signature_1 = (
                uploaded_file1.name,
                uploaded_file1.size
            )

            if (
                st.session_state.file_signature_1
                != signature_1
            ):

                new_df1 = read_uploaded_file(
                    uploaded_file1
                )

                if new_df1.empty:

                    st.error(
                        "❌ No student records could be read from Database 1."
                    )

                else:

                    if save_excel(
                        new_df1,
                        FILE_1
                    ):

                        st.session_state.file_signature_1 = (
                            signature_1
                        )

                        st.success(
                            f"✅ Database 1 uploaded successfully! "
                            f"{len(new_df1)} records found."
                        )

                        st.rerun()

            else:

                st.success(
                    f"✅ Database 1 loaded — "
                    f"{len(df1)} students"
                )


    # ========================================================
    # DATABASE 2 UPLOAD
    # ========================================================

    with upload_col2:

        st.markdown(
            "### 📄 Student Database 2"
        )

        uploaded_file2 = st.file_uploader(
            "Upload Excel File 2",
            type=[
                "xlsx",
                "xls",
                "csv"
            ],
            key="student_file_2"
        )

        if uploaded_file2 is not None:

            signature_2 = (
                uploaded_file2.name,
                uploaded_file2.size
            )

            if (
                st.session_state.file_signature_2
                != signature_2
            ):

                new_df2 = read_uploaded_file(
                    uploaded_file2
                )

                if new_df2.empty:

                    st.error(
                        "❌ No student records could be read from Database 2."
                    )

                else:

                    if save_excel(
                        new_df2,
                        FILE_2
                    ):

                        st.session_state.file_signature_2 = (
                            signature_2
                        )

                        st.success(
                            f"✅ Database 2 uploaded successfully! "
                            f"{len(new_df2)} records found."
                        )

                        st.rerun()

            else:

                st.success(
                    f"✅ Database 2 loaded — "
                    f"{len(df2)} students"
                )


    st.divider()


    # ========================================================
    # STATISTICS
    # ========================================================

    metric1, metric2, metric3 = st.columns(
        3
    )

    metric1.metric(
        "👨‍🎓 Total Students",
        len(df1) + len(df2)
    )

    metric2.metric(
        "📄 Database 1 Students",
        len(df1)
    )

    metric3.metric(
        "📄 Database 2 Students",
        len(df2)
    )


    st.divider()


    # ========================================================
    # DATABASE 1 - DIRECTLY EDITABLE TABLE
    # ========================================================

    st.subheader(
        "📊 Student Database 1"
    )

    if df1.empty:

        st.info(
            "No student records found in Database 1."
        )

    else:

        st.write(
            f"**Total records: {len(df1)}**"
        )

        edited_df1 = st.data_editor(
            df1,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="database_1_editor",

            column_config={

                "StudentID": st.column_config.TextColumn(
                    "Student ID",
                    required=True
                ),

                "Student Name": st.column_config.TextColumn(
                    "Student Name",
                    required=True
                ),

                "Email Address": st.column_config.TextColumn(
                    "Email Address"
                ),

                "Registered Course": st.column_config.TextColumn(
                    "Registered Course"
                ),

                "Current Status": st.column_config.SelectboxColumn(
                    "Current Status",
                    options=[
                        "Active",
                        "At Risk",
                        "Inactive",
                        "Completed"
                    ]
                ),

                "Internal Mentor Remarks": st.column_config.TextColumn(
                    "Internal Mentor Remarks"
                )
            }
        )

        st.caption(
            "✏️ Click any cell to edit it. "
            "You can also add or delete rows."
        )

        if st.button(
            "💾 Save Database 1 Changes",
            key="save_database_1"
        ):

            if save_excel(
                edited_df1,
                FILE_1
            ):

                st.success(
                    "✅ Database 1 changes saved successfully!"
                )

                st.rerun()


    st.divider()


    # ========================================================
    # DATABASE 2 - DIRECTLY EDITABLE TABLE
    # ========================================================

    st.subheader(
        "📊 Student Database 2"
    )

    if df2.empty:

        st.info(
            "No student records found in Database 2."
        )

    else:

        st.write(
            f"**Total records: {len(df2)}**"
        )

        edited_df2 = st.data_editor(
            df2,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="database_2_editor",

            column_config={

                "StudentID": st.column_config.TextColumn(
                    "Student ID",
                    required=True
                ),

                "Student Name": st.column_config.TextColumn(
                    "Student Name",
                    required=True
                ),

                "Email Address": st.column_config.TextColumn(
                    "Email Address"
                ),

                "Registered Course": st.column_config.TextColumn(
                    "Registered Course"
                ),

                "Current Status": st.column_config.SelectboxColumn(
                    "Current Status",
                    options=[
                        "Active",
                        "At Risk",
                        "Inactive",
                        "Completed"
                    ]
                ),

                "Internal Mentor Remarks": st.column_config.TextColumn(
                    "Internal Mentor Remarks"
                )
            }
        )

        st.caption(
            "✏️ Click any cell to edit it. "
            "You can also add or delete rows."
        )

        if st.button(
            "💾 Save Database 2 Changes",
            key="save_database_2"
        ):

            if save_excel(
                edited_df2,
                FILE_2
            ):

                st.success(
                    "✅ Database 2 changes saved successfully!"
                )

                st.rerun()


# ============================================================
# SEARCH & EDIT PAGE
# ============================================================

elif st.session_state.active_view == "search":

    st.title(
        "🔍 Search & Edit Student Records"
    )

    st.write(
        "Search students across both EduTrack databases."
    )

    st.divider()

    if combined_df.empty:

        st.warning(
            "No student records available."
        )

    else:

        search = st.text_input(
            "🔎 Search Student ID or Student Name",
            placeholder="Example: 101 or Rahul"
        )

        search = search.strip()

        if search:

            id_match = (
                combined_df[
                    "StudentID"
                ]
                .astype(str)
                .str.contains(
                    search,
                    case=False,
                    na=False,
                    regex=False
                )
            )

            name_match = (
                combined_df[
                    "Student Name"
                ]
                .astype(str)
                .str.contains(
                    search,
                    case=False,
                    na=False,
                    regex=False
                )
            )

            matches = combined_df[
                id_match | name_match
            ]

            if matches.empty:

                st.warning(
                    "❌ No student found."
                )

            else:

                st.success(
                    f"✅ {len(matches)} student record(s) found."
                )

                student_options = []

                for _, row in matches.iterrows():

                    option = (
                        str(row["Student Name"])
                        + " | ID: "
                        + str(row["StudentID"])
                        + " | "
                        + str(row["Source"])
                    )

                    student_options.append(
                        option
                    )

                selected_student = st.selectbox(
                    "Select student:",
                    student_options
                )

                selected_index = (
                    student_options.index(
                        selected_student
                    )
                )

                target = matches.iloc[
                    selected_index
                ]

                target_id = str(
                    target["StudentID"]
                )

                target_source = str(
                    target["Source"]
                )

                st.divider()

                st.subheader(
                    "👤 Student Profile"
                )

                st.caption(
                    f"Database: {target_source}"
                )

                with st.form(
                    "edit_student_form"
                ):

                    new_student_id = st.text_input(
                        "Student ID",
                        value=str(
                            target["StudentID"]
                        )
                    )

                    new_name = st.text_input(
                        "Student Name",
                        value=str(
                            target["Student Name"]
                        )
                    )

                    new_email = st.text_input(
                        "Email Address",
                        value=str(
                            target["Email Address"]
                        )
                    )

                    new_course = st.text_input(
                        "Registered Course",
                        value=str(
                            target["Registered Course"]
                        )
                    )

                    status_list = [
                        "Active",
                        "At Risk",
                        "Inactive",
                        "Completed"
                    ]

                    old_status = str(
                        target["Current Status"]
                    )

                    if old_status in status_list:

                        status_position = (
                            status_list.index(
                                old_status
                            )
                        )

                    else:

                        status_position = 0

                    new_status = st.selectbox(
                        "Current Status",
                        status_list,
                        index=status_position
                    )

                    new_remarks = st.text_area(
                        "Internal Mentor Remarks",
                        value=str(
                            target[
                                "Internal Mentor Remarks"
                            ]
                        )
                    )

                    update_button = st.form_submit_button(
                        "💾 Save Changes"
                    )

                    if update_button:

                        if (
                            target_source
                            == "Excel Sheet 1"
                        ):

                            database_file = FILE_1

                        else:

                            database_file = FILE_2

                        database = read_file_safely(
                            database_file
                        )

                        matching_rows = database[
                            database[
                                "StudentID"
                            ]
                            .astype(str)
                            .str.strip()
                            ==
                            target_id.strip()
                        ]

                        if matching_rows.empty:

                            st.error(
                                "❌ Student could not be found "
                                "in the database."
                            )

                        else:

                            original_index = (
                                matching_rows.index[0]
                            )

                            database.loc[
                                original_index,
                                "StudentID"
                            ] = new_student_id

                            database.loc[
                                original_index,
                                "Student Name"
                            ] = new_name

                            database.loc[
                                original_index,
                                "Email Address"
                            ] = new_email

                            database.loc[
                                original_index,
                                "Registered Course"
                            ] = new_course

                            database.loc[
                                original_index,
                                "Current Status"
                            ] = new_status

                            database.loc[
                                original_index,
                                "Internal Mentor Remarks"
                            ] = new_remarks

                            if save_excel(
                                database,
                                database_file
                            ):

                                st.success(
                                    "✅ Student record updated successfully!"
                                )

                                st.rerun()


# ============================================================
# REGISTER NEW STUDENT
# ============================================================

elif st.session_state.active_view == "register":

    st.title(
        "➕ Register New Student"
    )

    st.write(
        "Add a new student to your EduTrack database."
    )

    st.divider()

    with st.form(
        "register_student_form"
    ):

        target_sheet = st.selectbox(
            "Select Target Database",
            [
                "Excel Sheet 1",
                "Excel Sheet 2"
            ]
        )

        reg_id = st.text_input(
            "Student ID"
        )

        reg_name = st.text_input(
            "Student Name"
        )

        reg_email = st.text_input(
            "Email Address"
        )

        reg_course = st.text_input(
            "Registered Course"
        )

        reg_status = st.selectbox(
            "Current Status",
            [
                "Active",
                "At Risk",
                "Inactive",
                "Completed"
            ]
        )

        reg_remarks = st.text_area(
            "Internal Mentor Remarks"
        )

        register_button = st.form_submit_button(
            "➕ Save & Register Student"
        )

        if register_button:

            if not reg_id or not reg_name:

                st.error(
                    "❌ Student ID and Student Name are required."
                )

            else:

                if (
                    target_sheet
                    == "Excel Sheet 1"
                ):

                    file_to_update = FILE_1

                else:

                    file_to_update = FILE_2

                current_db = read_file_safely(
                    file_to_update
                )

                duplicate = current_db[
                    current_db[
                        "StudentID"
                    ]
                    .astype(str)
                    .str.strip()
                    ==
                    str(reg_id).strip()
                ]

                if not duplicate.empty:

                    st.error(
                        "❌ A student with this Student ID "
                        "already exists in this database."
                    )

                else:

                    new_row = pd.DataFrame([
                        {
                            "StudentID": reg_id,
                            "Student Name": reg_name,
                            "Email Address": reg_email,
                            "Registered Course": reg_course,
                            "Current Status": reg_status,
                            "Internal Mentor Remarks": reg_remarks
                        }
                    ])

                    updated_db = pd.concat(
                        [
                            current_db,
                            new_row
                        ],
                        ignore_index=True
                    )

                    if save_excel(
                        updated_db,
                        file_to_update
                    ):

                        st.success(
                            f"✅ Student successfully added "
                            f"to {target_sheet}!"
                        )

                        st.rerun()