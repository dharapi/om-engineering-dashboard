import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os
import json
from PIL import Image
from streamlit_autorefresh import st_autorefresh

# ============================================================
# PAGE CONFIGURATION (with OM Logo as Favicon)
# ============================================================
LOGO_PATH = os.path.join("assets", "logo.jpg")

favicon = "🎓"
if os.path.exists(LOGO_PATH):
    try:
        favicon = Image.open(LOGO_PATH)
    except Exception:
        favicon = "🎓"

st.set_page_config(
    page_title="OM Education - Student Performance System",
    page_icon=favicon,
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CUSTOM CSS
# ============================================================
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .main-header {
        background: linear-gradient(90deg, #0d7377 0%, #14a0a5 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        margin-bottom: 20px;
    }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        border-left: 5px solid #14a0a5;
    }
    .metric-value {
        font-size: 32px;
        font-weight: bold;
        color: #0d7377;
    }
    .metric-label {
        font-size: 14px;
        color: #666;
    }
    .branch-card {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        text-align: center;
        border-top: 4px solid #14a0a5;
        margin-bottom: 15px;
    }
    .branch-card h3 {
        color: #0d7377;
        margin: 0;
    }
    .branch-card .count {
        font-size: 36px;
        font-weight: bold;
        color: #14a0a5;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button[kind="primary"],
    .stFormSubmitButton>button {
        background: linear-gradient(90deg, #0d7377, #14a0a5) !important;
        color: white !important;
        border: none !important;
        font-weight: bold !important;
    }
    .stButton>button[kind="primary"]:hover,
    .stFormSubmitButton>button:hover {
        background: linear-gradient(90deg, #0a5c5f, #0d7377) !important;
        box-shadow: 0 4px 15px rgba(13, 115, 119, 0.4) !important;
        transform: translateY(-1px);
    }
    .stButton>button[kind="secondary"] {
        background: white !important;
        color: #0d7377 !important;
        border: 2px solid #14a0a5 !important;
    }
    .stButton>button[kind="secondary"]:hover {
        background: #f0fdfa !important;
        border-color: #0d7377 !important;
    }
    .creator-footer {
        text-align: center;
        padding: 25px;
        color: #64748b;
        font-size: 14px;
        border-top: 2px solid #14a0a5;
        margin-top: 40px;
        background: linear-gradient(135deg, #f0fdfa, #f8fafc);
        border-radius: 10px;
    }
    .creator-name {
        color: #0d7377;
        font-weight: bold;
        font-size: 18px;
        letter-spacing: 1px;
    }
    .live-clock {
        background: #f8fafc;
        padding: 12px;
        border-radius: 8px;
        border-left: 3px solid #14a0a5;
        font-size: 12px;
        color: #475569;
        line-height: 1.8;
    }
    .live-time {
        color: #0d7377;
        font-weight: bold;
    }
    .live-dot {
        display: inline-block;
        width: 8px;
        height: 8px;
        background: #10b981;
        border-radius: 50%;
        margin-right: 5px;
        animation: pulse 1.5s infinite;
    }
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.3; }
    }
    .stImage img {
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(13, 115, 119, 0.15);
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# DATA MANAGER
# ============================================================
DATA_DIR = "data"
ASSETS_DIR = "assets"
STUDENTS_FILE = os.path.join(DATA_DIR, "students.csv")
RESULTS_FILE = os.path.join(DATA_DIR, "weekly_results.json")
RESULTS_XLSX = "Weekly Test result.xlsx"

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(ASSETS_DIR, exist_ok=True)

BRANCHES = ["Civil Engineering", "Computer Engineering", "Electrical Engineering", "Mechanical Engineering"]
SEMESTER = "Sem-1"
ACADEMIC_YEAR = "2026-27"
PASS_MARKS = 8

# ============================================================
# SUBJECTS
# ============================================================
DEFAULT_SUBJECTS = {
    "Civil Engineering": [
        {"code": "DI01000121", "name": "Applied Physics", "faculty": "Prof. M.L. Lodhiya"},
        {"code": "DI01000021", "name": "Mathematics-I", "faculty": "Prof. M.J. Vasoya"},
        {"code": "DI01000611", "name": "Civil Engineering Drawing", "faculty": "Prof. D.R. Solanki"},
        {"code": "DI01000621", "name": "Construction Material", "faculty": "Prof. D.R. Solanki"},
        {"code": "DI01000111", "name": "Engineering Graphics", "faculty": "Prof. S.J. Chavda"},
        {"code": "DI01000031", "name": "Communication Skills in English", "faculty": "Prof. D.H. Tolia / Prof. P.M. Pithiya"},
        {"code": "DI01000041", "name": "Sports and Yoga", "faculty": "F-1"},
    ],
    "Computer Engineering": [
        {"code": "DI01000031", "name": "Communication Skills in English", "faculty": "Prof. D.H. Tolia / Prof. P.M. Pithiya"},
        {"code": "DI01000061", "name": "Modern Physics", "faculty": "Prof. M.L. Lodhiya"},
        {"code": "DI01000131", "name": "Computer Programming Fundamentals", "faculty": "Prof. D.C. Pithva / Prof. P.I. Mesiya"},
        {"code": "DI01000021", "name": "Mathematics-I", "faculty": "Prof. M.J. Vasoya"},
        {"code": "DI01000151", "name": "Basics of Electronics", "faculty": "Prof. K.K. Bhalani"},
        {"code": "DI01000141", "name": "Computer Basics and Static Web Page Designing", "faculty": "Prof. D.C. Pithva / Prof. P.I. Mesiya"},
        {"code": "DI01000041", "name": "Sports and Yoga", "faculty": "Prof. D.C. Pithva"},
    ],
    "Electrical Engineering": [
        {"code": "DI01000071", "name": "Engineering Chemistry", "faculty": "Prof. Y.J. Ratanpara"},
        {"code": "DI01000031", "name": "Communication Skills in English", "faculty": "Prof. D.H. Tolia / Prof. P.M. Pithiya"},
        {"code": "DI01000021", "name": "Mathematics-I", "faculty": "Prof. M.J. Vasoya"},
        {"code": "DI01000081", "name": "DC Circuit", "faculty": "Prof. B.A. Nandaniya"},
        {"code": "DI01000901", "name": "Electronics Devices & Circuits", "faculty": "Prof. K.K. Bhalani"},
        {"code": "DI01000691", "name": "Fundamentals of Information and Communication Technology", "faculty": "Prof. B.A. Nandaniya"},
        {"code": "DI01000041", "name": "Sports and Yoga", "faculty": "Prof. K.K. Bhalani"},
    ],
    "Mechanical Engineering": [
        {"code": "DI01000021", "name": "Mathematics-I", "faculty": "Prof. M.J. Vasoya"},
        {"code": "DI01000191", "name": "Elements of Mechanical Engineering", "faculty": "Prof. S.J. Chavda"},
        {"code": "DI01000031", "name": "Communication Skills in English", "faculty": "Prof. D.H. Tolia / Prof. P.M. Pithiya"},
        {"code": "DI01000111", "name": "Engineering Graphics", "faculty": "Prof. S.J. Chavda"},
        {"code": "DI01000181", "name": "Engineering Materials Science", "faculty": "Prof. P.M. Parmar"},
        {"code": "DI01000171", "name": "Basics of Electricals and Electronics Engineering", "faculty": "Prof. Y.J. Dodya / Prof. K.K. Bhalani"},
        {"code": "DI01000041", "name": "Sports and Yoga", "faculty": "Prof. S.J. Chavda"},
    ],
}

BRANCH_COORDINATORS = {
    "Civil Engineering": "Prof. D.R. Solanki",
    "Computer Engineering": "Prof. D.C. Pithva",
    "Electrical Engineering": "Prof. K.K. Bhalani",
    "Mechanical Engineering": "Prof. S.J. Chavda",
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================
def get_logo():
    if os.path.exists(LOGO_PATH):
        try:
            return Image.open(LOGO_PATH)
        except Exception:
            return None
    return None


def load_students_from_xls(xls_path):
    try:
        df = pd.read_excel(xls_path, engine='xlrd')
    except Exception:
        try:
            df = pd.read_excel(xls_path)
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            return pd.DataFrame()
    
    df.columns = [str(c).strip() for c in df.columns]
    rename_map = {
        "TEMP. NO.": "Roll No",
        "Branch 2": "Branch",
        "Full Name": "Name",
        "Gender": "Gender",
        "Student Email ID": "Email",
    }
    df = df.rename(columns=rename_map)
    
    required = ["Roll No", "Branch", "Name", "Gender", "Email"]
    for col in required:
        if col not in df.columns:
            df[col] = ""
    
    df = df[required]
    df["Roll No"] = df["Roll No"].astype(str).str.strip()
    df["Branch"] = df["Branch"].astype(str).str.strip()
    df["Name"] = df["Name"].astype(str).str.strip().str.upper()
    df["Gender"] = df["Gender"].astype(str).str.strip()
    df["Email"] = df["Email"].astype(str).str.strip()
    df = df.dropna(subset=["Roll No", "Name"])
    df = df[df["Roll No"] != ""]
    df["Semester"] = SEMESTER
    df["Status"] = "Active"
    df = df.reset_index(drop=True)
    df.insert(0, "Sr", range(1, len(df) + 1))
    return df


def load_students():
    if os.path.exists(STUDENTS_FILE):
        return pd.read_csv(STUDENTS_FILE)
    xls_path = "2026-2027.xls"
    if os.path.exists(xls_path):
        df = load_students_from_xls(xls_path)
        if not df.empty:
            df.to_csv(STUDENTS_FILE, index=False)
            return df
    return pd.DataFrame(columns=["Sr", "Roll No", "Branch", "Name", "Gender", "Email", "Semester", "Status"])


def load_json(path, default):
    if os.path.exists(path):
        with open(path, "r") as f:
            return json.load(f)
    return default


def save_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f, indent=2)


def load_weekly_results():
    if os.path.exists(RESULTS_FILE):
        return load_json(RESULTS_FILE, {})
    
    if not os.path.exists(RESULTS_XLSX):
        return {}
    
    try:
        results = {}
        sheet_map = {
            "Civil": "Civil Engineering",
            "Electrical": "Electrical Engineering",
            "Mechanical": "Mechanical Engineering",
            "Computer": "Computer Engineering",
        }
        
        for sheet_name, branch in sheet_map.items():
            try:
                df = pd.read_excel(RESULTS_XLSX, sheet_name=sheet_name, header=None)
            except Exception:
                continue
            
            header_row = None
            for i in range(min(15, len(df))):
                if df.iloc[i, 0] == "Sr.No." or (isinstance(df.iloc[i, 0], str) and "Sr.No" in str(df.iloc[i, 0])):
                    header_row = i
                    break
            
            if header_row is None:
                continue
            
            subject_row = header_row
            code_row = header_row + 1
            date_row = header_row + 2
            
            subjects = []
            for col in range(3, len(df.columns)):
                subj_name = df.iloc[subject_row, col]
                subj_code = df.iloc[code_row, col]
                subj_date = df.iloc[date_row, col]
                
                if pd.isna(subj_name) or subj_name == "" or "NAME OF SUBJECT" in str(subj_name):
                    continue
                
                subj_name_clean = str(subj_name).replace("\n", " ").strip()
                subj_code_clean = str(subj_code).replace("\n", " ").strip() if not pd.isna(subj_code) else "-"
                
                subj_date_clean = "-"
                if not pd.isna(subj_date):
                    if isinstance(subj_date, datetime):
                        subj_date_clean = subj_date.strftime("%d/%m/%Y")
                    else:
                        subj_date_clean = str(subj_date).strip()
                
                subjects.append({
                    "col": col,
                    "name": subj_name_clean,
                    "code": subj_code_clean,
                    "date": subj_date_clean,
                })
            
            students_data = []
            for row in range(date_row + 1, len(df)):
                sr = df.iloc[row, 0]
                name = df.iloc[row, 2]
                
                if pd.isna(sr) or pd.isna(name) or str(name).strip() == "":
                    continue
                if str(name).strip().upper() in ["NAME OF SUBJECT", "SUBJECT"]:
                    break
                
                student_marks = {}
                for subj in subjects:
                    val = df.iloc[row, subj["col"]]
                    if pd.isna(val):
                        marks_val = None
                    elif isinstance(val, str):
                        val_clean = val.strip().upper()
                        if val_clean in ["AB", "ABSENT", "A"]:
                            marks_val = "AB"
                        elif val_clean == "":
                            marks_val = None
                        else:
                            try:
                                marks_val = float(val)
                            except:
                                marks_val = None
                    else:
                        try:
                            marks_val = float(val)
                        except:
                            marks_val = None
                    
                    student_marks[subj["name"]] = marks_val
                
                students_data.append({
                    "sr": int(sr) if not pd.isna(sr) else 0,
                    "name": str(name).strip(),
                    "marks": student_marks,
                })
            
            results[branch] = {
                "subjects": [{"name": s["name"], "code": s["code"], "date": s["date"]} for s in subjects],
                "students": students_data,
            }
        
        save_json(RESULTS_FILE, results)
        return results
    
    except Exception as e:
        st.error(f"Error loading weekly results: {e}")
        return {}


# ============================================================
# SESSION STATE
# ============================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "page" not in st.session_state:
    st.session_state.page = "Dashboard"


# ============================================================
# LOGIN PAGE
# ============================================================
def login_page():
    logo = get_logo()
    if logo:
        col_l, col_c, col_r = st.columns([3.5, 2, 2])
        with col_c:
            st.image(logo, width=150)
    
    st.markdown("""
    <div style='text-align:center; padding: 10px 0 20px 0;'>
        <h1 style='color:#0d7377; margin: 0; font-size: 32px;'>
            OM Institute of Engineering & Technology
        </h1>
        <p style='color:#666; font-size:16px; margin-top: 10px;'>
            Weekly Test & Student Performance Management System
        </p>
        <p style='color:#999; font-size: 14px;'>Junagadh, Gujarat</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        with st.form("login_form"):
            st.markdown("### 🔐 Admin Login")
            username = st.text_input("Username", placeholder="admin")
            password = st.text_input("Password", type="password", placeholder="admin123")
            submit = st.form_submit_button("🔓 Login", use_container_width=True, type="primary")
            
            if submit:
                if username == "admin" and password == "admin123":
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("❌ Invalid credentials! Use admin / admin123")
        
        st.info("💡 **Default Credentials:** Username: `admin` | Password: `admin123`")


# ============================================================
# SIDEBAR (with LIVE Auto-Updating Clock)
# ============================================================
def sidebar():
    with st.sidebar:
        # AUTO-REFRESH EVERY 1 SECOND for live clock
        st_autorefresh(interval=1000, key="clock_refresh")
        
        logo = get_logo()
        if logo:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.image(logo, width=100)
        
        st.markdown("""
        <div style='background: linear-gradient(135deg, #0d7377, #14a0a5); 
                    padding: 15px; border-radius: 10px; color: white; 
                    margin-bottom: 20px; text-align: center;'>
            <h3 style='margin:0; font-size: 16px;'>🎓 OM Education</h3>
            <p style='margin:0; font-size:11px; opacity:0.9;'>Engg. & Technology</p>
        </div>
        """, unsafe_allow_html=True)
        
        menu = ["Dashboard", "Students", "Subjects", "Weekly Test Results", "Marks Entry", "Reports"]
        icons = ["🏠", "👥", "📚", "📋", "✏️", "📊"]
        
        for i, item in enumerate(menu):
            if st.button(f"{icons[i]}  {item}", use_container_width=True, 
                        key=f"nav_{item}",
                        type="primary" if st.session_state.page == item else "secondary"):
                st.session_state.page = item
                st.rerun()
        
        st.markdown("---")
        
        # ============================================
        # LIVE DATE & TIME (Auto-updates every second)
        # ============================================
        now = datetime.now()
        day_name = now.strftime("%A")
        date_str = f"{day_name}, {now.strftime('%d %b %Y')}"
        time_str = now.strftime("%I:%M:%S %p")
        
        st.markdown(f"""
        <div class='live-clock'>
            <b>👤 Logged in as:</b> Admin<br>
            <b>📅 Date:</b> {date_str}<br>
            <b>🕐 Time:</b> <span class='live-time'>{time_str}</span> 
            <span class='live-dot'></span>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()


# ============================================================
# DASHBOARD PAGE
# ============================================================
def dashboard_page():
    st.markdown("""
    <div class='main-header'>
        <h1 style='margin:0;'>📊 Dashboard</h1>
        <p style='margin:0; opacity:0.9;'>Overview & Quick Statistics</p>
    </div>
    """, unsafe_allow_html=True)
    
    df = load_students()
    results = load_weekly_results()
    
    total_students = len(df)
    total_subjects = sum(len(v) for v in DEFAULT_SUBJECTS.values())
    total_tests = sum(len(r.get("subjects", [])) for r in results.values())
    
    all_marks = []
    for branch, data in results.items():
        for s in data.get("students", []):
            for subj, m in s.get("marks", {}).items():
                if isinstance(m, (int, float)):
                    all_marks.append(m)
    avg_perf = (sum(all_marks) / len(all_marks)) / 20 * 100 if all_marks else 0
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>👥 Total Students</div>
            <div class='metric-value'>{total_students}</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>📚 Total Subjects</div>
            <div class='metric-value'>{total_subjects}</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>📅 Weekly Tests</div>
            <div class='metric-value'>{total_tests}</div>
        </div>""", unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>📈 Avg Performance</div>
            <div class='metric-value'>{avg_perf:.0f}%</div>
        </div>""", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("🏛️ Branch-wise Performance Summary")
    
    branch_summary = []
    for branch in BRANCHES:
        data = results.get(branch, {})
        marks_list = []
        for s in data.get("students", []):
            for m in s.get("marks", {}).values():
                if isinstance(m, (int, float)):
                    marks_list.append(m)
        
        avg = sum(marks_list) / len(marks_list) if marks_list else 0
        pass_count = sum(1 for m in marks_list if m >= PASS_MARKS)
        pass_pct = (pass_count / len(marks_list) * 100) if marks_list else 0
        
        branch_summary.append({
            "Branch": branch.split()[0],
            "Students": len(data.get("students", [])),
            "Avg Marks": round(avg, 1),
            "Avg %": f"{avg/20*100:.1f}%",
            "Pass %": f"{pass_pct:.1f}%",
        })
    
    summary_df = pd.DataFrame(branch_summary)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        fig = px.bar(summary_df, x="Branch", y="Avg %", color="Branch",
                    title="Branch-wise Average Performance",
                    color_discrete_sequence=px.colors.qualitative.Set2, text="Avg %")
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig, use_container_width=True)
    
    with col_b:
        fig2 = px.bar(summary_df, x="Branch", y="Pass %", color="Branch",
                     title="Branch-wise Pass Percentage",
                     color_discrete_sequence=px.colors.qualitative.Set3, text="Pass %")
        fig2.update_traces(textposition='outside')
        st.plotly_chart(fig2, use_container_width=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("🏛️ Student Distribution")
    
    branch_counts = df["Branch"].value_counts() if not df.empty else {}
    cols = st.columns(4)
    branch_icons = {"Civil Engineering": "🏗️", "Computer Engineering": "💻",
                    "Electrical Engineering": "⚡", "Mechanical Engineering": "⚙️"}
    
    for i, branch in enumerate(BRANCHES):
        count = branch_counts.get(branch, 0) if not df.empty else 0
        with cols[i]:
            st.markdown(f"""
            <div class='branch-card'>
                <h3>{branch_icons[branch]} {branch.split()[0]}</h3>
                <div class='count'>{count}</div>
                <div style='color:#666; font-size:12px;'>Students</div>
            </div>""", unsafe_allow_html=True)


# ============================================================
# STUDENTS PAGE
# ============================================================
def students_page():
    st.markdown("""
    <div class='main-header'>
        <h1 style='margin:0;'>👥 Student Management</h1>
        <p style='margin:0; opacity:0.9;'>Manage student records</p>
    </div>
    """, unsafe_allow_html=True)
    
    df = load_students()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Students", len(df))
    with col2:
        st.metric("Total Branches", df["Branch"].nunique() if not df.empty else 0)
    with col3:
        st.metric("Active Students", len(df[df["Status"] == "Active"]) if not df.empty else 0)
    
    st.markdown("---")
    
    c1, c2 = st.columns([2, 2])
    with c1:
        search = st.text_input("🔍 Search by Name or Roll No", placeholder="Enter name or roll no...")
    with c2:
        branch_filter = st.selectbox("Select Branch", ["All Branches"] + BRANCHES)
    
    filtered = df.copy()
    if search:
        filtered = filtered[
            filtered["Name"].str.contains(search.upper(), na=False) |
            filtered["Roll No"].str.contains(search, na=False)
        ]
    if branch_filter != "All Branches":
        filtered = filtered[filtered["Branch"] == branch_filter]
    
    st.markdown(f"### 📋 Student List ({len(filtered)})")
    
    if filtered.empty:
        st.warning("No students found.")
    else:
        display_df = filtered[["Sr", "Roll No", "Name", "Branch", "Semester", "Status"]].copy()
        st.dataframe(display_df, use_container_width=True, height=400)


# ============================================================
# SUBJECTS PAGE
# ============================================================
def subjects_page():
    st.markdown("""
    <div class='main-header'>
        <h1 style='margin:0;'>📚 Subject Management</h1>
        <p style='margin:0; opacity:0.9;'>View subjects with codes and faculty</p>
    </div>
    """, unsafe_allow_html=True)
    
    total = sum(len(v) for v in DEFAULT_SUBJECTS.values())
    st.metric("Total Active Subjects", total)
    
    branch_filter = st.selectbox("Filter by Branch", ["All Branches"] + BRANCHES)
    
    rows = []
    for branch, subs in DEFAULT_SUBJECTS.items():
        if branch_filter != "All Branches" and branch != branch_filter:
            continue
        for i, sub in enumerate(subs, 1):
            rows.append({
                "#": i,
                "Subject Code": sub["code"],
                "Subject Name": sub["name"],
                "Branch": branch,
                "Faculty": sub["faculty"],
                "Semester": SEMESTER,
                "Academic Year": ACADEMIC_YEAR,
                "Status": "Active"
            })
    
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, height=500)


# ============================================================
# WEEKLY TEST RESULTS PAGE (WITH EDIT FEATURES)
# ============================================================
def weekly_results_page():
    st.markdown("""
    <div class='main-header'>
        <h1 style='margin:0;'>📋 Weekly Test Results</h1>
        <p style='margin:0; opacity:0.9;'>Edit & manage test results</p>
    </div>
    """, unsafe_allow_html=True)
    
    results = load_weekly_results()
    
    if not results:
        st.warning("⚠️ No results loaded. Please ensure 'Weekly Test result.xlsx' is in the app folder.")
        return
    
    branch_tabs = st.tabs(BRANCHES)
    
    for idx, branch in enumerate(BRANCHES):
        with branch_tabs[idx]:
            data = results.get(branch, {})
            subjects = data.get("subjects", [])
            students = data.get("students", [])
            
            if not students:
                st.warning(f"No data for {branch}")
                continue
            
            st.markdown(f"#### 🏛️ {branch}")
            st.markdown(f"**Co-ordinator:** {BRANCH_COORDINATORS[branch]} | **Students:** {len(students)} | **Subjects:** {len(subjects)}")
            
            # =============================================
            # MANAGE SUBJECTS SECTION
            # =============================================
            with st.expander("⚙️ Manage Subjects (Add / Edit / Delete)", expanded=False):
                st.markdown("##### 📚 Current Subjects")
                
                for subj_idx, subj in enumerate(subjects):
                    c1, c2, c3, c4 = st.columns([2, 3, 2, 1])
                    with c1:
                        st.markdown(f"**{subj['code']}**")
                    with c2:
                        st.markdown(f"{subj['name']}")
                    with c3:
                        st.markdown(f"📅 {subj['date']}")
                    with c4:
                        if st.button("✏️", key=f"edit_subj_{branch}_{subj_idx}", help="Edit Subject"):
                            st.session_state[f"editing_subject_{branch}"] = subj_idx
                            st.rerun()
                
                if st.session_state.get(f"editing_subject_{branch}") is not None:
                    edit_idx = st.session_state[f"editing_subject_{branch}"]
                    if edit_idx < len(subjects):
                        subj_to_edit = subjects[edit_idx]
                        st.markdown(f"##### ✏️ Editing: **{subj_to_edit['name']}**")
                        
                        with st.form(f"edit_subject_form_{branch}_{edit_idx}"):
                            new_code = st.text_input("Subject Code", value=subj_to_edit['code'])
                            new_name = st.text_input("Subject Name", value=subj_to_edit['name'])
                            new_date = st.text_input("Test Date", value=subj_to_edit['date'])
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                save_btn = st.form_submit_button("💾 Save Changes", type="primary", use_container_width=True)
                            with c2:
                                cancel_btn = st.form_submit_button("❌ Cancel", use_container_width=True)
                            
                            if save_btn:
                                old_name = subj_to_edit['name']
                                subjects[edit_idx]['code'] = new_code
                                subjects[edit_idx]['name'] = new_name
                                subjects[edit_idx]['date'] = new_date
                                
                                if old_name != new_name:
                                    for s in students:
                                        if old_name in s['marks']:
                                            s['marks'][new_name] = s['marks'].pop(old_name)
                                
                                data['subjects'] = subjects
                                data['students'] = students
                                results[branch] = data
                                save_json(RESULTS_FILE, results)
                                st.session_state[f"editing_subject_{branch}"] = None
                                st.success("✅ Subject updated!")
                                st.rerun()
                            
                            if cancel_btn:
                                st.session_state[f"editing_subject_{branch}"] = None
                                st.rerun()
                
                st.markdown("---")
                
                st.markdown("##### ➕ Add New Subject")
                with st.form(f"add_subject_{branch}"):
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        new_subj_code = st.text_input("Subject Code", placeholder="DI0100XXXX")
                    with c2:
                        new_subj_name = st.text_input("Subject Name", placeholder="Subject Name")
                    with c3:
                        new_subj_date = st.text_input("Test Date", placeholder="DD/MM/YYYY")
                    
                    if st.form_submit_button("➕ Add Subject", type="primary", use_container_width=True):
                        if new_subj_code and new_subj_name:
                            new_subject = {
                                "code": new_subj_code.strip(),
                                "name": new_subj_name.strip(),
                                "date": new_subj_date.strip() or "-"
                            }
                            subjects.append(new_subject)
                            for s in students:
                                s['marks'][new_subj_name.strip()] = None
                            
                            data['subjects'] = subjects
                            data['students'] = students
                            results[branch] = data
                            save_json(RESULTS_FILE, results)
                            st.success(f"✅ Subject '{new_subj_name}' added!")
                            st.rerun()
                        else:
                            st.error("Subject Code and Name required!")
                
                st.markdown("---")
                
                st.markdown("##### 🗑️ Delete Subject")
                if subjects:
                    del_subj_name = st.selectbox(
                        "Select Subject to Delete",
                        [s['name'] for s in subjects],
                        key=f"del_subj_{branch}"
                    )
                    if st.button("🗑️ Delete Subject", type="primary", key=f"del_btn_{branch}"):
                        subjects = [s for s in subjects if s['name'] != del_subj_name]
                        for s in students:
                            s['marks'].pop(del_subj_name, None)
                        
                        data['subjects'] = subjects
                        data['students'] = students
                        results[branch] = data
                        save_json(RESULTS_FILE, results)
                        st.success(f"✅ '{del_subj_name}' deleted!")
                        st.rerun()
            
            # =============================================
            # SUBJECT-WISE PERFORMANCE
            # =============================================
            st.markdown("##### 📊 Subject-wise Performance")
            
            subject_analysis = []
            for subj in subjects:
                subj_name = subj["name"]
                marks_list = []
                for s in students:
                    m = s["marks"].get(subj_name)
                    if isinstance(m, (int, float)):
                        marks_list.append(m)
                
                total_count = len(marks_list)
                if total_count == 0:
                    continue
                
                passed = sum(1 for m in marks_list if m >= PASS_MARKS)
                failed = total_count - passed
                avg = sum(marks_list) / total_count
                
                subject_analysis.append({
                    "Subject Code": subj["code"],
                    "Subject": subj_name,
                    "Date": subj["date"],
                    "Appeared": total_count,
                    "Passed": passed,
                    "Failed": failed,
                    "Average": round(avg, 2),
                    "Pass %": f"{passed/total_count*100:.1f}%"
                })
            
            if subject_analysis:
                st.dataframe(pd.DataFrame(subject_analysis), use_container_width=True, hide_index=True)
            else:
                st.info("No marks entered yet.")
            
            # =============================================
            # EDIT STUDENT MARKS SECTION
            # =============================================
            with st.expander("✏️ Edit Student Marks", expanded=False):
                st.markdown("##### Update Marks for a Specific Student")
                
                if students:
                    student_options = [f"{s['sr']}. {s['name']}" for s in students]
                    selected_student = st.selectbox(
                        "Select Student",
                        student_options,
                        key=f"edit_student_{branch}"
                    )
                    student_idx = int(selected_student.split(".")[0]) - 1
                    current_student = students[student_idx]
                    
                    st.markdown(f"**Editing: {current_student['name']}**")
                    
                    with st.form(f"edit_marks_form_{branch}_{student_idx}"):
                        new_marks = {}
                        for subj_i, subj in enumerate(subjects):
                            subj_name = subj["name"]
                            current_val = current_student['marks'].get(subj_name)
                            
                            c1, c2, c3 = st.columns([3, 2, 2])
                            with c1:
                                st.markdown(f"**{subj_name}**")
                                st.caption(f"Code: {subj['code']}")
                            with c2:
                                if isinstance(current_val, (int, float)):
                                    default_val = int(current_val)
                                else:
                                    default_val = 0
                                m = st.number_input(
                                    "Marks (0-20)",
                                    min_value=0,
                                    max_value=20,
                                    value=default_val,
                                    key=f"edit_m_{branch}_{student_idx}_{subj_i}",
                                    label_visibility="collapsed"
                                )
                            with c3:
                                is_ab = st.checkbox(
                                    "AB",
                                    value=(current_val == "AB"),
                                    key=f"edit_ab_{branch}_{student_idx}_{subj_i}"
                                )
                            
                            new_marks[subj_name] = "AB" if is_ab else m
                        
                        if st.form_submit_button("💾 Save Marks", type="primary", use_container_width=True):
                            for subj_name, val in new_marks.items():
                                students[student_idx]['marks'][subj_name] = val
                            
                            data['students'] = students
                            results[branch] = data
                            save_json(RESULTS_FILE, results)
                            st.success(f"✅ Marks updated for {current_student['name']}!")
                            st.rerun()
            
            # =============================================
            # STUDENT-WISE MARKS TABLE
            # =============================================
            st.markdown("##### 📝 Student-wise Marks")
            
            table_rows = []
            for s in students:
                row = {"Sr": s["sr"], "Student Name": s["name"]}
                for subj in subjects:
                    m = s["marks"].get(subj["name"])
                    if m is None:
                        row[subj["name"][:15]] = "-"
                    elif m == "AB":
                        row[subj["name"][:15]] = "AB"
                    else:
                        row[subj["name"][:15]] = m
                table_rows.append(row)
            
            if table_rows:
                st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True, height=400)
            
            st.markdown("---")


# ============================================================
# MARKS ENTRY PAGE
# ============================================================
def marks_entry_page():
    st.markdown("""
    <div class='main-header'>
        <h1 style='margin:0;'>✏️ Marks Entry</h1>
        <p style='margin:0; opacity:0.9;'>Enter student marks (out of 20)</p>
    </div>
    """, unsafe_allow_html=True)
    
    results = load_weekly_results()
    
    if not results:
        st.warning("No test results available.")
        return
    
    st.info("📌 **Note:** Aa page ma tame weekly test na marks manually update kari shako.")
    
    c1, c2 = st.columns(2)
    with c1:
        selected_branch = st.selectbox("Select Branch", BRANCHES)
    data = results.get(selected_branch, {})
    subjects = data.get("subjects", [])
    
    if not subjects:
        st.warning(f"No subjects for {selected_branch}")
        return
    
    with c2:
        selected_subject = st.selectbox("Select Subject", [s["name"] for s in subjects])
    
    subj_info = next((s for s in subjects if s["name"] == selected_subject), {})
    st.markdown(f"**Code:** `{subj_info.get('code', '-')}` | **Date:** {subj_info.get('date', '-')} | **Max Marks:** 20")
    
    students = data.get("students", [])
    st.markdown(f"### 📝 Marks for {len(students)} students")
    
    with st.form("marks_form"):
        new_marks = {}
        for s in students:
            c1, c2, c3 = st.columns([3, 2, 2])
            with c1:
                st.markdown(f"**{s['sr']}.** {s['name']}")
            with c2:
                current = s["marks"].get(selected_subject)
                if isinstance(current, (int, float)):
                    val = int(current)
                else:
                    val = 0
                m = st.number_input("Marks", min_value=0, max_value=20, value=val,
                                   key=f"m_{selected_branch}_{s['sr']}",
                                   label_visibility="collapsed")
            with c3:
                is_ab = st.checkbox("AB/Absent", value=(current == "AB"),
                                   key=f"ab_{selected_branch}_{s['sr']}")
            new_marks[s["name"]] = "AB" if is_ab else m
        
        if st.form_submit_button("💾 Save Marks", type="primary", use_container_width=True):
            for s in students:
                s["marks"][selected_subject] = new_marks[s["name"]]
            results[selected_branch] = data
            save_json(RESULTS_FILE, results)
            st.success("✅ Marks updated!")
            st.rerun()


# ============================================================
# REPORTS PAGE
# ============================================================
def reports_page():
    st.markdown("""
    <div class='main-header'>
        <h1 style='margin:0;'>📊 Reports & Analysis</h1>
        <p style='margin:0; opacity:0.9;'>Generate & export performance reports</p>
    </div>
    """, unsafe_allow_html=True)
    
    df = load_students()
    results = load_weekly_results()
    
    tab1, tab2, tab3 = st.tabs(["👤 Student Performance", "📚 Subject Analysis", "📥 Export Data"])
    
    with tab1:
        st.subheader("👤 Individual Student Analysis")
        
        c1, c2 = st.columns(2)
        with c1:
            branch_sel = st.selectbox("Select Branch", BRANCHES, key="rp_branch")
        data = results.get(branch_sel, {})
        students = data.get("students", [])
        
        if not students:
            st.warning(f"No data for {branch_sel}")
            return
        
        with c2:
            student_sel = st.selectbox("Select Student", [f"{s['sr']}. {s['name']}" for s in students])
            student_idx = int(student_sel.split(".")[0]) - 1
            student = students[student_idx]
        
        st.markdown(f"#### 📋 Performance of **{student['name']}**")
        
        subjects = data.get("subjects", [])
        rows = []
        for subj in subjects:
            m = student["marks"].get(subj["name"])
            if isinstance(m, (int, float)):
                pct = f"{m/20*100:.1f}%"
                status = "Pass" if m >= PASS_MARKS else "Fail"
            elif m == "AB":
                pct = "-"
                status = "Absent"
            else:
                pct = "-"
                status = "-"
            
            rows.append({
                "Subject Code": subj["code"],
                "Subject": subj["name"],
                "Date": subj["date"],
                "Marks": m if m is not None else "-",
                "Max": 20,
                "Percentage": pct,
                "Status": status
            })
        
        report_df = pd.DataFrame(rows)
        st.dataframe(report_df, use_container_width=True, hide_index=True)
        
        valid = [r for r in rows if isinstance(r["Marks"], (int, float))]
        if valid:
            total_obt = sum(r["Marks"] for r in valid)
            total_max = 20 * len(valid)
            pct = total_obt / total_max * 100
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Tests Attended", len(valid))
            c2.metric("Total Obtained", f"{total_obt:.0f}/{total_max}")
            c3.metric("Overall %", f"{pct:.1f}%")
            grade = "A+" if pct >= 90 else "A" if pct >= 80 else "B" if pct >= 70 else "C" if pct >= 60 else "D"
            c4.metric("Grade", grade)
            
            chart_df = pd.DataFrame([r for r in rows if isinstance(r["Marks"], (int, float))])
            fig = px.bar(chart_df, x="Subject", y="Marks", text="Marks",
                        title=f"Performance - {student['name']}",
                        color="Status",
                        color_discrete_map={"Pass": "#14a0a5", "Fail": "#ef4444"})
            fig.update_traces(textposition='outside')
            fig.update_layout(xaxis_tickangle=-30)
            st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        st.subheader("📚 Subject-wise Analysis")
        
        branch_sel2 = st.selectbox("Select Branch", BRANCHES, key="rp_sub_branch")
        data2 = results.get(branch_sel2, {})
        subjects2 = data2.get("subjects", [])
        students2 = data2.get("students", [])
        
        if not subjects2:
            st.warning("No data.")
        else:
            analysis = []
            for subj in subjects2:
                marks_list = [s["marks"].get(subj["name"]) for s in students2]
                valid = [m for m in marks_list if isinstance(m, (int, float))]
                absent = sum(1 for m in marks_list if m == "AB")
                
                if valid:
                    passed = sum(1 for m in valid if m >= PASS_MARKS)
                    failed = len(valid) - passed
                    avg = sum(valid) / len(valid)
                    highest = max(valid)
                    lowest = min(valid)
                    
                    analysis.append({
                        "Code": subj["code"],
                        "Subject": subj["name"],
                        "Appeared": len(valid),
                        "Absent": absent,
                        "Passed": passed,
                        "Failed": failed,
                        "Pass %": f"{passed/len(valid)*100:.1f}%",
                        "Average": round(avg, 2),
                        "Highest": int(highest),
                        "Lowest": int(lowest)
                    })
            
            if analysis:
                analysis_df = pd.DataFrame(analysis)
                st.dataframe(analysis_df, use_container_width=True, hide_index=True)
                
                fig = px.bar(analysis_df, x="Subject", y="Average", text="Average",
                            title=f"{branch_sel2} - Subject-wise Average Marks",
                            color="Pass %", color_continuous_scale="Teal")
                fig.update_traces(textposition='outside')
                fig.update_layout(xaxis_tickangle=-30)
                st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        st.subheader("📥 Export Data")
        
        if not df.empty:
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button("📄 Students List (CSV)", csv,
                              "students_list.csv", "text/csv", use_container_width=True)
        
        for branch in BRANCHES:
            data_br = results.get(branch, {})
            if not data_br:
                continue
            
            subjects_br = data_br.get("subjects", [])
            students_br = data_br.get("students", [])
            
            rows = []
            for s in students_br:
                row = {"Sr": s["sr"], "Student Name": s["name"]}
                for subj in subjects_br:
                    m = s["marks"].get(subj["name"])
                    row[f"{subj['name']} ({subj['code']})"] = m if m is not None else "-"
                rows.append(row)
            
            if rows:
                rdf = pd.DataFrame(rows)
                csv_b = rdf.to_csv(index=False).encode('utf-8')
                st.download_button(f"📊 {branch} Results (CSV)", csv_b,
                                  f"{branch.replace(' ', '_')}_results.csv",
                                  "text/csv", use_container_width=True)


# ============================================================
# MAIN ROUTER
# ============================================================
def main():
    if not st.session_state.logged_in:
        login_page()
    else:
        sidebar()
        
        page = st.session_state.page
        if page == "Dashboard":
            dashboard_page()
        elif page == "Students":
            students_page()
        elif page == "Subjects":
            subjects_page()
        elif page == "Weekly Test Results":
            weekly_results_page()
        elif page == "Marks Entry":
            marks_entry_page()
        elif page == "Reports":
            reports_page()
        
        st.markdown("""
        <div class='creator-footer'>
            © 2026-27 OM Institute of Engineering & Technology, Junagadh<br>
            Weekly Test & Student Performance Management System<br>
            <span class='creator-name'>✨ Created by Dhara Pithva ✨</span>
        </div>
        """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()