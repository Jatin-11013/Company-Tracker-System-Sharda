import streamlit as st
import pandas as pd
import json, os, hashlib, hmac, random, smtplib
from datetime import datetime, date, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from io import BytesIO

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_OK = True
except ImportError:
    GSPREAD_OK = False

st.set_page_config(page_title="Floated Company Tracking System", page_icon="🎓", layout="wide")

SUPER_ADMIN_USER      = "jatin_123"
SUPER_ADMIN_PASS_HASH = hashlib.sha256("jatin@123".encode()).hexdigest()

SCHOOL_OPTIONS = sorted(["SAHS","SAS","SBS","SBSR","SDAP","SET","SHSS","SMFE","SOE","SSCSE"])
PROGRAM_OPTIONS = sorted(["MSc","BSc","MBA","BTech","BPT","MPT","BBA","MCom","BA","MA"])
ROUND_NAMES = ["PPT","Written Test","GD","Technical Interview","Final/HR Round"]
TOKEN_SECRET = "placement_app_secret_2024_xK9#mP"

DATA_HEADERS = [
    "Manager Name","Floated Date","Floated By","Opportunity Type","Batch",
    "Company Name","Core/Non-Core","Company Domain","Job Profile","Job Location",
    "School","Program","Specialization","CTC (LPA)","Company Current Status",
    "No. of Positions","No. of Registrations","Interview Date",
    "Participation PPT","Participation Written Test","Participation GD",
    "Participation Technical Interview","Participation Final/HR Round",
    "Shortlisting PPT","Shortlisting Written Test","Shortlisting GD",
    "Shortlisting Technical Interview","Shortlisting Final/HR Round",
    "HR Name","HR Email","HR Contact Number",
    "Remarks","Submitted By","Submitted At",
]

# ═══════════════════════════════════════════════════════════════════════════════
# EMAIL — OTP
# ═══════════════════════════════════════════════════════════════════════════════
def send_otp_email(to_email, otp, name=""):
    try:
        sender    = st.secrets["email"]["sender"]
        app_pass  = st.secrets["email"]["app_password"]
        msg = MIMEMultipart("alternative")
        msg["Subject"] = "🔐 Password Reset OTP — Floated Company Tracking System"
        msg["From"]    = sender
        msg["To"]      = to_email
        body = f"""
Hi {name},

Your OTP for password reset is:

  {otp}

This OTP is valid for 10 minutes. Do not share it with anyone.

If you did not request this, please ignore this email.

— Floated Company Tracking System
"""
        msg.attach(MIMEText(body, "plain"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(sender, app_pass)
            s.sendmail(sender, to_email, msg.as_string())
        return True
    except Exception as e:
        st.error(f"Email send failed: {e}")
        return False

def generate_otp():
    return str(random.randint(100000, 999999))

# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def hash_password(pw): return hashlib.sha256(pw.encode()).hexdigest()

@st.cache_resource(show_spinner=False)
def _get_client():
    if not GSPREAD_OK: return None
    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=["https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(creds)
    except Exception: return None

@st.cache_resource(show_spinner=False)
def _get_all_worksheets():
    client = _get_client()
    if not client: return None, None, None, None
    try:
        sh = client.open_by_key(st.secrets["google_sheet"]["sheet_id"])
        titles = [ws.title for ws in sh.worksheets()]
        def get_or_create(title, headers):
            if title not in titles:
                ws = sh.add_worksheet(title=title, rows=2000, cols=max(50,len(headers)+5))
                ws.append_row(headers)
            else:
                ws = sh.worksheet(title)
            return ws
        return (sh,
                get_or_create("Managers", ["username","name","password","email"]),
                get_or_create("Admins",   ["username","name","password","email"]),
                get_or_create("PlacementData", DATA_HEADERS))
    except Exception: return None, None, None, None

def _ws_mgr():  return _get_all_worksheets()[1]
def _ws_adm():  return _get_all_worksheets()[2]
def _ws_data(): return _get_all_worksheets()[3]

SHEETS_CONFIGURED = False
try:
    SHEETS_CONFIGURED = "gcp_service_account" in st.secrets and "google_sheet" in st.secrets
except Exception: pass

@st.cache_data(ttl=60, show_spinner=False)
def load_users():
    ws = _ws_mgr()
    if not ws: return {}
    try:
        rows = ws.get_all_records()
        return {r["username"]: {"name":r["name"],"password":r["password"],"email":r.get("email","")}
                for r in rows if r.get("username")}
    except Exception: return {}

@st.cache_data(ttl=60, show_spinner=False)
def load_admins():
    ws = _ws_adm()
    if not ws: return {}
    try:
        rows = ws.get_all_records()
        return {r["username"]: {"name":r["name"],"password":r["password"],"email":r.get("email","")}
                for r in rows if r.get("username")}
    except Exception: return {}

@st.cache_data(ttl=30, show_spinner=False)
def load_data():
    ws = _ws_data()
    if not ws: return []
    try: return [dict(r) for r in ws.get_all_records()]
    except Exception: return []

def save_users(users):
    ws = _ws_mgr()
    if not ws: return
    try:
        ws.clear(); ws.append_row(["username","name","password","email"])
        if users: ws.append_rows([[u,d["name"],d["password"],d.get("email","")] for u,d in users.items()])
        load_users.clear()
    except Exception as e: st.warning(f"Could not save managers: {e}")

def save_admins(admins):
    ws = _ws_adm()
    if not ws: return
    try:
        ws.clear(); ws.append_row(["username","name","password","email"])
        if admins: ws.append_rows([[u,d["name"],d["password"],d.get("email","")] for u,d in admins.items()])
        load_admins.clear()
    except Exception as e: st.warning(f"Could not save admins: {e}")

def save_data(data):
    ws = _ws_data()
    if not ws: return
    try:
        ws.clear(); ws.append_row(DATA_HEADERS)
        if data:
            df = pd.DataFrame(data)
            for h in DATA_HEADERS:
                if h not in df.columns: df[h] = ""
            ws.append_rows(df[DATA_HEADERS].astype(str).values.tolist())
        load_data.clear()
    except Exception as e: st.warning(f"Could not save data: {e}")

def push_row(row):
    ws = _ws_data()
    if not ws: return
    try:
        ws.append_row([str(row.get(h,"")) for h in DATA_HEADERS])
        load_data.clear()
    except Exception as e: st.warning(f"Could not push row: {e}")

def is_super_admin(): return st.session_state.get("username") == SUPER_ADMIN_USER

def excel_bytes(df):
    b = BytesIO()
    try:
        with pd.ExcelWriter(b, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="Report")
    except Exception:
        b = BytesIO(); df.to_csv(b, index=False)
    return b.getvalue()

def _make_token(username, role, page=""):
    msg = f"{username}:{role}:{page}"
    sig = hmac.new(TOKEN_SECRET.encode(), msg.encode(), hashlib.sha256).hexdigest()[:16]
    return f"{username}:{role}:{page}:{sig}"

def _verify_token(token):
    try:
        parts = token.split(":")
        if len(parts) != 4: return None, None, None
        username, role, page, sig = parts
        expected = hmac.new(TOKEN_SECRET.encode(), f"{username}:{role}:{page}".encode(), hashlib.sha256).hexdigest()[:16]
        if hmac.compare_digest(sig, expected): return username, role, page
    except Exception: pass
    return None, None, None

def _default_page(role):
    return "admin_home" if role in ("super_admin","admin") else "manager_home"

def _update_token():
    if st.session_state.get("logged_in"):
        st.query_params["t"] = _make_token(
            st.session_state.username, st.session_state.role, st.session_state.page)

# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════
for k,v in {"logged_in":False,"role":None,"username":None,"page":"login",
            "edit_row_idx":None,"confirm_del":None,"active_filters":None,
            "_session_restored":False,
            # OTP state
            "otp_code":None,"otp_expiry":None,"otp_username":None,"otp_role":None,
            "fp_step":1}.items():
    if k not in st.session_state: st.session_state[k] = v

if not st.session_state._session_restored:
    st.session_state._session_restored = True
    try:
        token = st.query_params.get("t","")
        if token:
            username, role, page = _verify_token(token)
            if username and role:
                st.session_state.logged_in = True
                st.session_state.role      = role
                st.session_state.username  = username
                st.session_state.page      = page if page else _default_page(role)
    except Exception: pass

# ═══════════════════════════════════════════════════════════════════════════════
# CSS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""<style>
.main-header{background:linear-gradient(135deg,#1e3a5f,#2d6a9f);color:white;
  padding:1.5rem 2rem;border-radius:12px;margin-bottom:1.5rem;text-align:center;}
.main-header h1{margin:0;font-size:2rem;}
.main-header p{margin:.3rem 0 0;opacity:.85;}
.metric-box{background:linear-gradient(135deg,#f0f6ff,#e8f0fe);border-radius:10px;
  padding:1rem;text-align:center;border:1px solid #c5d8f5;}
.metric-box h3{margin:0;color:#1e3a5f;font-size:2rem;}
.metric-box p{margin:0;color:#5a7fa5;font-size:.85rem;}
div[data-testid="stButton"]>button{border-radius:8px;font-weight:600;}
</style>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# FORGOT PASSWORD — 3 step flow
# ═══════════════════════════════════════════════════════════════════════════════
def page_forgot_password():
    st.markdown('<div class="main-header"><h1>🔐 Forgot Password</h1><p>Reset your password via OTP</p></div>', unsafe_allow_html=True)
    _,col,_ = st.columns([1,1.2,1])
    with col:
        # ── Step 1: Enter username ───────────────────────────────────────────
        if st.session_state.fp_step == 1:
            st.markdown("### Step 1 — Enter your username")
            uname = st.text_input("Username")
            if st.button("📧 Send OTP", use_container_width=True, type="primary"):
                if not uname:
                    st.error("Please enter your username."); return
                # Find user in managers or admins (not super admin)
                users  = load_users()
                admins = load_admins()
                found_role = None; found_email = ""; found_name = ""
                if uname in users:
                    found_role  = "manager"
                    found_email = users[uname].get("email","")
                    found_name  = users[uname].get("name", uname)
                elif uname in admins:
                    found_role  = "admin"
                    found_email = admins[uname].get("email","")
                    found_name  = admins[uname].get("name", uname)
                else:
                    st.error("❌ Username not found."); return

                if not found_email:
                    st.error("❌ No email registered for this account. Please contact Super Admin."); return

                otp = generate_otp()
                if send_otp_email(found_email, otp, found_name):
                    st.session_state.otp_code     = otp
                    st.session_state.otp_expiry   = datetime.now() + timedelta(minutes=10)
                    st.session_state.otp_username = uname
                    st.session_state.otp_role     = found_role
                    st.session_state.fp_step      = 2
                    # Show masked email
                    parts = found_email.split("@")
                    masked = parts[0][:2] + "***@" + parts[1] if len(parts)==2 else "***"
                    st.success(f"✅ OTP sent to {masked}"); st.rerun()

        # ── Step 2: Enter OTP ────────────────────────────────────────────────
        elif st.session_state.fp_step == 2:
            st.markdown("### Step 2 — Enter OTP")
            # Check expiry
            if datetime.now() > st.session_state.otp_expiry:
                st.error("⏰ OTP expired. Please start again.")
                st.session_state.fp_step = 1
                st.session_state.otp_code = None; st.rerun()
            remaining = int((st.session_state.otp_expiry - datetime.now()).total_seconds() // 60)
            st.info(f"OTP valid for ~{remaining} more minute(s). Check your email.")
            otp_input = st.text_input("Enter 6-digit OTP", max_chars=6)
            c1,c2 = st.columns(2)
            with c1:
                if st.button("✅ Verify OTP", use_container_width=True, type="primary"):
                    if otp_input == st.session_state.otp_code:
                        st.session_state.fp_step = 3; st.rerun()
                    else:
                        st.error("❌ Wrong OTP. Try again.")
            with c2:
                if st.button("🔄 Resend OTP", use_container_width=True):
                    st.session_state.fp_step = 1
                    st.session_state.otp_code = None; st.rerun()

        # ── Step 3: Set new password ─────────────────────────────────────────
        elif st.session_state.fp_step == 3:
            st.markdown("### Step 3 — Set New Password")
            st.success(f"✅ OTP verified! Set a new password for **{st.session_state.otp_username}**")
            new_pw  = st.text_input("New Password", type="password")
            new_pw2 = st.text_input("Confirm Password", type="password")
            if st.button("💾 Save New Password", use_container_width=True, type="primary"):
                if not new_pw:
                    st.error("Password cannot be empty."); return
                if new_pw != new_pw2:
                    st.error("❌ Passwords do not match."); return
                if len(new_pw) < 6:
                    st.error("Password must be at least 6 characters."); return
                # Save new password
                uname = st.session_state.otp_username
                role  = st.session_state.otp_role
                if role == "manager":
                    users = load_users()
                    users[uname]["password"] = hash_password(new_pw)
                    save_users(users)
                elif role == "admin":
                    admins = load_admins()
                    admins[uname]["password"] = hash_password(new_pw)
                    save_admins(admins)
                # Reset OTP state
                for k in ["otp_code","otp_expiry","otp_username","otp_role"]:
                    st.session_state[k] = None
                st.session_state.fp_step = 1
                st.success("✅ Password changed successfully! Please login.")
                st.session_state.page = "login"; st.rerun()

        st.divider()
        if st.button("← Back to Login", use_container_width=True):
            st.session_state.fp_step = 1
            st.session_state.otp_code = None
            st.session_state.page = "login"; st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ═══════════════════════════════════════════════════════════════════════════════
def page_login():
    st.markdown('<div class="main-header"><h1>🎓 Floated Company Tracking System</h1><p>Login to continue</p></div>', unsafe_allow_html=True)
    if not SHEETS_CONFIGURED:
        st.error("⚠️ Google Sheets not connected. Add credentials in Streamlit Cloud → Settings → Secrets.")
    _,col,_ = st.columns([1,1.2,1])
    with col:
        st.markdown("### 🔐 Login")
        user = st.text_input("Username")
        pw   = st.text_input("Password", type="password")
        if st.button("🚀 Login", use_container_width=True, type="primary"):
            if not user or not pw: st.error("Please enter both username and password."); return
            ph = hash_password(pw)
            admins = load_admins(); users = load_users()
            if user == SUPER_ADMIN_USER and ph == SUPER_ADMIN_PASS_HASH:
                st.session_state.update(logged_in=True, role="super_admin", username=user, page="admin_home")
                st.query_params["t"] = _make_token(user,"super_admin","admin_home"); st.rerun()
            elif user in admins and admins[user]["password"] == ph:
                st.session_state.update(logged_in=True, role="admin", username=user, page="admin_home")
                st.query_params["t"] = _make_token(user,"admin","admin_home"); st.rerun()
            elif user in users and users[user]["password"] == ph:
                st.session_state.update(logged_in=True, role="manager", username=user, page="manager_home")
                st.query_params["t"] = _make_token(user,"manager","manager_home"); st.rerun()
            else:
                st.error("❌ Invalid username or password.")
        st.markdown("")
        if st.button("🔑 Forgot Password?", use_container_width=True):
            st.session_state.page = "forgot_password"
            st.session_state.fp_step = 1
            st.session_state.otp_code = None; st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
def render_sidebar():
    role = st.session_state.role
    lbl = {"super_admin":"Super Admin","admin":"Admin","manager":"Manager"}.get(role,role)
    with st.sidebar:
        st.markdown(f"### 👤 {st.session_state.username}")
        st.markdown(f"Role: **{lbl}**")
        st.divider()
        if role in ("super_admin","admin"):
            for pg,ic,tx in [("admin_home","🏠","Dashboard"),("manage_managers","👥","Manage Managers"),("admin_reports","📋","All Reports")]:
                if st.button(f"{ic} {tx}", use_container_width=True):
                    st.session_state.page=pg; _update_token(); st.rerun()
            if role=="super_admin":
                if st.button("🔑 Manage Admins", use_container_width=True):
                    st.session_state.page="manage_admins"; _update_token(); st.rerun()
            if st.button("🔄 Refresh Data", use_container_width=True):
                st.cache_resource.clear(); st.cache_data.clear(); st.rerun()
        else:
            for pg,ic,tx in [("manager_home","🏠","Home"),("data_entry","➕","New Entry"),("preview","📋","My Reports")]:
                if st.button(f"{ic} {tx}", use_container_width=True):
                    st.session_state.page=pg; _update_token(); st.rerun()
        st.divider()
        if st.button("🚪 Logout", use_container_width=True):
            st.query_params.clear()
            for k in list(st.session_state.keys()): del st.session_state[k]
            st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# ADMIN DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
def page_admin_home():
    lbl = "Super Admin" if is_super_admin() else "Admin"
    st.markdown(f'<div class="main-header"><h1>🏠 {lbl} Dashboard</h1></div>', unsafe_allow_html=True)
    data=load_data(); users=load_users()
    df = pd.DataFrame(data) if data else pd.DataFrame()
    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(f'<div class="metric-box"><h3>{len(users)}</h3><p>Managers</p></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="metric-box"><h3>{len(df)}</h3><p>Total Entries</p></div>', unsafe_allow_html=True)
    with c3:
        co = df["Company Name"].nunique() if not df.empty and "Company Name" in df.columns else 0
        st.markdown(f'<div class="metric-box"><h3>{co}</h3><p>Companies</p></div>', unsafe_allow_html=True)
    with c4:
        sel = int(df["Shortlisting Final/HR Round"].apply(pd.to_numeric,errors="coerce").sum()) if not df.empty and "Shortlisting Final/HR Round" in df.columns else 0
        st.markdown(f'<div class="metric-box"><h3>{sel}</h3><p>Final Round Shortlisted</p></div>', unsafe_allow_html=True)
    if not df.empty:
        st.divider(); st.subheader("📊 Recent Entries")
        st.dataframe(df.tail(10), use_container_width=True, hide_index=True)
        st.divider()
        c1,c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ Download All (Excel)", excel_bytes(df),
                file_name="all_placement_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with c2: st.info("☁️ Data is live-synced to Google Sheets automatically.")

# ═══════════════════════════════════════════════════════════════════════════════
# USER MANAGEMENT — with email field
# ═══════════════════════════════════════════════════════════════════════════════
def user_mgmt_ui(users, singular, save_fn, reserved):
    st.subheader(f"➕ Add New {singular}")
    with st.form(f"add_{singular}"):
        c1,c2 = st.columns(2)
        with c1: nu = st.text_input("Username")
        with c2: nn = st.text_input("Full Name")
        c3,c4 = st.columns(2)
        with c3: np_ = st.text_input("Password", type="password")
        with c4: ne  = st.text_input("Email *", placeholder="for OTP password reset")
        if st.form_submit_button(f"Add {singular}", type="primary", use_container_width=True):
            if not nu or not nn or not np_ or not ne: st.error("All fields required (including Email).")
            elif nu in users or nu in reserved: st.error("Username already exists!")
            else:
                users[nu]={"name":nn,"password":hash_password(np_),"email":ne}
                save_fn(users); st.success(f"✅ {singular} '{nn}' added!"); st.rerun()

    st.subheader(f"📋 Current {singular}s")
    if not users: st.info(f"No {singular.lower()}s yet.")
    for uname,ud in list(users.items()):
        c1,c2,c3,c4,c5 = st.columns([1.8,1.8,1.8,1,1])
        with c1: st.write(f"**{ud['name']}**")
        with c2: st.write(f"`{uname}`")
        with c3: st.write(f"📧 {ud.get('email','—')}")
        with c4:
            with st.popover("🔑 Change"):
                pw2=st.text_input("New PW",type="password",key=f"pw_{singular}_{uname}")
                em2=st.text_input("New Email",key=f"em_{singular}_{uname}",value=ud.get("email",""))
                if st.button("Update",key=f"upd_{singular}_{uname}"):
                    if pw2: users[uname]["password"]=hash_password(pw2)
                    if em2: users[uname]["email"]=em2
                    save_fn(users); st.success("Updated!"); st.rerun()
        with c5:
            ck=f"confirm_rm_{singular}_{uname}"
            if st.session_state.get(ck):
                st.warning(f"Remove **{ud['name']}**?")
                y,n=st.columns(2)
                with y:
                    if st.button("✅ Yes",key=f"yes_{singular}_{uname}",use_container_width=True,type="primary"):
                        del users[uname]; save_fn(users)
                        st.session_state.pop(ck,None); st.success(f"Removed {ud['name']}"); st.rerun()
                with n:
                    if st.button("❌ No",key=f"no_{singular}_{uname}",use_container_width=True):
                        st.session_state.pop(ck,None); st.rerun()
            else:
                if st.button("🗑️ Remove",key=f"rm_{singular}_{uname}",use_container_width=True):
                    st.session_state[ck]=True; st.rerun()
        st.divider()

def page_manage_managers():
    st.markdown('<div class="main-header"><h1>👥 Manage Managers</h1></div>', unsafe_allow_html=True)
    user_mgmt_ui(load_users(),"Manager",save_users,set(load_admins().keys())|{SUPER_ADMIN_USER})

def page_manage_admins():
    st.markdown('<div class="main-header"><h1>🔑 Manage Admins</h1></div>', unsafe_allow_html=True)
    user_mgmt_ui(load_admins(),"Admin",save_admins,set(load_users().keys())|{SUPER_ADMIN_USER})

# ═══════════════════════════════════════════════════════════════════════════════
# DATA ENTRY
# ═══════════════════════════════════════════════════════════════════════════════
def data_entry_form(prefill=None, edit_idx=None):
    p=prefill or {}; is_edit=(edit_idx is not None)
    st.markdown(f'<div class="main-header"><h1>{"✏️ Edit Entry" if is_edit else "➕ New Data Entry"}</h1></div>', unsafe_allow_html=True)
    with st.form("entry_form"):
        st.subheader("👤 Basic Information")
        _un=st.session_state.username; _rl=st.session_state.role
        _fn = load_users().get(_un,{}).get("name",_un) if _rl=="manager" else \
              (load_admins().get(_un,{}).get("name",_un) if _rl=="admin" else _un)
        c1,c2=st.columns(2)
        with c1: mgr_name=st.text_input("Manager Name *",value=p.get("Manager Name",_fn))
        with c2:
            fd=datetime.strptime(p["Floated Date"],"%Y-%m-%d").date() if p.get("Floated Date") else date.today()
            floated_date=st.date_input("Floated Date *",value=fd)
        c1,c2,c3=st.columns(3)
        with c1:
            fb=["Erp","Google Form"]
            floated_by=st.selectbox("Floated By *",fb,index=fb.index(p["Floated By"]) if p.get("Floated By") in fb else 0)
        with c2:
            ot=["Full Time","Intern cum PPO","Only Internship"]
            opp_type=st.selectbox("Opportunity Type *",ot,index=ot.index(p["Opportunity Type"]) if p.get("Opportunity Type") in ot else 0)
        with c3:
            bt=["2026-2027","2027-2028"]
            batch=st.selectbox("Batch *",bt,index=bt.index(p["Batch"]) if p.get("Batch") in bt else 0)

        st.subheader("🏢 Company Details")
        c1,c2,c3=st.columns(3)
        with c1: company_name=st.text_input("Company Name *",value=p.get("Company Name",""))
        with c2:
            cn=["Core","Non-Core"]
            core_nc=st.selectbox("Core / Non-Core *",cn,index=cn.index(p["Core/Non-Core"]) if p.get("Core/Non-Core") in cn else 0)
        with c3: company_domain=st.text_input("Company Domain",value=p.get("Company Domain",""))
        c1,c2=st.columns(2)
        with c1: job_profile=st.text_input("Job Profile",value=p.get("Job Profile",""))
        with c2: job_location=st.text_input("Job Location",value=p.get("Job Location",""))

        st.subheader("🏫 School, Program & Specialization")
        ds=[s.strip() for s in p["School"].split(",")] if p.get("School") else []
        dp=[s.strip() for s in p["Program"].split(",")] if p.get("Program") else []
        c1,c2,c3=st.columns(3)
        with c1: schools=st.multiselect("School *",SCHOOL_OPTIONS,default=[s for s in ds if s in SCHOOL_OPTIONS])
        with c2: programs=st.multiselect("Program *",PROGRAM_OPTIONS,default=[s for s in dp if s in PROGRAM_OPTIONS])
        with c3: specialization=st.text_input("Specialization",value=p.get("Specialization",""))

        st.subheader("💰 CTC & Status")
        c1,c2,c3,c4=st.columns(4)
        with c1:
            try: cv=float(str(p.get("CTC (LPA)","0")).replace(" LPA","").strip())
            except: cv=0.0
            ctc=st.number_input("CTC (LPA) *",min_value=0.0,step=0.1,value=cv,format="%.2f")
        with c2:
            cs=["In Process","Cancelled","Hold","Completed","Postponed"]
            company_status=st.selectbox("Company Status *",cs,index=cs.index(p["Company Current Status"]) if p.get("Company Current Status") in cs else 0)
        with c3:
            _pos_val=str(p.get("No. of Positions","Not Disclosed"))
            _slider_opts=["Not Disclosed"]+list(range(1,201))
            try: _pos_default=int(_pos_val) if _pos_val!="Not Disclosed" else "Not Disclosed"
            except: _pos_default="Not Disclosed"
            no_pos=st.select_slider("No. of Positions",options=_slider_opts,value=_pos_default)
        with c4:
            no_reg=st.number_input("No. of Registrations",min_value=0,step=1,value=int(p.get("No. of Registrations",0) or 0))
        iv=datetime.strptime(p["Interview Date"],"%Y-%m-%d").date() if p.get("Interview Date") else date.today()
        interview_date=st.date_input("Interview Date",value=iv)

        st.subheader("📊 Participations & Shortlistings")
        parts={}; c5=st.columns(5)
        for i,rnd in enumerate(ROUND_NAMES):
            with c5[i]:
                k=f"Participation {rnd}"
                parts[k]=st.number_input(f"Part. {rnd}",min_value=0,step=1,value=int(p.get(k,0) or 0),key=f"p{i}")
        shorts={}; c5b=st.columns(5)
        for i,rnd in enumerate(ROUND_NAMES):
            with c5b[i]:
                k=f"Shortlisting {rnd}"
                shorts[k]=st.number_input(f"Short. {rnd}",min_value=0,step=1,value=int(p.get(k,0) or 0),key=f"s{i}")

        st.subheader("👔 HR Details")
        c1,c2,c3=st.columns(3)
        with c1: hr_name=st.text_input("HR Name *",value=p.get("HR Name",""))
        with c2: hr_email=st.text_input("HR Email *",value=p.get("HR Email",""))
        with c3: hr_contact=st.text_input("HR Contact Number *",value=p.get("HR Contact Number",""))

        remarks=st.text_area("Remarks",value=p.get("Remarks",""),height=100)
        submitted=st.form_submit_button("💾 Update Entry" if is_edit else "💾 Submit Entry",
            type="primary",use_container_width=True)

    if submitted:
        errors=[]
        if not mgr_name:   errors.append("Manager Name")
        if not company_name: errors.append("Company Name")
        if not schools:    errors.append("School")
        if not programs:   errors.append("Program")
        if not hr_name:    errors.append("HR Name")
        if not hr_email:   errors.append("HR Email")
        if not hr_contact: errors.append("HR Contact Number")
        if errors: st.error(f"Please fill required fields: {', '.join(errors)}"); return
        row={
            "Manager Name":mgr_name,"Floated Date":str(floated_date),
            "Floated By":floated_by,"Opportunity Type":opp_type,"Batch":batch,
            "Company Name":company_name,"Core/Non-Core":core_nc,
            "Company Domain":company_domain,"Job Profile":job_profile,
            "Job Location":job_location,"School":", ".join(schools),
            "Program":", ".join(programs),"Specialization":specialization,
            "CTC (LPA)":f"{ctc:.2f} LPA","Company Current Status":company_status,
            "No. of Positions":str(no_pos),"No. of Registrations":no_reg,
            "Interview Date":str(interview_date),
            **parts,**shorts,
            "HR Name":hr_name,"HR Email":hr_email,"HR Contact Number":hr_contact,
            "Remarks":remarks,
            "Submitted By":st.session_state.username,
            "Submitted At":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        data=load_data()
        if is_edit: data[edit_idx]=row; st.success("✅ Entry updated!")
        else: data.append(row); push_row(row); st.success("✅ Entry submitted!")
        save_data(data)
        st.session_state.page="preview"; st.session_state.edit_row_idx=None
        _update_token(); st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# PREVIEW / REPORTS
# ═══════════════════════════════════════════════════════════════════════════════
def page_preview(admin_view=False):
    title="📋 All Reports" if admin_view else "📋 My Reports"
    st.markdown(f'<div class="main-header"><h1>{title}</h1></div>', unsafe_allow_html=True)
    data=load_data()
    if not data: st.info("No data found yet."); return
    df=pd.DataFrame(data)
    if not admin_view:
        df=df[df["Submitted By"]==st.session_state.username].reset_index(drop=True)
        if df.empty:
            st.info("No entries yet.")
            if st.button("➕ Add New Entry",type="primary"):
                st.session_state.page="data_entry"; _update_token(); st.rerun()
            return

    def _opts(col):
        if col not in df.columns: return []
        vals=set()
        for v in df[col].dropna().astype(str):
            for part in v.split(","):
                part=part.strip()
                if part and part!="nan": vals.add(part)
        return sorted(vals)

    with st.expander("🔍 Filter Data",expanded=False):
        with st.form("filter_form"):
            st.markdown("**🔎 Dynamic Dropdowns**")
            r1c1,r1c2,r1c3,r1c4=st.columns(4)
            with r1c1: f_co=st.multiselect("Company Name",_opts("Company Name"))
            with r1c2: f_do=st.multiselect("Company Domain",_opts("Company Domain"))
            with r1c3: f_jp=st.multiselect("Job Profile",_opts("Job Profile"))
            with r1c4: f_jl=st.multiselect("Job Location",_opts("Job Location"))
            r2c1,r2c2,r2c3,r2c4=st.columns(4)
            with r2c1: f_mn=st.multiselect("Manager Name",_opts("Manager Name"))
            with r2c2: f_sp=st.multiselect("Specialization",_opts("Specialization"))
            with r2c3: f_sc=st.multiselect("School",_opts("School"))
            with r2c4: f_pr=st.multiselect("Program",_opts("Program"))
            r3c1,r3c2,r3c3,r3c4=st.columns(4)
            with r3c1: f_hr=st.multiselect("HR Name",_opts("HR Name"))
            with r3c2: f_he=st.multiselect("HR Email",_opts("HR Email"))
            with r3c3: f_hc=st.multiselect("HR Contact",_opts("HR Contact Number"))
            with r3c4: f_re=st.multiselect("Remarks",_opts("Remarks"))
            st.markdown("**📋 Fixed Dropdowns**")
            d1,d2,d3,d4=st.columns(4)
            with d1: f_st=st.multiselect("Company Status",["In Process","Cancelled","Hold","Completed","Postponed"])
            with d2: f_fb=st.multiselect("Floated By",["Superset","Google Form"])
            with d3: f_ot=st.multiselect("Opportunity Type",["Full Time","Intern cum PPO","Only Internship"])
            with d4: f_cn=st.multiselect("Core/Non-Core",["Core","Non-Core"])
            d5,d6,_,__=st.columns(4)
            with d5: f_ba=st.multiselect("Batch",["2026-2027","2027-2028"])
            with d6:
                if admin_view: f_sb=st.multiselect("Submitted By",sorted(df["Submitted By"].unique().tolist()))
                else: f_sb=[]
            st.markdown("**📅 Date Filters**")
            dr1,dr2,dr3,dr4=st.columns(4)
            with dr1: f_fdf=st.date_input("Floated From",value=None,key="fdf")
            with dr2: f_fdt=st.date_input("Floated To",value=None,key="fdt")
            with dr3: f_idf=st.date_input("Interview From",value=None,key="idf")
            with dr4: f_idt=st.date_input("Interview To",value=None,key="idt")
            st.markdown("**🔢 Number Filters**")
            nr1,nr2,nr3,nr4=st.columns(4)
            with nr1: f_pmn=st.number_input("Min Positions",min_value=0,value=0,step=1)
            with nr2: f_pmx=st.number_input("Max Positions",min_value=0,value=99999,step=1)
            with nr3: f_smn=st.number_input("Min Shortlisting Final",min_value=0,value=0,step=1)
            with nr4: f_smx=st.number_input("Max Shortlisting Final",min_value=0,value=99999,step=1)
            apply_clicked=st.form_submit_button("✅ Apply Filters",type="primary",use_container_width=True)

        if apply_clicked:
            st.session_state.active_filters=dict(
                f_co=f_co,f_do=f_do,f_jp=f_jp,f_jl=f_jl,f_mn=f_mn,f_sp=f_sp,
                f_sc=f_sc,f_pr=f_pr,f_hr=f_hr,f_he=f_he,f_hc=f_hc,f_re=f_re,
                f_st=f_st,f_fb=f_fb,f_ot=f_ot,f_cn=f_cn,f_ba=f_ba,f_sb=f_sb,
                f_fdf=f_fdf,f_fdt=f_fdt,f_idf=f_idf,f_idt=f_idt,
                f_pmn=f_pmn,f_pmx=f_pmx,f_smn=f_smn,f_smx=f_smx)
        if st.button("🔄 Clear Filters"):
            st.session_state.active_filters=None; st.rerun()

    af=st.session_state.get("active_filters")
    if af is None:
        af=dict(f_co=[],f_do=[],f_jp=[],f_jl=[],f_mn=[],f_sp=[],f_sc=[],f_pr=[],
                f_hr=[],f_he=[],f_hc=[],f_re=[],f_st=[],f_fb=[],f_ot=[],f_cn=[],
                f_ba=[],f_sb=[],f_fdf=None,f_fdt=None,f_idf=None,f_idt=None,
                f_pmn=0,f_pmx=99999,f_smn=0,f_smx=99999)

    flt=df.copy()
    if af["f_co"]: flt=flt[flt["Company Name"].isin(af["f_co"])]
    if af["f_do"]: flt=flt[flt["Company Domain"].isin(af["f_do"])]
    if af["f_jp"]: flt=flt[flt["Job Profile"].isin(af["f_jp"])]
    if af["f_jl"]: flt=flt[flt["Job Location"].isin(af["f_jl"])]
    if af["f_mn"]: flt=flt[flt["Manager Name"].isin(af["f_mn"])]
    if af["f_sp"]: flt=flt[flt["Specialization"].isin(af["f_sp"])]
    if af["f_re"] and "Remarks" in flt.columns: flt=flt[flt["Remarks"].isin(af["f_re"])]
    if af["f_hr"] and "HR Name" in flt.columns: flt=flt[flt["HR Name"].isin(af["f_hr"])]
    if af["f_he"] and "HR Email" in flt.columns: flt=flt[flt["HR Email"].isin(af["f_he"])]
    if af["f_hc"] and "HR Contact Number" in flt.columns: flt=flt[flt["HR Contact Number"].isin(af["f_hc"])]
    if af["f_sc"]:
        flt=flt[flt["School"].apply(lambda x: any(s in [v.strip() for v in str(x).split(",")] for s in af["f_sc"]))]
    if af["f_pr"]:
        flt=flt[flt["Program"].apply(lambda x: any(s in [v.strip() for v in str(x).split(",")] for s in af["f_pr"]))]
    if af["f_st"]: flt=flt[flt["Company Current Status"].isin(af["f_st"])]
    if af["f_ba"]: flt=flt[flt["Batch"].isin(af["f_ba"])]
    if af["f_fb"]: flt=flt[flt["Floated By"].isin(af["f_fb"])]
    if af["f_ot"]: flt=flt[flt["Opportunity Type"].isin(af["f_ot"])]
    if af["f_cn"]: flt=flt[flt["Core/Non-Core"].isin(af["f_cn"])]
    if af["f_sb"]: flt=flt[flt["Submitted By"].isin(af["f_sb"])]
    if af["f_fdf"] and "Floated Date" in flt.columns:
        flt=flt[pd.to_datetime(flt["Floated Date"],errors="coerce")>=pd.Timestamp(af["f_fdf"])]
    if af["f_fdt"] and "Floated Date" in flt.columns:
        flt=flt[pd.to_datetime(flt["Floated Date"],errors="coerce")<=pd.Timestamp(af["f_fdt"])]
    if af["f_idf"] and "Interview Date" in flt.columns:
        flt=flt[pd.to_datetime(flt["Interview Date"],errors="coerce")>=pd.Timestamp(af["f_idf"])]
    if af["f_idt"] and "Interview Date" in flt.columns:
        flt=flt[pd.to_datetime(flt["Interview Date"],errors="coerce")<=pd.Timestamp(af["f_idt"])]
    if "No. of Positions" in flt.columns and (af["f_pmn"]>0 or af["f_pmx"]<99999):
        flt=flt[flt["No. of Positions"].apply(lambda x: pd.to_numeric(x,errors="coerce") if x!="Not Disclosed" else None).between(af["f_pmn"],af["f_pmx"])]
    if "Shortlisting Final/HR Round" in flt.columns:
        flt=flt[pd.to_numeric(flt["Shortlisting Final/HR Round"],errors="coerce").fillna(0).between(af["f_smn"],af["f_smx"])]

    flt=flt.reset_index(drop=False).rename(columns={"index":"_oi"})
    disp=flt.drop(columns=["_oi"])
    st.markdown(f"**{len(flt)} entries found**")
    st.download_button("⬇️ Download as Excel",excel_bytes(disp),
        file_name=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.divider()
    if flt.empty: st.info("No entries match the filters."); return

    can_modify=(st.session_state.role=="manager")
    if st.session_state.confirm_del is not None:
        oi=st.session_state.confirm_del; all_d=load_data()
        co_name=all_d[oi].get("Company Name","this entry") if oi<len(all_d) else "this entry"
        st.warning(f"⚠️ Delete **{co_name}**? This cannot be undone.")
        ya,na=st.columns(2)
        with ya:
            if st.button("✅ Yes, Delete",type="primary",use_container_width=True):
                all_d.pop(oi); save_data(all_d)
                st.session_state.confirm_del=None; st.success("Deleted!"); st.rerun()
        with na:
            if st.button("❌ Cancel",use_container_width=True):
                st.session_state.confirm_del=None; st.rerun()
        st.divider()

    st.markdown("#### 📊 Full Data Table")
    st.caption("Scroll right to see all columns.")
    st.dataframe(disp,use_container_width=True,hide_index=True,height=min(60+35*len(flt),500))
    st.divider()
    if not can_modify: return

    st.markdown("#### ✏️ Edit / Delete Entries")
    SHOW_COLS=[("Company Name",2.0),("Floated Date",1.1),("Batch",0.9),
               ("Opportunity Type",1.3),("Core/Non-Core",0.9),("CTC (LPA)",0.9),
               ("Company Current Status",1.3),("HR Name",1.0),("Submitted By",1.0)]
    SHOW_COLS=[(c,w) for c,w in SHOW_COLS if c in disp.columns]
    col_names=[c for c,_ in SHOW_COLS]; col_widths=[w for _,w in SHOW_COLS]
    all_widths=col_widths+[0.6,0.6]
    hcols=st.columns(all_widths)
    for hc,lbl in zip(hcols,col_names+["✏️","🗑️"]):
        hc.markdown(f"<div style='font-size:.72rem;font-weight:700;color:#2d6a9f;padding:4px 2px;border-bottom:2px solid #2d6a9f'>{lbl}</div>",unsafe_allow_html=True)
    for _,row in flt.iterrows():
        oi=int(row["_oi"]); rcols=st.columns(all_widths)
        for ci,cname in enumerate(col_names):
            val=str(row.get(cname,"—")); dv=val[:24]+"…" if len(val)>24 else val
            rcols[ci].markdown(f"<div style='font-size:.80rem;padding:4px 2px;border-bottom:1px solid #eef1f7'><span title='{val}'>{dv}</span></div>",unsafe_allow_html=True)
        with rcols[len(col_names)]:
            if st.button("✏️",key=f"ed_{oi}",use_container_width=True,type="primary",help="Edit"):
                st.session_state.edit_row_idx=oi; st.session_state.page="edit_entry"
                _update_token(); st.rerun()
        with rcols[len(col_names)+1]:
            if st.button("🗑️",key=f"dl_{oi}",use_container_width=True,help="Delete"):
                st.session_state.confirm_del=oi; st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# MANAGER HOME
# ═══════════════════════════════════════════════════════════════════════════════
def page_manager_home():
    uname=st.session_state.username
    name=load_users().get(uname,{}).get("name",uname)
    st.markdown(f'<div class="main-header"><h1>👋 Welcome, {name}!</h1><p>Floated Company Tracking Portal</p></div>', unsafe_allow_html=True)
    data=load_data(); my=[r for r in data if r.get("Submitted By")==uname]
    c1,c2,c3=st.columns(3)
    with c1: st.markdown(f'<div class="metric-box"><h3>{len(my)}</h3><p>My Total Entries</p></div>',unsafe_allow_html=True)
    with c2:
        co=len(set(r.get("Company Name","") for r in my))
        st.markdown(f'<div class="metric-box"><h3>{co}</h3><p>Companies</p></div>',unsafe_allow_html=True)
    with c3:
        fin=sum(int(r.get("Shortlisting Final/HR Round",0) or 0) for r in my)
        st.markdown(f'<div class="metric-box"><h3>{fin}</h3><p>Final Round Shortlisted</p></div>',unsafe_allow_html=True)
    st.divider()
    c1,c2=st.columns(2)
    with c1:
        if st.button("➕ New Data Entry",use_container_width=True,type="primary"):
            st.session_state.page="data_entry"; _update_token(); st.rerun()
    with c2:
        if st.button("📋 View My Reports",use_container_width=True):
            st.session_state.page="preview"; _update_token(); st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# ROUTER
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    if not st.session_state.logged_in:
        if st.session_state.page == "forgot_password":
            page_forgot_password()
        else:
            page_login()
        return
    render_sidebar()
    pg=st.session_state.page; role=st.session_state.role
    if   pg=="admin_home":      page_admin_home()
    elif pg=="manage_managers": page_manage_managers()
    elif pg=="manage_admins":
        if role=="super_admin": page_manage_admins()
        else: st.error("Access denied.")
    elif pg=="admin_reports":   page_preview(admin_view=True)
    elif pg=="manager_home":    page_manager_home()
    elif pg=="data_entry":      data_entry_form()
    elif pg=="edit_entry":
        idx=st.session_state.edit_row_idx
        if idx is not None: data_entry_form(prefill=load_data()[idx],edit_idx=idx)
        else: st.session_state.page="preview"; _update_token(); st.rerun()
    elif pg=="preview": page_preview(admin_view=(role in ("admin","super_admin")))

main()
