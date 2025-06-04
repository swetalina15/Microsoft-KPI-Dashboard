import streamlit as st
import requests
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timezone
from io import BytesIO
import msal

# ------------------ App Setup ------------------
st.set_page_config(page_title="📊 KPI Dashboard", layout="wide")
st.markdown("""
    <style>
        .main-title {
            font-size: 40px;
            font-weight: bold;
            color: #1F4E79;
            text-align: center;
            margin-bottom: 20px;
        }
        .kpi-card {
            background: linear-gradient(145deg, #ffffff, #f2f2f2);
            padding: 25px;
            border-radius: 20px;
            box-shadow: 0 6px 16px rgba(0, 0, 0, 0.08);
            text-align: center;
            transition: transform 0.2s ease;
        }
        .kpi-card:hover {
            transform: scale(1.02);
        }
        .kpi-label {
            font-size: 18px;
            font-weight: 500;
            color: #333333;
        }
        .kpi-value {
            font-size: 36px;
            font-weight: bold;
            margin-top: 10px;
        }
        hr {
            margin: 40px 0;
            border: 1px solid #e0e0e0;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>📊 Microsoft Planner KPI Dashboard</div>", unsafe_allow_html=True)

# ------------------ MSAL Device Code Login ------------------
def msal_device_flow_login():
    if "access_token" not in st.session_state:
        app = msal.PublicClientApplication(
            st.secrets["CLIENT_ID"],
            authority=f"https://login.microsoftonline.com/{st.secrets['TENANT_ID']}"
        )
        flow = app.initiate_device_flow(scopes=["User.Read"])
        if "user_code" not in flow:
            st.error("Failed to create device flow. Please check your Azure app registration.")
            st.stop()
        st.info(f"Please authenticate by visiting: {flow['verification_uri']} and enter code: {flow['user_code']}")
        result = app.acquire_token_by_device_flow(flow)
        if "access_token" in result:
            st.session_state["access_token"] = result["access_token"]
            graph_resp = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers={"Authorization": f"Bearer {result['access_token']}"}
            )
            if graph_resp.status_code == 200:
                email = graph_resp.json().get("mail") or graph_resp.json().get("userPrincipalName")
                st.session_state["user_email"] = email.lower()
                return email.lower()
            else:
                st.error("Failed to fetch user profile from Microsoft Graph.")
                st.stop()
        else:
            st.error(f"Authentication failed: {result.get('error_description')}")
            st.stop()
    else:
        return st.session_state.get("user_email")

user_email = msal_device_flow_login()
st.write(f"Logged in as: {user_email}")

# ------------------ Access Token for Graph API (Client Credentials) ------------------
def get_access_token():
    url = f"https://login.microsoftonline.com/{st.secrets['TENANT_ID']}/oauth2/v2.0/token"
    payload = {
        "client_id": st.secrets["CLIENT_ID"],
        "client_secret": st.secrets["CLIENT_SECRET"],
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    res = requests.post(url, data=payload)
    res.raise_for_status()
    return res.json().get("access_token")

access_token = get_access_token()
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# ------------------ Load Reporting Hierarchy ------------------
df_hierarchy = pd.read_excel("reporting_hierarchy.xlsx")
df_hierarchy.columns = df_hierarchy.columns.str.strip()
df_hierarchy["Employee EmailID"] = df_hierarchy["Employee EmailID"].str.lower().str.strip()
df_hierarchy["Reporting Manager EmailID"] = df_hierarchy["Reporting Manager EmailID"].str.lower().str.strip()

# ------------------ Determine Access Scope ------------------
if user_email not in df_hierarchy["Reporting Manager EmailID"].values:
    scope_emails = [user_email]
else:
    reports = df_hierarchy[df_hierarchy["Reporting Manager EmailID"] == user_email]["Employee EmailID"].tolist()
    scope_emails = [user_email] + reports

# ------------------ Date Filter ------------------
date_filter = st.selectbox("🗕️ Filter by Date", ["This Month", "Last Month", "All Time"])
now = datetime.now(timezone.utc)
start_this_month = datetime(now.year, now.month, 1, tzinfo=timezone.utc)
start_next_month = datetime(now.year + 1, 1, 1, tzinfo=timezone.utc) if now.month == 12 else datetime(now.year, now.month + 1, 1, tzinfo=timezone.utc)
start_last_month = datetime(now.year - 1, 12, 1, tzinfo=timezone.utc) if now.month == 1 else datetime(now.year, now.month - 1, 1, tzinfo=timezone.utc)

def filter_by_date(df):
    if date_filter == "This Month":
        return df[(df["CreatedDate"] >= start_this_month) & (df["CreatedDate"] < start_next_month)]
    elif date_filter == "Last Month":
        return df[(df["CreatedDate"] >= start_last_month) & (df["CreatedDate"] < start_this_month)]
    return df

# ------------------ Planner Plan IDs ------------------
plan_ids = ["-dg9FJCoHkeg04AlKb_22ckAB08q", "1qTmx04ZQ0aUmfMRl-qDAMkAAShd", "9MwY0H0E1UipbdU_MQN1pskACY44", "HZUriORIbU2o6gb5wRpcPskAAOku", "LcvQROmlP0mjBFaizgn-6MkACnHV", "PJVx-ra-lU65RVcF_zOPcMkAHDIm", "Q-dOJFb1SkiuSMQiCIEZ2ckAEcKR", "SjFKBXJCqkucjHDUXmqfFckADR6Y", "_CSis4zCf0eODLqCuYG2iskACLvW", "hO9_bkDTgES372fKeT0QZckAC9JU", "rPvsaKHA3Eqt5QpO1TAlGckAEJEU", "s1IswOPOxkWD8AXZOv6EmskABJ4o", "Ny5u_Gfh9kygH1HZ4xOGKckABUX7"]

# ------------------ Fetch plan names ------------------
plan_names = {}
for plan_id in plan_ids:
    plan_resp = requests.get(f"https://graph.microsoft.com/v1.0/planner/plans/{plan_id}", headers=headers)
    if plan_resp.status_code == 200:
        plan_names[plan_id] = plan_resp.json().get("title", "Unknown Plan")
    else:
        plan_names[plan_id] = "Unknown Plan"

# ------------------ Fetch Tasks ------------------
all_tasks = []
user_cache = {}

for plan_id in plan_ids:
    plan_name = plan_names.get(plan_id, "Unknown Plan")

    # Fetch buckets for the plan
    bucket_url = f"https://graph.microsoft.com/v1.0/planner/plans/{plan_id}/buckets"
    res_buckets = requests.get(bucket_url, headers=headers)
    if res_buckets.status_code != 200:
        st.warning(f"Failed to fetch buckets for plan {plan_name}: {res_buckets.status_code}")
        continue
    bucket_map = {bucket["id"]: bucket["name"] for bucket in res_buckets.json().get("value", [])}

    # Fetch tasks for the plan
    task_url = f"https://graph.microsoft.com/v1.0/planner/plans/{plan_id}/tasks"
    res = requests.get(task_url, headers=headers)
    if res.status_code != 200:
        st.warning(f"Failed to fetch tasks for plan {plan_name}: {res.status_code}")
        continue

    for task in res.json().get("value", []):
        title = task.get("title", "")
        due = task.get("dueDateTime")
        created = task.get("createdDateTime")
        progress = task.get("percentComplete", 0)
        bucket_id = task.get("bucketId")
        bucket_name = bucket_map.get(bucket_id, "Unknown")

        for assigned_id in task.get("assignments", {}):
            if assigned_id not in user_cache:
                user_info = requests.get(f"https://graph.microsoft.com/v1.0/users/{assigned_id}", headers=headers)
                if user_info.status_code == 200:
                    user_cache[assigned_id] = user_info.json().get("userPrincipalName", "").lower()
                else:
                    user_cache[assigned_id] = None
            email = user_cache.get(assigned_id)
            if email:
                all_tasks.append({
                    "Title": title,
                    "AssignedTo": email,
                    "DueDate": due,
                    "CreatedDate": created,
                    "Progress": progress,
                    "Bucket": bucket_name,
                    "Plan": plan_name
                })

# ------------------ Task DataFrame ------------------
df_tasks = pd.DataFrame(all_tasks)
df_tasks = df_tasks[df_tasks["AssignedTo"].isin(scope_emails)]
df_tasks["DueDate"] = pd.to_datetime(df_tasks["DueDate"], utc=True, errors='coerce')
df_tasks["CreatedDate"] = pd.to_datetime(df_tasks["CreatedDate"], utc=True, errors='coerce')
df_tasks = df_tasks[df_tasks["CreatedDate"].notna()]
df_tasks["Status"] = df_tasks["Progress"].apply(lambda x: "✅ Completed" if x == 100 else ("🟡 Not Started" if x == 0 else "🔀 In Progress"))
df_tasks = filter_by_date(df_tasks)

# ------------------ Employee Filter ------------------
employee_list = sorted(df_tasks["AssignedTo"].unique())
selected_employee = st.selectbox("👥 Filter by Employee", ["All"] + employee_list)
if selected_employee != "All":
    df_tasks = df_tasks[df_tasks["AssignedTo"] == selected_employee]

# ------------------ KPI Function ------------------
def compute_kpi(df):
    total = len(df)
    completed = len(df[df["Progress"] == 100])
    not_started = len(df[df["Progress"] == 0])
    in_progress = len(df[(df["Progress"] > 0) & (df["Progress"] < 100)])
    overdue = len(df[(df["Progress"] < 100) & (df["DueDate"] < datetime.now(timezone.utc))])
    on_time = len(df[(df["Progress"] == 100) & (df["DueDate"] >= df["CreatedDate"])])
    kpi = int(((completed / total) * 50 if total else 0) + ((on_time / completed) * 50 if completed else 0))
    return total, completed, not_started, in_progress, overdue, on_time, kpi

assigned, completed, not_started, in_progress, overdue, on_time, kpi_score = compute_kpi(df_tasks)

# ------------------ KPI Display as Cards ------------------
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("### 📈 KPI Summary")
with st.container():
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6, kpi7 = st.columns(7)

    with kpi1:
        st.markdown(f"""<div class='kpi-card'>
        <div class='kpi-label'>📦 Assigned</div>
        <div class='kpi-value' style='color:#4B8BBE'>{assigned}</div>
        </div>""", unsafe_allow_html=True)

    with kpi2:
        st.markdown(f"""<div class='kpi-card'>
        <div class='kpi-label'>✅ Completed</div>
        <div class='kpi-value' style='color:green'>{completed}</div>
        </div>""", unsafe_allow_html=True)

    with kpi3:
        st.markdown(f"""<div class='kpi-card'>
        <div class='kpi-label'>🟡 Not Started</div>
        <div class='kpi-value' style='color:#E3B505'>{not_started}</div>
        </div>""", unsafe_allow_html=True)

    with kpi4:
        st.markdown(f"""<div class='kpi-card'>
        <div class='kpi-label'>🔀 In Progress</div>
        <div class='kpi-value' style='color:#007ACC'>{in_progress}</div>
        </div>""", unsafe_allow_html=True)

    with kpi5:
        st.markdown(f"""<div class='kpi-card'>
        <div class='kpi-label'>⏳ Overdue</div>
        <div class='kpi-value' style='color:#DE4B4B'>{overdue}</div>
        </div>""", unsafe_allow_html=True)

    with kpi6:
        st.markdown(f"""<div class='kpi-card'>
        <div class='kpi-label'>🚀 On Time</div>
        <div class='kpi-value' style='color:#2A9D8F'>{on_time}</div>
        </div>""", unsafe_allow_html=True)

    with kpi7:
        st.markdown(f"""<div class='kpi-card'>
        <div class='kpi-label'>⭐ KPI Score</div>
        <div class='kpi-value' style='color:#F4A261'>{kpi_score}</div>
        </div>""", unsafe_allow_html=True)

# ------------------ KPI Gauge Chart ------------------
fig_kpi = go.Figure(go.Indicator(
    mode="gauge+number",
    value=kpi_score,
    title={'text': "KPI Score"},
    gauge={
        'axis': {'range': [0, 100]},
        'bar': {'color': "seagreen"},
        'steps': [
            {'range': [0, 50], 'color': "lightcoral"},
            {'range': [50, 80], 'color': "khaki"},
            {'range': [80, 100], 'color': "lightgreen"}
        ]
    }
))
st.plotly_chart(fig_kpi, use_container_width=True)

# ------------------ Breakdown for Managers ------------------
if user_email in df_hierarchy["Reporting Manager EmailID"].values:
    st.markdown("### 👥 Employee KPI Breakdown")
    def kpi_breakdown_func(x):
        a, c, n, p, o, t, k = compute_kpi(x)
        return pd.Series({
            "Assigned": a,
            "Completed": c,
            "Not Started": n,
            "In Progress": p,
            "Overdue": o,
            "On Time": t,
            "KPI Score": k
        })
    breakdown = df_tasks.groupby("AssignedTo").apply(kpi_breakdown_func).reset_index()
    st.dataframe(breakdown)

# ------------------ Task Assignment Table ------------------
st.markdown("### 📋 Task Assignment Table")
st.dataframe(df_tasks[["Title", "AssignedTo", "Bucket", "Plan", "Status", "DueDate", "CreatedDate"]].sort_values(by="DueDate"))

# ------------------ Export to Excel ------------------
df_export = df_tasks.copy()
if pd.api.types.is_datetime64tz_dtype(df_export["DueDate"]):
    df_export["DueDate"] = df_export["DueDate"].dt.tz_localize(None)
if pd.api.types.is_datetime64tz_dtype(df_export["CreatedDate"]):
    df_export["CreatedDate"] = df_export["CreatedDate"].dt.tz_localize(None)

output = BytesIO()
df_export.to_excel(output, index=False, sheet_name="Tasks")
output.seek(0)

st.download_button(
    label="🗕 Export Tasks to Excel",
    data=output,
    file_name="kpi_tasks.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
