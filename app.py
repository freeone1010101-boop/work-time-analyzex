import io
import re
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Work Time Analyzer", page_icon="⏱️", layout="wide")

# ---------------- Theme ----------------
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg,#07111f 0%,#0d1b2a 45%,#132238 100%);
    color:#f4f7fb;
}
.block-container {max-width: 1400px; padding-top: 1.5rem;}
.hero {
    padding: 28px 30px; border-radius: 24px; margin-bottom: 22px;
    background: linear-gradient(120deg,#17365d,#0b6b6b,#274e6f);
    box-shadow: 0 18px 50px rgba(0,0,0,.28);
    animation: floatIn .7s ease-out;
}
.hero h1 {margin:0; font-size:2.25rem;}
.hero p {margin:.4rem 0 0; opacity:.82;}
.card {
    padding:18px; border-radius:18px; background:rgba(255,255,255,.07);
    border:1px solid rgba(255,255,255,.1); margin-bottom:12px;
}
@keyframes floatIn {from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}
div.stButton > button, div.stDownloadButton > button {
    border-radius:12px; font-weight:700; min-height:44px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>⏱️ Work Time Analyzer</h1>
<p>Daily & monthly attendance • 09:30–18:30 company rules • WOP-aware overtime</p>
</div>
""", unsafe_allow_html=True)

# ---------------- Login ----------------
USERS = {
    "admin": ("Admin@123", "admin"),
    "user1": ("Work@123", "user"),
    "user2": ("Work@123", "user"),
    "user3": ("Work@123", "user"),
    "user4": ("Work@123", "user"),
}
if "auth" not in st.session_state:
    st.session_state.auth = False
if not st.session_state.auth:
    st.markdown('<div class="card"><h2>🔐 Sign in</h2><p>Authorized company users only.</p></div>', unsafe_allow_html=True)
    with st.form("login"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Sign in", use_container_width=True):
            if u in USERS and USERS[u][0] == p:
                st.session_state.auth = True
                st.session_state.user = u
                st.rerun()
            else:
                st.error("Invalid username or password.")
    st.stop()

with st.sidebar:
    st.success(f"Signed in: {st.session_state.user}")
    if st.button("Sign out"):
        st.session_state.clear()
        st.rerun()
    st.markdown("---")
    st.markdown("### Company rules")
    st.write("Start: **09:30**")
    st.write("End: **18:30**")
    st.write("Saturday: **WOP working day**")
    st.write("Sunday: **WO weekly off**")

# ---------------- Helpers ----------------
def mins(v):
    if pd.isna(v) or v is None:
        return 0
    if isinstance(v, timedelta):
        return int(v.total_seconds() // 60)
    s = str(v).strip()
    if not s or s.lower() in {"nan", "nat", "none"}:
        return 0
    m = re.match(r"^(\d+):(\d{1,2})(?::(\d{1,2}))?$", s)
    if m:
        return int(m.group(1))*60 + int(m.group(2))
    try:
        x = float(s)
        return int(round(x * 24 * 60)) if x < 1 else int(round(x))
    except:
        return 0

def hm(n):
    n = int(max(0, n))
    return f"{n//60:02d}:{n%60:02d}"

def status_norm(v):
    s = str(v).strip().upper()
    if s in {"P","PRESENT"}: return "P"
    if s in {"A","ABSENT"}: return "A"
    if s in {"WOP","SAT","SATURDAY"}: return "WOP"
    if s in {"WO","SUN","SUNDAY"}: return "WO"
    return s

def parse_time(v):
    if pd.isna(v): return None
    if isinstance(v, datetime): return v
    s = str(v).strip()
    if not s or s.lower() in {"nan","nat","none","-"}: return None
    for fmt in ("%H:%M:%S","%H:%M","%I:%M %p","%I:%M:%S %p"):
        try:
            return datetime.strptime(s, fmt)
        except: pass
    try:
        if isinstance(v, (float,int)) and 0 <= float(v) < 1:
            return datetime(1900,1,1) + timedelta(days=float(v))
    except: pass
    return None

def actual_duration(in_v, out_v):
    a,b = parse_time(in_v), parse_time(out_v)
    if not a or not b: return 0
    delta = b-a
    if delta.total_seconds() < 0: delta += timedelta(days=1)
    return int(delta.total_seconds()//60)

def late_minutes(in_v, status):
    if status in {"A","WO"}: return 0
    t = parse_time(in_v)
    if not t: return 0
    start = datetime(1900,1,1,9,30)
    return max(0, int((t-start).total_seconds()//60))

def early_minutes(out_v, status):
    if status in {"A","WO"}: return 0
    t = parse_time(out_v)
    if not t: return 0
    end = datetime(1900,1,1,18,30)
    return max(0, int((end-t).total_seconds()//60))

def overtime_minutes(out_v, status):
    # WOP OT is explicitly excluded from regular OT.
    if status == "WOP":
        return 0
    if status in {"A","WO"}:
        return 0
    t = parse_time(out_v)
    if not t: return 0
    end = datetime(1900,1,1,18,30)
    return max(0, int((t-end).total_seconds()//60))

# ---------------- Source parser ----------------
def find_header(raw):
    for i in range(min(len(raw), 30)):
        vals = [str(x).strip().lower() for x in raw.iloc[i].tolist()]
        joined = " ".join(vals)
        if ("employee" in joined or "name" in joined) and ("status" in joined or "in" in joined):
            return i
    return 0

def clean_daily_sheet(df, sheet_name):
    # First try a conventional table.
    raw = df.copy()
    h = find_header(raw)
    table = pd.read_excel(io.BytesIO(st.session_state.current_bytes), sheet_name=sheet_name, header=h)
    table.columns = [str(c).strip() for c in table.columns]
    lower = {str(c).lower(): c for c in table.columns}

    def col(*names):
        for n in names:
            if n.lower() in lower: return lower[n.lower()]
        for c in table.columns:
            cl = str(c).lower()
            if any(n.lower() in cl for n in names): return c
        return None

    emp = col("Employee Name","Employee","Name")
    status = col("Status","Attendance Status")
    intime = col("InTime","In Time","In")
    outtime = col("OutTime","Out Time","Out")
    duration = col("Duration")
    otcol = col("OT","Overtime")
    datecol = col("Date","Attendance Date","Work Date")

    # If source has one row per employee/day, parse directly.
    if emp and (status or intime or outtime):
        rows=[]
        for _,r in table.iterrows():
            name=str(r.get(emp,"")).strip()
            if not name or name.lower() in {"nan","employee name","name"}: continue
            st=status_norm(r.get(status,"")) if status else ""
            iv=r.get(intime,"") if intime else ""
            ov=r.get(outtime,"") if outtime else ""
            dur=mins(r.get(duration,0)) if duration else actual_duration(iv,ov)
            actual=actual_duration(iv,ov)
            # For WOP, source duration can be 00:00; use actual In/Out duration.
            if st=="WOP" and actual>0: dur=actual
            source_ot=mins(r.get(otcol,0)) if otcol else 0
            reg_ot=overtime_minutes(ov,st)
            # If source OT exists for normal days, prefer it only when it agrees with
            # the post-18:30 rule; WOP is always removed from regular OT.
            if st=="WOP": reg_ot=0
            elif source_ot>0: reg_ot=source_ot
            date_val=r.get(datecol,"") if datecol else ""
            date=pd.to_datetime(date_val,errors="coerce") if datecol else pd.NaT
            rows.append({
                "Date": date.strftime("%Y-%m-%d") if pd.notna(date) else "",
                "Employee": name,
                "Status": st,
                "In Time": iv,
                "Out Time": ov,
                "Work Duration": dur,
                "Regular OT": reg_ot,
                "Late By": late_minutes(iv,st),
                "Early By": early_minutes(ov,st),
                "WOP OT Excluded": source_ot if st=="WOP" else 0
            })
        if rows: return pd.DataFrame(rows)

    return pd.DataFrame()

def process_workbook(file_bytes):
    xls=pd.ExcelFile(io.BytesIO(file_bytes))
    frames=[]
    for sh in xls.sheet_names:
        try:
            f=clean_daily_sheet(pd.read_excel(io.BytesIO(file_bytes),sheet_name=sh,header=None),sh)
            if not f.empty: frames.append(f)
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    out=pd.concat(frames,ignore_index=True)
    out=out.drop_duplicates(subset=["Date","Employee","Status","In Time","Out Time"],keep="first")
    return out

# ---------------- Upload ----------------
st.markdown("### 📁 Upload attendance report(s)")
files=st.file_uploader("Upload .xlsx or .xls attendance files",type=["xlsx","xls"],accept_multiple_files=True)
if not files:
    st.markdown('<div class="card">Upload one or more daily/monthly attendance reports to begin.</div>',unsafe_allow_html=True)
    st.stop()

all_frames=[]
for f in files:
    st.session_state.current_bytes=f.getvalue()
    parsed=process_workbook(f.getvalue())
    if not parsed.empty:
        parsed["Source File"]=f.name
        all_frames.append(parsed)

if not all_frames:
    st.error("No employee attendance rows could be detected. Please check the Excel format.")
    st.stop()

data=pd.concat(all_frames,ignore_index=True)
data["Date"]=data["Date"].replace("",pd.NA)
data["Date"]=pd.to_datetime(data["Date"],errors="coerce")
data["Status"]=data["Status"].map(status_norm)

# ---------------- Summary ----------------
st.success(f"Loaded {len(files)} file(s) • {data['Employee'].nunique()} employees • {data['Date'].dropna().nunique()} dates")

def summary_for(df):
    rows=[]
    for name,g in df.groupby("Employee",sort=True):
        present=int((g.Status=="P").sum())
        absent=int((g.Status=="A").sum())
        wop=int((g.Status=="WOP").sum())
        wo=int((g.Status=="WO").sum())
        work=int(g["Work Duration"].sum())
        ot=int(g["Regular OT"].sum())
        late=int(g["Late By"].sum())
        early=int(g["Early By"].sum())
        rows.append({
            "Employee":name,
            "Total Work Duration":hm(work),
            "Total OT":hm(ot),
            "Present":present,
            "Absent":absent,
            "WOP":wop,
            "WeeklyOff":wo,
            "Late By Hrs":hm(late),
            "Late By Days":int((g["Late By"]>0).sum()),
            "Early By Hrs":hm(early),
            "Early Going By Days":int((g["Early By"]>0).sum()),
            "WOP OT Excluded":hm(int(g["WOP OT Excluded"].sum())),
            "Total Duration (+OT)":hm(work+ot),
            "Average Working Hrs":hm(round(work/present) if present else 0)
        })
    return pd.DataFrame(rows)

tab1,tab2,tab3,tab4=st.tabs(["📊 Dashboard","📅 Daily / Monthly","👤 Employee","📤 Export"])

with tab1:
    s=summary_for(data)
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Employees",data.Employee.nunique())
    c2.metric("Present",int((data.Status=="P").sum()))
    c3.metric("Regular OT",hm(int(data["Regular OT"].sum())))
    c4.metric("WOP OT excluded",hm(int(data["WOP OT Excluded"].sum())))
    st.dataframe(s,use_container_width=True,hide_index=True)

with tab2:
    st.markdown("#### Date / status records")
    show=data.copy()
    show["Date"]=show["Date"].dt.strftime("%Y-%m-%d")
    st.dataframe(show[["Date","Employee","Status","In Time","Out Time","Work Duration","Regular OT","Late By","Early By","WOP OT Excluded","Source File"]],use_container_width=True,hide_index=True)
    st.download_button("⬇️ Download monthly CSV",show.to_csv(index=False).encode(),file_name="work_time_monthly.csv",mime="text/csv")

with tab3:
    names=sorted(data.Employee.dropna().unique())
    name=st.selectbox("Select employee",names)
    ed=data[data.Employee==name].copy()
    ss=summary_for(ed).iloc[0]
    st.markdown(f"### 👤 {name}")
    cols=st.columns(4)
    for i,k in enumerate(["Total Work Duration","Total OT","Present","WOP"]):
        cols[i].metric(k,str(ss[k]))
    ed["Date"]=ed["Date"].dt.strftime("%Y-%m-%d")
    st.dataframe(ed[["Date","Status","In Time","Out Time","Work Duration","Regular OT","Late By","Early By","WOP OT Excluded"]],use_container_width=True,hide_index=True)

with tab4:
    st.markdown("### 📤 Export")
    period=st.radio("Report period",["Daily","Monthly"],horizontal=True)
    emp_choice=st.radio("Employees",["All Employees","Selected Employees"],horizontal=True)
    if emp_choice=="Selected Employees":
        selected=st.multiselect("Select employees",sorted(data.Employee.unique()))
        export_data=data[data.Employee.isin(selected)].copy()
    else:
        export_data=data.copy()

    if period=="Daily":
        dates=sorted([d for d in data.Date.dropna().dt.strftime("%Y-%m-%d").unique()])
        selected_date=st.selectbox("Select date",dates) if dates else None
        if selected_date:
            export_data=export_data[export_data.Date.dt.strftime("%Y-%m-%d")==selected_date]
    else:
        months=sorted(data.Date.dropna().dt.to_period("M").astype(str).unique())
        selected_month=st.selectbox("Select month",months) if months else None
        if selected_month:
            export_data=export_data[export_data.Date.dt.to_period("M").astype(str)==selected_month]

    fmt=st.radio("Export format",["PDF","Word","Image"],horizontal=True)

    summary=summary_for(export_data)
    if summary.empty:
        st.warning("No records match the selected filters.")
    else:
        st.dataframe(summary,use_container_width=True,hide_index=True)
        if fmt=="PDF":
            try:
                from reportlab.lib import colors
                from reportlab.lib.pagesizes import landscape, A4
                from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
                from reportlab.lib.styles import getSampleStyleSheet
                bio=io.BytesIO()
                doc=SimpleDocTemplate(bio,pagesize=landscape(A4))
                styles=getSampleStyleSheet()
                elements=[Paragraph("Work Time Analyzer Report",styles["Title"])]
                vals=[list(summary.columns)]+summary.astype(str).values.tolist()
                t=Table(vals,repeatRows=1)
                t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365d")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.25,colors.grey),("FONTSIZE",(0,0),(-1,-1),6)]))
                elements.append(t); doc.build(elements)
                st.download_button("⬇️ Download PDF",bio.getvalue(),file_name="work_time_report.pdf",mime="application/pdf")
            except Exception as e: st.error(f"PDF export unavailable: {e}")
        elif fmt=="Word":
            try:
                from docx import Document
                bio=io.BytesIO(); doc=Document(); doc.add_heading("Work Time Analyzer Report",0)
                t=doc.add_table(rows=1,cols=len(summary.columns))
                for i,c in enumerate(summary.columns): t.rows[0].cells[i].text=str(c)
                for _,r in summary.iterrows():
                    cells=t.add_row().cells
                    for i,c in enumerate(summary.columns): cells[i].text=str(r[c])
                doc.save(bio)
                st.download_button("⬇️ Download Word",bio.getvalue(),file_name="work_time_report.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e: st.error(f"Word export unavailable: {e}")
        else:
            try:
                import matplotlib.pyplot as plt
                from PIL import Image
                fig,ax=plt.subplots(figsize=(18,max(4,len(summary)*.45)))
                ax.axis("off")
                tbl=ax.table(cellText=summary.astype(str).values,colLabels=summary.columns,loc="center")
                tbl.auto_set_font_size(False); tbl.set_fontsize(7); tbl.scale(1,1.4)
                bio=io.BytesIO(); fig.savefig(bio,format="png",dpi=180,bbox_inches="tight"); plt.close(fig)
                st.download_button("⬇️ Download Image",bio.getvalue(),file_name="work_time_report.png",mime="image/png")
            except Exception as e: st.error(f"Image export unavailable: {e}")

st.caption("Work Time Analyzer • P=Present • A=Absent • WOP=Saturday working day • WO=Sunday weekly off • WOP OT excluded from regular OT")
