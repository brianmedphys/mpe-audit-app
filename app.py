import streamlit as st
import json
import pandas as pd
import io
from docx import Document
from datetime import datetime

# ==============================================================================
# KSC DATASET (TABLE A, TABLE B, TABLE C2)
# ==============================================================================
KSC_DATA = [
    {
        "table": "Table A", "code": "A1.1", "title": "Radiation Physics & Matter Interactions",
        "description": "Comprehensive understanding of atomic/nuclear physics, radiation production, photon/electron/hadron interactions with matter, and attenuation physics in tissue."
    },
    {
        "table": "Table A", "code": "A1.2", "title": "Radiation Dosimetry Principles & Cavity Theory",
        "description": "In-depth knowledge of cavity theory (Spencer-Attix, Bragg-Gray), absolute and relative dosimetry instrumentation (ionization chambers, diodes, MOSFETs, film, gels)."
    },
    {
        "table": "Table A", "code": "A1.3", "title": "Radiation Detection & Measurement Instrumentation",
        "description": "Operating principles, calibration techniques, energy dependence, directional response, and signal processing of radiation detection systems."
    },
    {
        "table": "Table A", "code": "A2.1", "title": "Legislative & Regulatory Frameworks",
        "description": "Detailed mastery of Irish radiation protection legislation (S.I. 256 of 2018, S.I. 30 of 2019), EU Directives (2013/59/Euratom), and HIQA framework compliance."
    },
    {
        "table": "Table A", "code": "A2.2", "title": "Radiation Safety Administration & Shielding Design",
        "description": "Design and calculation of structural radiation shielding for high-energy linacs (NCRP 151) and brachytherapy suites; safety committee governance and hazard management."
    },
    {
        "table": "Table A", "code": "A3.1", "title": "Anatomy, Physiology & Medical Imaging Physics",
        "description": "Applied knowledge of cross-sectional human anatomy, radio-pathology, and medical imaging modalities (CT, MRI, PET, SPECT) used in radiotherapy planning."
    },
    {
        "table": "Table A", "code": "A3.2", "title": "Biological Effects of Radiation & Radiobiology",
        "description": "Linear-Quadratic model physics, 5 Rs of radiobiology, cell survival curves, fractionations, TCP/NTCP modeling, and acute/late tissue toxicity mechanisms."
    },
    {
        "table": "Table A", "code": "A4.1", "title": "Medical Informatics, Networking & Cybersecurity",
        "description": "DICOM/DICOM-RT standards, PACS/OIS database management, system networking, data security, HIPAA/GDPR compliance, and algorithm software validation."
    },
    {
        "table": "Table B", "code": "B1.1", "title": "Clinical Governance & Quality Management Systems",
        "description": "Implementation of comprehensive departmental QMS, ISO quality standards, clinical audit protocols, and document control procedures."
    },
    {
        "table": "Table B", "code": "B1.2", "title": "Risk Management & Patient Safety",
        "description": "Proactive risk assessment methodologies (FMEA, FTA), incident reporting systems (SAFR, ROSIS), root cause analysis (RCA), and human factors engineering."
    },
    {
        "table": "Table B", "code": "B2.1", "title": "Ethics, Patient Autonomy & Bioethics",
        "description": "Adherence to medical ethics, patient confidentiality, informed consent frameworks in medical exposures, and professional code of conduct."
    },
    {
        "table": "Table B", "code": "B2.2", "title": "Communication, Leadership & Interdisciplinary Collaboration",
        "description": "Effective technical and clinical communication with radiation oncologists, radiation therapists, nurses, hospital management, and regulatory authorities."
    },
    {
        "table": "Table B", "code": "B3.1", "title": "Research, Innovation & Health Technology Assessment",
        "description": "Design of clinical research protocols, statistical analysis, peer-reviewed scientific publishing, evaluation of emerging technologies, and cost-benefit analysis."
    },
    {
        "table": "Table B", "code": "B3.2", "title": "Continuing Professional Development (CPD) & Education",
        "description": "Maintaining personal CPD portfolios, delivering technical training to multidisciplinary teams, and mentoring junior physicists and trainees."
    },
    {
        "table": "Table C2", "code": "C2.1", "title": "External Beam Equipment Commissioning & QA",
        "description": "Acceptance testing, beam data commissioning, and routine quality assurance (daily, monthly, annual) of C-arm linacs, specialized platforms (CyberKnife, Halcyon, MR-Linac), and kilovoltage units."
    },
    {
        "table": "Table C2", "code": "C2.2", "title": "Reference Dosimetry & International Codes of Practice",
        "description": "Absolute dose determination using IAEA TRS-398 or AAPM TG-51 protocols for megavoltage photon and electron beams under reference conditions."
    },
    {
        "table": "Table C2", "code": "C2.3", "title": "Treatment Planning System (TPS) Commissioning & Modeling",
        "description": "Beam modeling in TPS, algorithm validation (AAA, Collapsed Cone Convolution, Monte Carlo, Acuros XB), grid size sensitivity, and heterogeneous medium dose calculation."
    },
    {
        "table": "Table C2", "code": "C2.4", "title": "Advanced Delivery Techniques (IMRT, VMAT, SBRT/SRS)",
        "description": "Planning, optimization, inverse planning parameters, leaf sequencing, small field dosimetry (IAEA TRS-483), and high-dose stereotactic delivery techniques."
    },
    {
        "table": "Table C2", "code": "C2.5", "title": "Patient-Specific Quality Assurance (PSQA)",
        "description": "Pre-treatment verification using 2D/3D detector arrays, EPID dosimetry, portal dose image prediction, and gamma index analysis (3%/2mm, 2%/2mm criteria)."
    },
    {
        "table": "Table C2", "code": "C2.6", "title": "Image-Guided Radiotherapy (IGRT) & Surface-Guided RT (SGRT)",
        "description": "Commissioning and QA of 2D/3D in-room imaging systems (kV/MV CBCT, planar imaging), optical surface tracking, respiratory gating, and motion management."
    },
    {
        "table": "Table C2", "code": "C2.7", "title": "Brachytherapy Equipment, Treatment Planning & Physics",
        "description": "High-Dose-Rate (HDR) and Low-Dose-Rate (LDR) source handling, TG-43 dosimetry formalism, applicator reconstruction, dwell time optimization, emergency procedures, and source exchange calibration."
    },
    {
        "table": "Table C2", "code": "C2.8", "title": "Unintended & Accidental Medical Exposures in Radiotherapy",
        "description": "Safety interlocks, error prevention in treatment delivery, dose miscalculation investigations, reporting protocols to national competent authorities (HIQA), and preventive actions."
    }
]

RATING_OPTIONS = [
    "5 - Master: Departmental lead, innovator, national subject authority",
    "4 - Expert: Independent clinical expert, routine performer, supervisor level",
    "3 - Proficient: Competent execution under general supervision",
    "2 - Foundation: Theoretical knowledge, limited clinical practice",
    "1 - In Development: Gap identified; active clinical training required"
]

# ==============================================================================
# APPLICATION STATE & HELPER FUNCTIONS
# ==============================================================================
def init_state():
    """Initializes Streamlit session state."""
    if "candidate_name" not in st.session_state: st.session_state.candidate_name = ""
    if "hospital" not in st.session_state: st.session_state.hospital = ""
    if "supervisors" not in st.session_state: st.session_state.supervisors = ""
    if "answers" not in st.session_state: st.session_state.answers = {}
    if "current_idx" not in st.session_state: st.session_state.current_idx = 0

def reset_state():
    """Clears the session state to start over."""
    st.session_state.candidate_name = ""
    st.session_state.hospital = ""
    st.session_state.supervisors = ""
    st.session_state.answers = {}
    st.session_state.current_idx = 0

def handle_backup_upload():
    """Callback function triggered ONLY when a JSON backup is uploaded."""
    uploaded_file = st.session_state.get("backup_uploader")
    if uploaded_file is not None:
        try:
            data = json.load(uploaded_file)
            st.session_state.candidate_name = data.get("candidate_name", "")
            st.session_state.hospital = data.get("hospital", "")
            st.session_state.supervisors = data.get("supervisors", "")
            st.session_state.answers = data.get("answers", {})
            st.session_state.current_idx = data.get("current_idx", 0)
            st.toast("Progress restored successfully!", icon="✅")
        except Exception:
            st.error("Failed to parse JSON backup file.")

def generate_word_report():
    """Generates the Section 4b Word document based on user inputs."""
    doc = Document()
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    doc.add_heading("MPE Equivalence Demonstration Audit Report (Section 4b)", level=0)
    
    doc.add_paragraph("Specialty Domain: Radiotherapy Physics")
    doc.add_paragraph(f"Generated Date: {today_str}")
    
    doc.add_heading("1. Candidate & Clinical Infrastructure", level=1)
    doc.add_paragraph(f"Candidate Name: {st.session_state.candidate_name or 'N/A'}", style="List Bullet")
    doc.add_paragraph(f"Hospital / Institution: {st.session_state.hospital or 'N/A'}", style="List Bullet")
    doc.add_paragraph(f"Supervising MPE Supporter(s): {st.session_state.supervisors or 'N/A'}", style="List Bullet")
    
    doc.add_heading("2. Detailed Knowledge, Skills & Competency (KSC) Audit Matrix", level=1)
    
    for k in KSC_DATA:
        ans = st.session_state.answers.get(k["code"], {})
        rating = ans.get("rating", "Unrated / Pending Evaluation")
        evidence = ans.get("evidence", "No clinical documentary evidence provided.")
        
        doc.add_heading(f"[{k['code']}] {k['title']}", level=2)
        doc.add_paragraph(f"Table Category: {k['table']}")
        doc.add_paragraph(f"Requirement Scope: {k['description']}")
        doc.add_paragraph(f"Assessed Competency Level: {rating}")
        doc.add_paragraph("Clinical Documentary Evidence:")
        doc.add_paragraph(evidence, style="Quote")

    doc.add_heading("3. Applicant & Supervisor Declarations", level=1)
    doc.add_paragraph("I confirm that the ratings and clinical documentary evidence detailed in this cross-reference report accurately reflect my professional experience and practice.\n")
    
    doc.add_paragraph(f"Candidate Signature: ___________________________    Date: {today_str}")
    doc.add_paragraph("Supervising MPE Signature: ______________________    Date: _______________")
    
    # Save the document to an in-memory buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==============================================================================
# MAIN APPLICATION LAYOUT
# ==============================================================================
st.set_page_config(page_title="MPE Radiotherapy Assessment", layout="wide")
init_state()

# Calculate Progress 
completed_count = sum(1 for k in KSC_DATA if k["code"] in st.session_state.answers and st.session_state.answers[k["code"]].get("rating"))
total_count = len(KSC_DATA)
progress_pct = completed_count / total_count if total_count > 0 else 0

# Sidebar Configuration
with st.sidebar:
    st.header("MPE Equivalence")
    st.caption("Section 4b Audit Suite")
    st.divider()
    
    st.subheader("Applicant Profile")
    st.session_state.candidate_name = st.text_input("Candidate Name", value=st.session_state.candidate_name, placeholder="Dr. Jane Doe")
    st.session_state.hospital = st.text_input("Hospital / Institution", value=st.session_state.hospital, placeholder="University Hospital Galway")
    st.session_state.supervisors = st.text_input("Supervising MPE Supporter(s)", value=st.session_state.supervisors, placeholder="Prof. J. Smith, MPE")
    
    st.divider()
    st.metric(label="Total Completion", value=f"{int(progress_pct * 100)}%", delta=f"{completed_count} of {total_count} rated", delta_color="normal")
    st.progress(progress_pct)
    
    st.divider()
    st.subheader("Save & Restore Progress")
    
    # 1. Export Current State
    export_data = json.dumps({
        "candidate_name": st.session_state.candidate_name,
        "hospital": st.session_state.hospital,
        "supervisors": st.session_state.supervisors,
        "answers": st.session_state.answers,
        "current_idx": st.session_state.current_idx
    }, indent=2)
    
    st.download_button(
        "Download Backup JSON", 
        data=export_data, 
        file_name="MPE_Backup.json", 
        mime="application/json",
        use_container_width=True
    )
    
    # 2. Import Saved State via Callback
    st.file_uploader(
        "Restore from Backup", 
        type=["json"], 
        key="backup_uploader", 
        on_change=handle_backup_upload
    )
            
    st.divider()
    if st.button("Reset All Data", type="secondary", use_container_width=True):
        reset_state()
        st.rerun()

# Native Tab Navigation
tab1, tab2, tab3 = st.tabs(["Sequential Evaluation", "Competency Matrix", "Section 4b Audit Report"])

# ------------------------------------------------------------------------------
# TAB 1: SEQUENTIAL EVALUATION
# ------------------------------------------------------------------------------
with tab1:
    col_jump, col_blank = st.columns([2, 1])
    with col_jump:
        jump_options = [f"[{k['code']}] {k['title']}" for k in KSC_DATA]
        selected_jump = st.selectbox(
            "Jump to Requirement", 
            options=range(len(jump_options)), 
            format_func=lambda x: jump_options[x], 
            index=st.session_state.current_idx
        )
        if selected_jump != st.session_state.current_idx:
            st.session_state.current_idx = selected_jump
            st.rerun()

    current_ksc = KSC_DATA[st.session_state.current_idx]
    current_code = current_ksc["code"]
    current_ans = st.session_state.answers.get(current_code, {"rating": None, "evidence": ""})

    col1, col2, col3 = st.columns(3)
    col1.metric("Specialty Domain", "Radiotherapy Physics")
    col2.metric("Current Requirement", f"{current_code} ({st.session_state.current_idx + 1} of {len(KSC_DATA)})")
    col3.metric("Table Section", current_ksc["table"])

    st.subheader(f"[{current_code}] {current_ksc['title']}")
    st.info(current_ksc["description"])

    def save_evaluation():
        rating_val = st.session_state.get(f"f_rating_{current_code}")
        evidence_val = st.session_state.get(f"f_evidence_{current_code}")
        
        st.session_state.answers[current_code] = {
            "rating": rating_val.split(":")[0] if rating_val else None,
            "evidence": evidence_val
        }
    
    # Map stored short rating back to the full dropdown string
    default_rating_idx = 0
    if current_ans["rating"]:
        for i, opt in enumerate(RATING_OPTIONS):
            if current_ans["rating"] in opt:
                default_rating_idx = i
                break

    st.radio(
        "1. Self-Assessed Competency Level", 
        options=RATING_OPTIONS, 
        index=default_rating_idx, 
        key=f"f_rating_{current_code}", 
        on_change=save_evaluation
    )
    
    st.text_area(
        "2. Clinical Documentary Evidence & Portfolio References (Section 4b)", 
        value=current_ans["evidence"], 
        height=150, 
        key=f"f_evidence_{current_code}", 
        on_change=save_evaluation, 
        placeholder="Detail specific clinical tasks, TPS modeling, TRS-398 calibrations, QA protocols, or published reports..."
    )

    # Navigation Controls
    c1, c2, c3 = st.columns([1, 8, 1])
    
    with c1:
        if st.session_state.current_idx > 0:
            if st.button("Previous", use_container_width=True):
                st.session_state.current_idx -= 1
                st.rerun()
                
    with c3:
        if st.session_state.current_idx < len(KSC_DATA) - 1:
            if st.button("Next Item", type="primary", use_container_width=True):
                st.session_state.current_idx += 1
                st.rerun()

# ------------------------------------------------------------------------------
# TAB 2: COMPETENCY MATRIX
# ------------------------------------------------------------------------------
with tab2:
    st.subheader("Competency Overview Matrix")
    st.caption("Tables A, B & C2 Summary")
    
    matrix_data = []
    for k in KSC_DATA:
        ans = st.session_state.answers.get(k["code"], {})
        matrix_data.append({
            "Code": k["code"],
            "Requirement Title": k["title"],
            "Self-Rating": ans.get("rating", "Unrated"),
            "Clinical Evidence": ans.get("evidence", "No evidence recorded.")
        })
    
    df = pd.DataFrame(matrix_data)
    st.dataframe(df, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# TAB 3: SECTION 4B AUDIT REPORT
# ------------------------------------------------------------------------------
with tab3:
    st.subheader("Section 4b Audit Summary Report")
    
    word_buffer = generate_word_report()
    
    st.download_button(
        label="Download Word Audit File",
        data=word_buffer,
        file_name=f"MPE_Section4b_Audit_Report_{st.session_state.candidate_name.replace(' ', '_') or 'Draft'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )

    st.markdown("*(The report content is compiled directly into the Word document available for download above.)*")
