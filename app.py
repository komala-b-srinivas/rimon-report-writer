"""
Rimon Health - Psychological ASD Evaluation Report Writer
Phase 1 - Built from actual Rimon Health report template
"""
import streamlit as st
from groq import Groq
from dotenv import load_dotenv
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os, json, base64, io, smtplib, subprocess, tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime

load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

st.set_page_config(page_title="Rimon Health - ASD Report Writer", layout="wide")

# ══════════════════════════════════════════════════════════════════════
# PASSWORD GATE
# ══════════════════════════════════════════════════════════════════════
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.markdown("""
        <style>
        .login-box {
            max-width: 420px; margin: 80px auto 0; background: white;
            border-radius: 16px; padding: 48px 40px;
            box-shadow: 0 8px 32px rgba(75,174,232,0.15);
            border-top: 4px solid #4BAEE8; text-align: center;
        }
        .login-box img { height: 72px; margin-bottom: 20px; }
        .login-box h2 { font-size: 20px; font-weight: 700; color: #1a1a2e; margin: 0 0 6px; }
        .login-box p { font-size: 13px; color: #6b7280; margin: 0 0 28px; }
        </style>
        <div class="login-box">
            <img src="https://static.wixstatic.com/media/022991_02a105832a4745979b94f16debb093a8~mv2.png" />
            <h2>Report Writer</h2>
            <p>Enter your access password to continue</p>
        </div>
        """, unsafe_allow_html=True)
        pw = st.text_input("Password", type="password", label_visibility="collapsed", placeholder="Enter password")
        if st.button("Login", use_container_width=True):
            correct = st.secrets.get("APP_PASSWORD", os.getenv("APP_PASSWORD", ""))
            if pw == correct:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password.")
        st.stop()

check_password()

# ══════════════════════════════════════════════════════════════════════
# SCORE HELPERS
# ══════════════════════════════════════════════════════════════════════

SS_PCT = {
    40:0,44:0,45:0,50:0,55:0,60:1,65:1,68:2,69:2,70:2,72:3,74:4,75:5,
    76:5,78:7,80:9,82:12,84:14,85:16,87:19,88:21,90:25,92:30,95:37,
    98:45,100:50,102:55,105:63,108:70,110:75,112:79,115:84,118:88,
    120:91,122:93,125:95,128:97,130:98,135:99,140:100,145:100,
}

def ss_to_pct(ss):
    ss = max(40, min(145, int(round(ss))))
    if ss in SS_PCT: return SS_PCT[ss]
    keys = sorted(SS_PCT)
    for i in range(len(keys)-1):
        if keys[i] < ss < keys[i+1]:
            lo, hi = keys[i], keys[i+1]
            f = (ss - lo) / (hi - lo)
            return round(SS_PCT[lo] + f * (SS_PCT[hi] - SS_PCT[lo]))
    return 50

def ss_adaptive_level(ss):
    if ss >= 130: return "High"
    if ss >= 115: return "Moderately High"
    if ss >= 86:  return "Adequate"
    if ss >= 71:  return "Moderately Low"
    return "Low"

def t_classification(t, adaptive=False):
    """Adaptive scales flag low scores (opposite direction)."""
    if adaptive:
        if t <= 30: return "Clinically Significant"
        if t <= 40: return "At-Risk"
        return "Adequate"
    else:
        if t >= 70: return "Clinically Significant"
        if t >= 60: return "At-Risk"
        return "Average"

def t_flag(t, adaptive=False):
    c = t_classification(t, adaptive)
    if c == "Clinically Significant": return "red"
    if c == "At-Risk": return "orange"
    return "normal"

# ══════════════════════════════════════════════════════════════════════
# ADOS-2 AUTO-CLASSIFY
# ══════════════════════════════════════════════════════════════════════

ADOS_THRESHOLDS = {
    "Toddler (T)": {"sa_autism": 12, "sa_spectrum": 8, "combined_autism": 17, "combined_spectrum": 12},
    "Module 1":    {"sa_autism": 12, "sa_spectrum": 8,  "combined_autism": 16, "combined_spectrum": 11},
    "Module 2":    {"sa_autism": 9,  "sa_spectrum": 7,  "combined_autism": 13, "combined_spectrum": 8},
    "Module 3":    {"sa_autism": 8,  "sa_spectrum": 6,  "combined_autism": 13, "combined_spectrum": 7},
    "Module 4":    {"sa_autism": 8,  "sa_spectrum": 6,  "combined_autism": 13, "combined_spectrum": 7},
}

def ados_classify(module, sa, rrb):
    combined = sa + rrb
    t = ADOS_THRESHOLDS.get(module, ADOS_THRESHOLDS["Module 1"])
    if combined >= t["combined_autism"] or sa >= t["sa_autism"]:
        return "Autism"
    if combined >= t["combined_spectrum"] or sa >= t["sa_spectrum"]:
        return "Autism Spectrum"
    return "Non-Spectrum"

def ados_sa_label(module, sa):
    t = ADOS_THRESHOLDS.get(module, ADOS_THRESHOLDS["Module 1"])
    if sa >= t["sa_autism"]:   return "Autism"
    if sa >= t["sa_spectrum"]: return "Autism Spectrum"
    return "Non-Spectrum"

def ados_rrb_label(rrb):
    # RRB ≥ 2 generally considered elevated
    return "Elevated" if rrb >= 2 else "Non-Spectrum"

def comparison_score_label(cs):
    if cs <= 3: return f"Level {cs} = Minimal to Low Symptoms"
    if cs <= 5: return f"Level {cs} = Moderate Symptoms"
    if cs <= 7: return f"Level {cs} = High Symptoms"
    return f"Level {cs} = High Symptoms"

# ══════════════════════════════════════════════════════════════════════
# NARRATIVE GENERATORS  (rule-based - no LLM needed for these sections)
# ══════════════════════════════════════════════════════════════════════

BASC3_INTERP = {
    # Clinical scales - high T = problem
    "Hyperactivity":           ("engages in many disruptive, impulsive, and uncontrolled behaviors.",
                                "tends not to be excessively active, impulsive, or disruptive."),
    "Aggression":              ("acts aggressively and can be difficult to control.",
                                "tends not to act aggressively any more often than others of the same age."),
    "Conduct Problems":        ("exhibits behavioral problems that may be in violation of rules or the rights of others.",
                                "demonstrates rule-breaking behavior no more often than others of the same age."),
    "Anxiety":                 ("displays a significant number of anxiety-based behaviors.",
                                "displays relatively few anxiety-based behaviors compared to others of the same age."),
    "Depression":              ("displays a significant number of depressive behaviors.",
                                "displays depressive behaviors no more often than others of the same age."),
    "Somatization":            ("complains of health-related problems to a much greater degree than others of the same age.",
                                "complains of health-related problems to about the same degree as others of the same age."),
    "Atypicality":             ("engages in behaviors that are considered strange or odd, and generally seems disconnected from their surroundings.",
                                "does not display significantly unusual or atypical behaviors compared to others of the same age."),
    "Withdrawal":              ("is seemingly alone, has difficulty making friends, and/or is sometimes unwilling to join group activities.",
                                "does not appear to withdraw from social contact more than others of the same age."),
    "Attention Problems":      ("has significant difficulty maintaining necessary levels of attention. The problems experienced are probably interfering with academic performance and functioning in other areas.",
                                "does not appear to have significantly more difficulty with attention than others of the same age."),
    # Adaptive scales - low T = problem
    "Adaptability":            ("has difficulty adapting to a variety of situations.",
                                "is able to adapt as well as most others of the same age to a variety of situations."),
    "Social Skills":           ("has difficulty complimenting others and making suggestions for improvement in a tactful and socially acceptable manner.",
                                "demonstrates adequate social skills comparable to others of the same age."),
    "Leadership":              ("has difficulty making decisions, lacks creativity, and/or has difficulty getting others to work together effectively.",
                                "demonstrates leadership skills comparable to others of the same age."),
    "Activities of Daily Living": ("has difficulty performing simple daily tasks in a safe and efficient manner.",
                                   "is able to perform daily living tasks at an age-appropriate level."),
    "Functional Communication":("demonstrates unusually poor expressive and receptive communication skills and has significant difficulty seeking out and finding information on their own.",
                                "demonstrates functional communication skills comparable to others of the same age."),
}

def generate_basc3_narrative(data, patient_name, respondent_name):
    """
    Generate the BASC-3 section narrative programmatically.
    data = dict with all composite and subscale T + percentile + CI values.
    Matches Q-Global report language exactly.
    """
    rater = respondent_name or "the respondent"
    name  = patient_name or "The patient"

    def scale_sentence(scale, t, pct, adaptive=False):
        cls  = t_classification(t, adaptive)
        hi_t = t >= 60 and not adaptive
        lo_t = t <= 40 and adaptive
        flagged = hi_t or lo_t
        interp_hi, interp_lo = BASC3_INTERP.get(scale, ("displays notable behaviors in this area.", "does not display notable behaviors in this area."))
        interp = interp_hi if flagged else interp_lo
        warn = " This T score falls in the Clinically Significant classification range and usually warrants follow-up." if cls == "Clinically Significant" else \
               " This T score falls in the At-Risk classification range and follow-up may be necessary." if cls == "At-Risk" else ""
        return (f"{name}'s T score on {scale} is {t} and has a percentile rank of {pct}.{warn} "
                f"{rater.capitalize()} reports {name} {interp}")

    def composite_sentence(label, t, ci_lo, ci_hi, pct, adaptive=False):
        cls = t_classification(t, adaptive)
        warn = f" {name}'s T score on this composite scale falls in the Clinically Significant classification range." if cls == "Clinically Significant" else ""
        return (f"The {label} composite scale T score is {t}, with a 90% confidence interval range of "
                f"{ci_lo}-{ci_hi} and a percentile rank of {pct}.{warn}")

    lines = []

    # Externalizing
    lines.append(composite_sentence("Externalizing Problems",
        data["ext_t"], data["ext_ci_lo"], data["ext_ci_hi"], data["ext_pct"]))
    for s in ["Hyperactivity", "Aggression", "Conduct Problems"]:
        key = s.lower().replace(" ", "_")
        lines.append(scale_sentence(s, data[f"{key}_t"], data[f"{key}_pct"]))
    lines.append("")

    # Internalizing
    lines.append(composite_sentence("Internalizing Problems",
        data["int_t"], data["int_ci_lo"], data["int_ci_hi"], data["int_pct"]))
    for s in ["Anxiety", "Depression", "Somatization"]:
        key = s.lower()
        lines.append(scale_sentence(s, data[f"{key}_t"], data[f"{key}_pct"]))
    lines.append("")

    # BSI
    lines.append(composite_sentence("Behavioral Symptoms Index (BSI)",
        data["bsi_t"], data["bsi_ci_lo"], data["bsi_ci_hi"], data["bsi_pct"]))
    lines.append("Scale summary information for Hyperactivity, Aggression, and Depression (scales included in the BSI) has been provided above. Scale summary information for the remaining BSI scales is given next.")
    for s in ["Atypicality", "Withdrawal", "Attention Problems"]:
        key = s.lower().replace(" ", "_")
        lines.append(scale_sentence(s, data[f"{key}_t"], data[f"{key}_pct"]))
    lines.append("")

    # Adaptive Skills
    lines.append(composite_sentence("Adaptive Skills",
        data["adp_t"], data["adp_ci_lo"], data["adp_ci_hi"], data["adp_pct"], adaptive=True))
    for s in ["Adaptability", "Social Skills", "Leadership", "Activities of Daily Living", "Functional Communication"]:
        key = s.lower().replace(" ", "_")
        lines.append(scale_sentence(s, data[f"{key}_t"], data[f"{key}_pct"], adaptive=True))

    return "\n".join(lines)


def generate_vineland_narrative(abc, comm, daily, social, comm_pct, daily_pct, social_pct,
                                 patient_name, respondent_name, date_completed):
    name = patient_name or "The patient"
    rater = respondent_name or "the caregiver"
    abc_pct = ss_to_pct(abc)
    abc_pct_str = f"<1" if abc_pct == 0 else str(abc_pct)
    comm_pct_str = f"<1" if comm_pct == 0 else str(comm_pct)
    daily_pct_str = f"<1" if daily_pct == 0 else str(daily_pct)
    social_pct_str = f"<1" if social_pct == 0 else str(social_pct)

    comm_rel  = "relative weakness" if comm < min(daily, social) else ("relative strength" if comm > max(daily, social) else "within the same range as other domains")
    daily_note = ""
    social_note = ""

    text = f"""{name} was evaluated using the Vineland-3 Comprehensive Parent/Caregiver Form on {date_completed}. {rater.capitalize()}, {name}'s caregiver, completed the form.

{name}'s overall level of adaptive functioning is described by the score on the Adaptive Behavior Composite (ABC). The ABC score is {abc}, which is {"well " if abc < 70 else ""}{"below" if abc < 85 else "within"} the normative mean of 100 (the normative standard deviation is 15). The percentile rank for this overall score is {abc_pct_str}. The ABC score is based on scores for three specific adaptive behavior domains: Communication, Daily Living Skills, and Socialization.

The Communication domain measures how well {name} listens and understands, expresses themselves through speech, and reads and writes. The Communication standard score is {comm}. This corresponds to a percentile rank of {comm_pct_str}. This domain is a {comm_rel} for {name}.

The Daily Living Skills domain assesses {name}'s performance of the practical, everyday tasks of living that are appropriate for age. The standard score for Daily Living Skills is {daily}, which corresponds to a percentile rank of {daily_pct_str}.

{name}'s score for the Socialization domain reflects functioning in social situations. The Socialization standard score is {social}. The percentile rank is {social_pct_str}.

{name} fell in the {ss_adaptive_level(abc)} Range across all domain Standard Scores. This placed the overall Adaptive Behavior Composite Score into the {ss_adaptive_level(abc)} range."""
    return text


def generate_ados_narrative(module, sa, rrb, comparison_score, classification,
                             dsm5_a_met, dsm5_a_age, dsm5_b_desc, dsm5_c_desc,
                             intellectual_impairment, language_impairment,
                             criteria_a_level, criteria_b_level, patient_name):
    combined = sa + rrb
    sa_label  = ados_sa_label(module, sa)
    rrb_label = ados_rrb_label(rrb)
    cs_label  = comparison_score_label(comparison_score)
    name = patient_name or "The patient"

    text = f"""The ADOS-2 is a semi-structured, standardized assessment of communication, social behavior and play, which is important in the diagnosis and evaluation of autism and other autism spectrum disorders. The ADOS-2 consists of four modules of standard activities designed to be used with children and adults. The focus of the ADOS-2 is to conduct observations of social behavior and communication rather than to assess specific cognitive abilities and academic skills.

{module} was selected for this evaluation because it is designed for use with clients who demonstrate {"low verbal ability, only stating one word and unable to form 2-3 word phrases or requests" if module in ["Module 1","Toddler (T)"] else "phrase speech and emerging verbal fluency" if module == "Module 2" else "verbally fluent abilities"}. {module} consists of 10 activities with accompanying ratings. The activities focus on playful/imaginative use of objects, initiation and response to social interaction with the examiner, joint attention, and the quality of shared enjoyment.

Ratings of behavior are organized into three groupings: Communication, Reciprocal Social Interaction, and Restricted and Repetitive Behavior. Observations are scored on a 0-3 continuum, ranging from no evidence of deficit to significant atypicality.

ADOS Ratings Areas:                              Domain Total         Autism Related Symptoms
Social Affect (SA):                              {sa:<30}{sa_label}
Restricted and Repetitive Behavior (RRB):        {rrb:<30}{rrb_label}
Algorithm Combined Total (SA) and (RRB):         {combined:<30}{cs_label}
ADOS-2 Classification:

Based upon the Criteria for the ADOS, the classification of {classification} {"has been met" if classification != "Non-Spectrum" else "has not been met"}.

Symptoms must be present in the early developmental period (but may not become fully manifest until social demands exceed limited capacities, or may be masked by learned strategies in later life):
           {"This criteria was met when " + name + " was " + str(dsm5_a_age) + " years old." if dsm5_a_met else "This criteria was not clearly established based on available history."}

Symptoms cause clinically significant impairment in social, occupational, or other important areas of current function:
{dsm5_b_desc or "Global impact across all areas."}

The disturbance is not better accounted for by intellectual disability or global developmental delay:
{dsm5_c_desc or "There is evidence of an intellectual disability and global developmental delay. Cognitive ability is impacted by receptive and expressive speech processing."}

Specifiers:
Intellectual Impairment: {"Yes. There is Intellectual Impairment." if intellectual_impairment else "No. There is no intellectual impairment."}
Language Impairment: {language_impairment or "Not specified."}

Criteria A - Social Communication Severity: Level {criteria_a_level} - {"Requiring Very Substantial Support" if criteria_a_level == 3 else "Requiring Substantial Support" if criteria_a_level == 2 else "Requiring Support"}
Criteria B - Restricted / Repetitive Behaviors Severity: Level {criteria_b_level} - {"Requiring Very Substantial Support" if criteria_b_level == 3 else "Requiring Substantial Support" if criteria_b_level == 2 else "Requiring Support"}"""
    return text


# ══════════════════════════════════════════════════════════════════════
# AUDIO / IMAGE EXTRACTION  (reused from prior build)
# ══════════════════════════════════════════════════════════════════════

WALK_OPTS   = ["On time (~12 months)","Delayed (>15 months)","Significantly delayed (>24 months)","Not yet achieved","Unknown"]
WORDS_OPTS  = ["On time (~12 months)","Delayed (>18 months)","Significantly delayed (>24 months)","Not yet achieved","Unknown"]
PHRASE_OPTS = ["On time (~24 months)","Delayed (>30 months)","Significantly delayed (>36 months)","Not yet achieved","Unknown"]
IEP_OPTS    = ["Active IEP","504 Plan","No IEP/504","Under evaluation","Unknown"]
COMM_OPTS   = ["Fully verbal (age-appropriate)","Verbal with delays","Phrase speech (2-4 word combinations)","Single words only","Primarily nonverbal","AAC device / PECS"]
SERVICES_OPTS = ["Speech-Language Therapy (ST)","Occupational Therapy (OT)","Physical Therapy (PT)",
                 "ABA Therapy","Resource Room","1:1 Para Support","SEIT","Counseling","Extended School Year (ESY)","None"]

def opt_idx(opts, val, default=0):
    try: return opts.index(val)
    except ValueError: return default

def transcribe_audio(audio_bytes, filename):
    ext  = filename.rsplit(".", 1)[-1].lower()
    mime = {"mp3":"audio/mpeg","mp4":"audio/mp4","m4a":"audio/mp4","wav":"audio/wav","webm":"audio/webm"}.get(ext,"audio/mpeg")
    return client.audio.transcriptions.create(
        file=(filename, audio_bytes, mime),
        model="whisper-large-v3-turbo",
        response_format="text",
        language="en",
    )

def extract_background_from_transcript(transcript):
    prompt = f"""Extract background information from this session transcript. Return ONLY valid JSON:
{{
  "referral_reason": "string",
  "birth_complications": "string or null",
  "medical_diagnoses": "string or null",
  "current_medications": "string or null",
  "milestone_walking": one of {WALK_OPTS},
  "milestone_words": one of {WORDS_OPTS},
  "milestone_phrases": one of {PHRASE_OPTS},
  "regression": true or false,
  "regression_detail": "string or null",
  "iep_status": one of {IEP_OPTS},
  "classification": "string or null",
  "prior_evals": "string or null",
  "prior_diagnoses": "string or null",
  "services": [],
  "comm_level": one of {COMM_OPTS},
  "home_language": "string",
  "family_hx": "string or null",
  "social_hx": "string or null",
  "school_placement": "string or null",
  "behavioral_concerns": "string or null",
  "toilet_trained": true or false or null,
  "feeding_difficulties": "string or null"
}}
TRANSCRIPT: {transcript[:6000]}"""
    resp = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role":"user","content":prompt}],
        temperature=0.1,
    )
    raw = resp.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"): raw = raw[4:]
    return json.loads(raw.strip())

def extract_scores_from_image(image_bytes, context="neuropsychological evaluation"):
    b64   = base64.b64encode(image_bytes).decode("utf-8")
    sig   = image_bytes[:4]
    mime  = "image/jpeg" if sig[:3]==b'\xff\xd8\xff' else ("image/png" if sig[:4]==b'\x89PNG' else "image/jpeg")
    prompt = f"""You are reading a neuropsychological score sheet from a {context} evaluation.
Extract ALL numerical scores visible. Return ONLY valid JSON:
{{"test_battery": "name", "scores": {{"field name": number}}, "notes": "any issues"}}
For T-scores label as "<scale> T". For standard scores label as "<index> SS" or just the index name.
Return ONLY JSON."""
    resp = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{"role":"user","content":[
            {"type":"image_url","image_url":{"url":f"data:{mime};base64,{b64}"}},
            {"type":"text","text":prompt}
        ]}],
        temperature=0.1, max_tokens=2048,
    )
    raw = resp.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"): raw = raw[4:]
    return json.loads(raw.strip())

def extract_obs_from_image(image_bytes):
    b64  = base64.b64encode(image_bytes).decode("utf-8")
    sig  = image_bytes[:4]
    mime = "image/jpeg" if sig[:3]==b'\xff\xd8\xff' else ("image/png" if sig[:4]==b'\x89PNG' else "image/jpeg")
    resp = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{"role":"user","content":[
            {"type":"image_url","image_url":{"url":f"data:{mime};base64,{b64}"}},
            {"type":"text","text":"Transcribe these handwritten behavioral observation notes and rewrite as polished clinical language for a psychological evaluation report. Only include content that is clearly clinical in nature - observations about the patient's behavior, attention, responses, affect, or test performance. Ignore any annotations that appear to be developer or tester notes (e.g. 'null handling', 'missing score logic', 'useful to test', 'TODO', etc.). Return only the clinical paragraph(s)."}
        ]}],
        temperature=0.2, max_tokens=1024,
    )
    return resp.choices[0].message.content.strip()


# ══════════════════════════════════════════════════════════════════════
# THEME / CUSTOM CSS
# ══════════════════════════════════════════════════════════════════════

st.markdown("""
<style>
/* ── Global ── */
html, body, [class*="css"] {
    font-family: 'Inter', 'Segoe UI', sans-serif;
    background-color: #f7f9fc;
    color: #1a1a2e;
}

/* ── Hide Streamlit chrome ── */
#MainMenu, footer, header { visibility: hidden; }

/* ── Main container ── */
.block-container {
    padding-top: 0 !important;
    max-width: 1100px;
}

/* ── Top header bar ── */
.rh-header {
    background: #ffffff;
    border-bottom: 3px solid #4BAEE8;
    padding: 18px 32px;
    display: flex;
    align-items: center;
    gap: 20px;
    margin-bottom: 28px;
    border-radius: 0 0 12px 12px;
    box-shadow: 0 2px 12px rgba(75,174,232,0.10);
}
.rh-header img { height: 56px; }
.rh-header-text h1 {
    margin: 0;
    font-size: 22px;
    font-weight: 700;
    color: #1a1a2e;
    letter-spacing: -0.3px;
}
.rh-header-text p {
    margin: 2px 0 0;
    font-size: 13px;
    color: #4BAEE8;
    font-weight: 500;
}
.rh-badge {
    margin-left: auto;
    background: #eef7fd;
    color: #4BAEE8;
    border: 1px solid #4BAEE8;
    border-radius: 20px;
    padding: 4px 14px;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.05em;
}

/* ── Section headers ── */
h2, h3 { color: #1a1a2e !important; font-weight: 700 !important; }

/* ── Subheader pill ── */
.stSubheader {
    background: linear-gradient(90deg, #eef7fd, #f7f9fc);
    border-left: 4px solid #4BAEE8;
    padding: 10px 16px !important;
    border-radius: 0 8px 8px 0;
    margin-bottom: 16px !important;
}

/* ── Input fields ── */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea,
.stSelectbox > div > div {
    border: 1px solid #d0e8f7 !important;
    border-radius: 8px !important;
    background: #ffffff !important;
    font-size: 14px !important;
}
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
    border-color: #4BAEE8 !important;
    box-shadow: 0 0 0 3px rgba(75,174,232,0.15) !important;
}

/* ── Buttons ── */
.stButton > button {
    background: #4BAEE8 !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    padding: 10px 24px !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    background: #2a8fd4 !important;
    box-shadow: 0 4px 12px rgba(75,174,232,0.35) !important;
    transform: translateY(-1px) !important;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: #eef7fd;
    border-radius: 10px;
    padding: 4px;
    gap: 4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px !important;
    font-weight: 500 !important;
    font-size: 13px !important;
    color: #4BAEE8 !important;
}
.stTabs [aria-selected="true"] {
    background: #4BAEE8 !important;
    color: white !important;
}

/* ── Cards (expanders) ── */
.streamlit-expanderHeader {
    background: #f0f8fe !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    color: #1a1a2e !important;
}
.streamlit-expanderContent {
    border: 1px solid #d0e8f7 !important;
    border-top: none !important;
    border-radius: 0 0 8px 8px !important;
    background: #ffffff !important;
}

/* ── Info / success / error boxes ── */
.stAlert {
    border-radius: 8px !important;
    font-size: 13px !important;
}

/* ── Divider ── */
hr { border-color: #d0e8f7 !important; margin: 28px 0 !important; }

/* ── Checkbox ── */
.stCheckbox > label { font-size: 14px !important; color: #1a1a2e !important; }

/* ── Download buttons ── */
.stDownloadButton > button {
    background: #1a1a2e !important;
    color: white !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}
.stDownloadButton > button:hover {
    background: #2d2d4e !important;
}

/* ── Spinner ── */
.stSpinner > div { border-top-color: #4BAEE8 !important; }
</style>
""", unsafe_allow_html=True)

# ── TOP HEADER BAR ──
st.markdown("""
<div class="rh-header">
    <img src="https://static.wixstatic.com/media/022991_02a105832a4745979b94f16debb093a8~mv2.png" />
    <div class="rh-header-text">
        <h1>ASD Evaluation Report Writer</h1>
        <p>Accessible Neuropsych Assessments</p>
    </div>
    <span class="rh-badge">INTERNAL USE ONLY</span>
</div>
""", unsafe_allow_html=True)

st.caption("De-identified inputs only · All drafts require clinician review and signature before release")
st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 1 - HEADER INFO
# ══════════════════════════════════════════════════════════════════════
st.subheader("Patient & Evaluation Info")
st.info("Use first name or initials only. No last name, no full DOB.")

col1, col2 = st.columns(2)
with col1:
    patient_name    = st.text_input("Patient First Name / Initials", placeholder="e.g. M.B. or Mckenzie")
    dob_text        = st.text_input("Date of Birth (year only for HIPAA)", placeholder="e.g. 2020")
    chron_age       = st.text_input("Chronological Age", placeholder="e.g. 5 years, 3 months")
    eval_date       = st.date_input("Date of Evaluation", value=datetime.today())
    language        = st.selectbox("Language of Testing", ["English","Spanish","Both","Other"])

with col2:
    school_name     = st.text_input("School Name", placeholder="e.g. P.S. 169")
    grade_placement = st.text_input("Grade / Classroom Placement", placeholder="e.g. Pre-K, 6:1+1 classroom")
    examiner_name   = st.text_input("Examiner Name + Credentials", placeholder="e.g. Daniella Abekassis M.S.")
    supervisor_name = st.text_input("Supervising Psychologist + Credentials", placeholder="e.g. Dr. Gabrielle Kirby, Psy.D")
    supervisor_npi  = st.text_input("Supervisor NPI", placeholder="e.g. 1598984932")
    supervisor_lic  = st.text_input("Supervisor License #", placeholder="e.g. 01694")

st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 2 - REASON FOR REFERRAL
# ══════════════════════════════════════════════════════════════════════
st.subheader("Reason for Referral")
referral_by      = st.selectbox("Referred by", ["Parent/Guardian","School District","Pediatrician","OPWDD Office","Court/Legal","Self-Referral","Other"])
referral_concern = st.text_area("Presenting concerns", placeholder="e.g. Parent reports possible neurological issues and/or ASD symptoms. Referred to assess if child is on the Autism Spectrum with the attempt of getting services through the OPWDD.", height=80)
st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 3 - TESTS ADMINISTERED
# ══════════════════════════════════════════════════════════════════════
st.subheader("Evaluation Materials")
col1, col2 = st.columns(2)
with col1:
    use_wppsi   = st.checkbox("WPPSI-IV (Cognitive)", value=True)
    use_wisc    = st.checkbox("WISC-V (Cognitive - if age 6+)")
    use_ados    = st.checkbox("ADOS-2", value=True)
with col2:
    use_basc    = st.checkbox("BASC-3 PRS (Behavior - Parent Rating)", value=True)
    use_vineland= st.checkbox("Vineland-3 (Adaptive Behavior - Parent)", value=True)
    use_case    = st.checkbox("Case Materials / Records Review", value=True)
extra_tests = st.text_input("Other tests (comma-separated)", placeholder="e.g. CARS-2, SRS-2")
st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 4 - BACKGROUND
# ══════════════════════════════════════════════════════════════════════
st.subheader("Background Information")
_bg = st.session_state.get("extracted_bg", {})

bg_auto, bg_manual = st.tabs([
    "🎙️  Extract from Session Recording / Transcript  ← start here",
    "✏️  Manual Entry"
])

with bg_auto:
    st.markdown("Upload the intake/background session recording or paste the transcript. AI fills all fields below.")
    audio_mode = st.radio("Input", ["Upload audio file","Paste transcript"], horizontal=True)
    if audio_mode == "Upload audio file":
        af = st.file_uploader("Session recording", type=["mp3","mp4","m4a","wav","webm"])
        if af:
            st.audio(af)
            if st.button("Transcribe Audio", type="primary"):
                with st.spinner("Transcribing..."):
                    try:
                        t = transcribe_audio(af.read(), af.name)
                        st.session_state["transcript"] = t
                        st.success("Done. See transcript below.")
                    except Exception as e:
                        st.error(f"Transcription failed: {e}")
        if "transcript" in st.session_state:
            with st.expander("View transcript"): st.text_area("", st.session_state["transcript"], height=180, disabled=True)
    else:
        pasted = st.text_area("Paste transcript", value=st.session_state.get("transcript",""), height=200)
        if pasted: st.session_state["transcript"] = pasted

    if st.session_state.get("transcript") and st.button("Extract Background from Transcript", type="primary"):
        with st.spinner("Extracting fields..."):
            try:
                ex = extract_background_from_transcript(st.session_state["transcript"])
                st.session_state["extracted_bg"] = ex
                _bg = ex
                st.success("Extracted. Review in Manual Entry tab.")
            except Exception as e:
                st.error(f"Extraction failed: {e}")

with bg_manual:
    st.caption("Pre-filled from transcript if extracted. Edit anything needed.")

    # Prior evaluations
    with st.expander("Prior Evaluations & Diagnoses", expanded=True):
        prior_evals_text = st.text_area("Prior evaluations (date, test, result summary)",
            value=_bg.get("prior_evals",""),
            placeholder="e.g. Nov 2022 - Stanford Binet-5: Mildly Impaired range; CARS-2: Autism Spectrum; Vineland-3: Moderately Low to Low across all domains",
            height=80)
        prior_diagnoses = st.text_input("Prior diagnoses on record",
            value=_bg.get("prior_diagnoses",""),
            placeholder="e.g. Chung-Jansen Syndrome, Intellectual Disability, ASD, Developmental Delay, hip dysplasia")

    # Parent interview
    with st.expander("Parent Interview", expanded=True):
        interview_date = st.text_input("Interview date", value=eval_date.strftime("%m/%d/%Y"))
        col1, col2 = st.columns(2)
        with col1:
            family_composition = st.text_input("Family composition",
                value=_bg.get("social_hx",""),
                placeholder="e.g. Resides with mother, father, and two step-siblings")
            family_hx = st.text_input("Family history of ASD/IDD",
                value=_bg.get("family_hx",""),
                placeholder="e.g. Stepbrother with ASD. No other reported family history.")
            comm_level = st.selectbox("Communication level", COMM_OPTS,
                index=opt_idx(COMM_OPTS, _bg.get("comm_level","Single words only")))
            echolalia  = st.checkbox("Echolalia present", value=True)
            if echolalia:
                echolalia_detail = st.text_input("Echolalia description",
                    placeholder="e.g. Repeats phrases multiple times after hearing them")
            else:
                echolalia_detail = ""
        with col2:
            toilet_trained = st.checkbox("Toilet trained", value=not _bg.get("toilet_trained", False) if _bg.get("toilet_trained") is not None else False)
            feeding_difficulties = st.text_input("Feeding difficulties",
                value=_bg.get("feeding_difficulties",""),
                placeholder="e.g. Recently began self-feeding, does not consume solid foods, milk from bottle")
            behavioral_concerns  = st.text_area("Behavioral concerns at home",
                value=_bg.get("behavioral_concerns",""),
                placeholder="e.g. Head banging, biting self when needs not met, pacing, rocking",
                height=80)

    # Educational history
    with st.expander("Educational History"):
        col1, col2 = st.columns(2)
        with col1:
            school_placement_type = st.text_input("Classroom placement type",
                value=_bg.get("school_placement",""),
                placeholder="e.g. 6:1+1, 8:1+1, 12:1+4, mainstream")
            iep_status = st.selectbox("IEP Status", IEP_OPTS,
                index=opt_idx(IEP_OPTS, _bg.get("iep_status","Active IEP")))
            cse_classification = st.text_input("CSE Classification",
                value=_bg.get("classification",""),
                placeholder="e.g. Autism, Multiple Disabilities, None")
        with col2:
            services = st.multiselect("Current services", SERVICES_OPTS,
                default=[s for s in (_bg.get("services",[]) or []) if s in SERVICES_OPTS])

    # Developmental history
    with st.expander("Developmental & Medical History"):
        col1, col2 = st.columns(2)
        with col1:
            birth_complications = st.text_input("Birth/prenatal complications",
                value=_bg.get("birth_complications",""),
                placeholder="e.g. Premature at 34 weeks, or 'None reported'")
            medical_diagnoses   = st.text_input("Medical diagnoses",
                value=_bg.get("medical_diagnoses",""),
                placeholder="e.g. Chung-Jansen Syndrome, hip dysplasia, gait disorder")
            current_meds        = st.text_input("Current medications",
                value=_bg.get("current_medications",""),
                placeholder="e.g. None, or list medications")
        with col2:
            milestone_walking = st.selectbox("Walking milestone", WALK_OPTS,
                index=opt_idx(WALK_OPTS, _bg.get("milestone_walking","Unknown")))
            milestone_words   = st.selectbox("First words", WORDS_OPTS,
                index=opt_idx(WORDS_OPTS, _bg.get("milestone_words","Unknown")))
            milestone_phrases = st.selectbox("Two-word phrases", PHRASE_OPTS,
                index=opt_idx(PHRASE_OPTS, _bg.get("milestone_phrases","Unknown")))
            regression = st.checkbox("Developmental regression", value=bool(_bg.get("regression",False)))
            if regression:
                regression_detail = st.text_input("Regression description",
                    value=_bg.get("regression_detail",""))
            else:
                regression_detail = ""

st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 5 - BEHAVIORAL OBSERVATIONS
# ══════════════════════════════════════════════════════════════════════
st.subheader("Behavioral Observations")
obs_photo, obs_type = st.tabs(["📷  Upload Photo of Handwritten Notes  ← start here", "✏️  Type / Dictate"])

with obs_photo:
    obs_img = st.file_uploader("Handwritten observation notes", type=["jpg","jpeg","png","webp"], key="obs_upload")
    if obs_img:
        st.image(obs_img, use_container_width=True)
        if st.button("Extract Observations from Photo", type="primary"):
            with st.spinner("Reading handwriting..."):
                try:
                    obs_extracted = extract_obs_from_image(obs_img.read())
                    st.session_state["observations"] = obs_extracted
                    st.success("Done. Review in Type tab.")
                    st.text_area("Preview", obs_extracted, height=150, disabled=True)
                except Exception as e:
                    st.error(f"Failed: {e}")

with obs_type:
    observations = st.text_area(
        "Behavioral observations",
        value=st.session_state.get("observations",""),
        height=200,
        placeholder="e.g. Patient arrived with mother. Minimal eye contact. Echolalic speech noted (counting '1,2,3'). Did not respond to bids for joint attention. Stereotyped hand movements observed. Effort genuine; results valid estimate of current functioning.",
    )
    st.session_state["observations"] = observations

st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 6a - WPPSI-IV / WISC-V
# ══════════════════════════════════════════════════════════════════════
cog_scores = {}
if use_wppsi or use_wisc:
    cog_label = "WPPSI-IV" if use_wppsi else "WISC-V"
    st.subheader(f"Cognitive Assessment - {cog_label}")

    cog_photo, cog_manual = st.tabs(["📷  Upload Score Sheet Photo", "✏️  Manual Entry"])
    _cog = st.session_state.get("extracted_cog_scores", {})

    with cog_photo:
        cog_img = st.file_uploader(f"{cog_label} score sheet", type=["jpg","jpeg","png","webp"], key="cog_upload")
        if cog_img:
            st.image(cog_img, use_container_width=True)
            if st.button(f"Extract {cog_label} Scores", type="primary"):
                with st.spinner("Reading score sheet..."):
                    try:
                        r = extract_scores_from_image(cog_img.read(), cog_label)
                        st.session_state["extracted_cog_scores"] = r.get("scores",{})
                        _cog = st.session_state["extracted_cog_scores"]
                        st.success("Scores extracted. Review in Manual Entry tab.")
                        if r.get("notes"): st.warning(r["notes"])
                    except Exception as e:
                        st.error(f"Extraction failed: {e}")

    with cog_manual:
        obtained = st.radio("Were scores obtained?", ["Yes - full battery administered", "No - unable to obtain scores"], horizontal=True)
        if "No" in obtained:
            cog_not_obtained_reason = st.text_area("Reason scores not obtained",
                placeholder="e.g. Patient was unable to sustain attention and participate during this evaluation. Patient often uttered nonsense words to herself and provided no responses to any questions.",
                height=80)
            cog_scores["obtained"] = False
            cog_scores["reason"]   = cog_not_obtained_reason
        else:
            cog_scores["obtained"] = True
            st.caption("Enter index scores (Standard Score, mean=100, SD=15)")
            col1, col2 = st.columns(2)
            with col1:
                for label, default in [("Full Scale IQ (FSIQ)",85),("Verbal Comprehension Index (VCI)",82),("Visual Spatial Index (VSI)",88)]:
                    v = st.number_input(label, 40, 160, int(_cog.get(label, default)))
                    cog_scores[label] = v
            with col2:
                for label, default in [("Fluid Reasoning Index (FRI)",84),("Working Memory Index (WMI)",80),("Processing Speed Index (PSI)",78)]:
                    v = st.number_input(label, 40, 160, int(_cog.get(label, default)))
                    cog_scores[label] = v

    st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 6b - BASC-3 PRS
# ══════════════════════════════════════════════════════════════════════
basc_data = {}
if use_basc:
    st.subheader("BASC-3 - Behavior Assessment System for Children (Parent Rating Scales)")

    basc_photo, basc_manual = st.tabs(["📷  Upload Score Sheet Photo", "✏️  Manual Entry"])
    _basc = st.session_state.get("extracted_basc_scores", {})

    with basc_photo:
        basc_img = st.file_uploader("BASC-3 PRS score sheet", type=["jpg","jpeg","png","webp"], key="basc_upload")
        if basc_img:
            st.image(basc_img, use_container_width=True)
            if st.button("Extract BASC-3 Scores", type="primary"):
                with st.spinner("Reading BASC-3 scores..."):
                    try:
                        r = extract_scores_from_image(basc_img.read(), "BASC-3 Parent Rating Scales")
                        st.session_state["extracted_basc_scores"] = r.get("scores",{})
                        _basc = st.session_state["extracted_basc_scores"]
                        st.success("Extracted. Review in Manual Entry.")
                    except Exception as e:
                        st.error(f"Failed: {e}")

    with basc_manual:
        b_respondent = st.text_input("Respondent name (parent/caregiver)", placeholder="e.g. Thangiere P. Burns")
        b_form = st.selectbox("Form version", ["Preschool (ages 2-5)","Child (ages 6-11)","Adolescent (ages 12-21)"])

        st.markdown("#### Composite Scores  *(T-score + 90% CI + Percentile)*")
        st.caption("Get these from Q-Global printout or score report.")

        def composite_inputs(label, key_prefix, default_t=55, default_pct=50):
            c1,c2,c3,c4 = st.columns([2,1,1,1])
            c1.markdown(f"**{label}**")
            t    = c2.number_input(f"T", 20, 100, max(20, default_t),   key=f"{key_prefix}_t")
            ci_lo= c3.number_input(f"CI lo", 20, 100, max(20, default_t-4), key=f"{key_prefix}_ci_lo")
            ci_hi= c4.number_input(f"CI hi", 20, 100, max(20, min(100, default_t+4)), key=f"{key_prefix}_ci_hi")
            pct  = st.number_input(f"Percentile ({label})", 0, 99, default_pct, key=f"{key_prefix}_pct")
            return t, ci_lo, ci_hi, pct

        def subscale_inputs(scales, adaptive=False):
            data = {}
            for s in scales:
                key = s.lower().replace(" ", "_")
                default_t = 45 if adaptive else 55
                c1,c2,c3 = st.columns([3,1,1])
                c1.write(s)
                t   = c2.number_input("T", 20, 100, int(_basc.get(f"{s} T", default_t)), key=f"{key}_t")
                pct = c3.number_input("%ile", 0, 99, int(_basc.get(f"{s} pct", 50)),      key=f"{key}_pct")
                data[f"{key}_t"]   = t
                data[f"{key}_pct"] = pct
                color = t_flag(t, adaptive)
                cls   = t_classification(t, adaptive)
                if color == "red":    st.error(f"  → {cls}", icon=None)
                elif color == "orange": st.warning(f"  → {cls}", icon=None)
            return data

        ext_t,ext_ci_lo,ext_ci_hi,ext_pct = composite_inputs("Externalizing Problems", "ext", 57, 80)
        basc_data.update({"ext_t":ext_t,"ext_ci_lo":ext_ci_lo,"ext_ci_hi":ext_ci_hi,"ext_pct":ext_pct})
        st.markdown("*Externalizing subscales:*")
        basc_data.update(subscale_inputs(["Hyperactivity","Aggression","Conduct Problems"]))

        st.markdown("---")
        int_t,int_ci_lo,int_ci_hi,int_pct = composite_inputs("Internalizing Problems", "int", 47, 47)
        basc_data.update({"int_t":int_t,"int_ci_lo":int_ci_lo,"int_ci_hi":int_ci_hi,"int_pct":int_pct})
        st.markdown("*Internalizing subscales:*")
        basc_data.update(subscale_inputs(["Anxiety","Depression","Somatization"]))

        st.markdown("---")
        bsi_t,bsi_ci_lo,bsi_ci_hi,bsi_pct = composite_inputs("Behavioral Symptoms Index (BSI)", "bsi", 76, 98)
        basc_data.update({"bsi_t":bsi_t,"bsi_ci_lo":bsi_ci_lo,"bsi_ci_hi":bsi_ci_hi,"bsi_pct":bsi_pct})
        st.markdown("*Additional BSI subscales (Atypicality, Withdrawal, Attention Problems):*")
        basc_data.update(subscale_inputs(["Atypicality","Withdrawal","Attention Problems"]))

        st.markdown("---")
        adp_t,adp_ci_lo,adp_ci_hi,adp_pct = composite_inputs("Adaptive Skills", "adp", 23, 1)
        basc_data.update({"adp_t":adp_t,"adp_ci_lo":adp_ci_lo,"adp_ci_hi":adp_ci_hi,"adp_pct":adp_pct})
        st.markdown("*Adaptive subscales (low T = problem):*")
        basc_data.update(subscale_inputs(["Adaptability","Social Skills","Leadership",
                                          "Activities of Daily Living","Functional Communication"], adaptive=True))
        basc_data["respondent"] = b_respondent
        basc_data["form"]       = b_form

    st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 6c - VINELAND-3
# ══════════════════════════════════════════════════════════════════════
vineland_data = {}
if use_vineland:
    st.subheader("Vineland-3 - Adaptive Behavior Scales")

    vin_photo, vin_manual = st.tabs(["📷  Upload Score Sheet Photo", "✏️  Manual Entry"])
    _vin = st.session_state.get("extracted_vin_scores", {})

    with vin_photo:
        vin_img = st.file_uploader("Vineland-3 score sheet", type=["jpg","jpeg","png","webp"], key="vin_upload")
        if vin_img:
            st.image(vin_img, use_container_width=True)
            if st.button("Extract Vineland-3 Scores", type="primary"):
                with st.spinner("Reading Vineland scores..."):
                    try:
                        r = extract_scores_from_image(vin_img.read(), "Vineland-3 Adaptive Behavior Scales")
                        st.session_state["extracted_vin_scores"] = r.get("scores",{})
                        _vin = st.session_state["extracted_vin_scores"]
                        st.success("Extracted. Review in Manual Entry.")
                    except Exception as e:
                        st.error(f"Failed: {e}")

    with vin_manual:
        col1, col2 = st.columns(2)
        with col1:
            vin_respondent    = st.text_input("Respondent name + relationship", placeholder="e.g. Thangiere P. Burns, mother")
            vin_date_completed= st.text_input("Date form completed", value=eval_date.strftime("%m/%d/%Y"))
        with col2:
            st.caption("Standard Scores (mean=100, SD=15)")

        c1,c2,c3,c4 = st.columns(4)
        vin_abc   = c1.number_input("ABC",          20,160, int(_vin.get("ABC",44)))
        vin_comm  = c2.number_input("Communication",20,160, int(_vin.get("Communication",30)))
        vin_daily = c3.number_input("Daily Living",  20,160, int(_vin.get("Daily Living Skills",45)))
        vin_social= c4.number_input("Socialization", 20,160, int(_vin.get("Socialization",42)))

        # Auto percentiles
        ap = ss_to_pct(vin_abc); cp = ss_to_pct(vin_comm); dp = ss_to_pct(vin_daily); sp = ss_to_pct(vin_social)
        c1.caption(f"{'<1' if ap==0 else ap}th %ile · {ss_adaptive_level(vin_abc)}")
        c2.caption(f"{'<1' if cp==0 else cp}th %ile · {ss_adaptive_level(vin_comm)}")
        c3.caption(f"{'<1' if dp==0 else dp}th %ile · {ss_adaptive_level(vin_daily)}")
        c4.caption(f"{'<1' if sp==0 else sp}th %ile · {ss_adaptive_level(vin_social)}")

        vineland_data = {
            "abc": vin_abc, "comm": vin_comm, "daily": vin_daily, "social": vin_social,
            "comm_pct": cp, "daily_pct": dp, "social_pct": sp,
            "respondent": vin_respondent, "date_completed": vin_date_completed
        }

        # OPWDD eligibility check
        if use_ados or True:
            if vin_abc <= 70:
                st.error(f"ABC = {vin_abc} ≤ 70 - Meets adaptive deficit criterion for OPWDD / IDD eligibility")
            elif vin_abc <= 85:
                st.warning(f"ABC = {vin_abc} - Moderately Low range (below average but above IDD threshold)")
            else:
                st.success(f"ABC = {vin_abc} - Within or above average adaptive range")

    st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 6d - ADOS-2
# ══════════════════════════════════════════════════════════════════════
ados_data = {}
if use_ados:
    st.subheader("ADOS-2 - Autism Diagnostic Observation Schedule")

    ados_photo, ados_manual = st.tabs(["📷  Upload Score Sheet Photo", "✏️  Manual Entry"])
    _ados = st.session_state.get("extracted_ados_scores", {})

    with ados_photo:
        ados_img = st.file_uploader("ADOS-2 score sheet", type=["jpg","jpeg","png","webp"], key="ados_upload")
        if ados_img:
            st.image(ados_img, use_container_width=True)
            if st.button("Extract ADOS-2 Scores", type="primary"):
                with st.spinner("Reading ADOS-2 scores..."):
                    try:
                        r = extract_scores_from_image(ados_img.read(), "ADOS-2 Autism Diagnostic Observation Schedule")
                        st.session_state["extracted_ados_scores"] = r.get("scores",{})
                        _ados = st.session_state["extracted_ados_scores"]
                        st.success("Extracted. Review in Manual Entry.")
                    except Exception as e:
                        st.error(f"Failed: {e}")

    with ados_manual:
        col1, col2 = st.columns(2)
        with col1:
            ados_module = st.selectbox("Module", ["Toddler (T)","Module 1","Module 2","Module 3","Module 4"])
            ados_module_reason = st.text_input("Why this module was selected",
                placeholder="e.g. Selected because patient demonstrates low verbal ability, stating only single words")
            ados_sa  = st.number_input("Social Affect (SA) raw score", 0, 28, int(_ados.get("SA",20)))
            ados_rrb = st.number_input("Restricted & Repetitive Behavior (RRB) raw score", 0, 10, int(_ados.get("RRB",8)))
            ados_combined = ados_sa + ados_rrb
            st.metric("Algorithm Combined Total (SA + RRB)", ados_combined)
            ados_comparison = st.number_input("Comparison Score (1-10)", 1, 10, int(_ados.get("Comparison Score",10)))

        with col2:
            # Auto-classify
            auto_sa_label   = ados_sa_label(ados_module, ados_sa)
            auto_rrb_label  = ados_rrb_label(ados_rrb)
            auto_class      = ados_classify(ados_module, ados_sa, ados_rrb)
            auto_cs_label   = comparison_score_label(ados_comparison)

            st.markdown("**Auto-Classification**")
            st.write(f"SA ({ados_sa}): **{auto_sa_label}**")
            st.write(f"RRB ({ados_rrb}): **{auto_rrb_label}**")
            st.write(f"Combined ({ados_combined}): **{auto_cs_label}**")
            if auto_class == "Autism":
                st.error(f"ADOS-2 Classification: **{auto_class}**")
            elif auto_class == "Autism Spectrum":
                st.warning(f"ADOS-2 Classification: **{auto_class}**")
            else:
                st.success(f"ADOS-2 Classification: **{auto_class}**")

        st.markdown("---")
        st.markdown("#### DSM-5 Criteria Evaluation")

        col1, col2 = st.columns(2)
        with col1:
            dsm5_a_met = st.checkbox("Criterion A: Symptoms present in early developmental period", value=True)
            if dsm5_a_met:
                dsm5_a_age = st.number_input("Age symptoms first noted (years)", 0, 18, 2)
            else:
                dsm5_a_age = None
            dsm5_b_desc = st.text_area("Criterion B: Clinically significant impairment description",
                value="Global impact across all areas.",
                height=70)
        with col2:
            dsm5_c_desc = st.text_area("Criterion C: Not better explained by ID alone",
                value="There is evidence of an intellectual disability and global developmental delay. Cognitive ability is impacted by receptive and expressive speech processing.",
                height=70)

        st.markdown("#### Specifiers")
        col1, col2 = st.columns(2)
        with col1:
            intellectual_impairment = st.checkbox("Intellectual Impairment present", value=True)
            language_impairment     = st.text_input("Language Impairment type",
                placeholder="e.g. Area of Receptive-Expressive Language Processing")
        with col2:
            criteria_a_level = st.selectbox("Criteria A Severity (Social Communication)", [1,2,3], index=2,
                format_func=lambda x: f"Level {x} - {'Requiring Support' if x==1 else 'Requiring Substantial Support' if x==2 else 'Requiring Very Substantial Support'}")
            criteria_b_level = st.selectbox("Criteria B Severity (RRB)", [1,2,3], index=0,
                format_func=lambda x: f"Level {x} - {'Requiring Support' if x==1 else 'Requiring Substantial Support' if x==2 else 'Requiring Very Substantial Support'}")

        ados_data = {
            "module": ados_module, "module_reason": ados_module_reason,
            "sa": ados_sa, "rrb": ados_rrb, "combined": ados_combined,
            "comparison": ados_comparison, "classification": auto_class,
            "dsm5_a_met": dsm5_a_met, "dsm5_a_age": dsm5_a_age,
            "dsm5_b_desc": dsm5_b_desc, "dsm5_c_desc": dsm5_c_desc,
            "intellectual_impairment": intellectual_impairment,
            "language_impairment": language_impairment,
            "criteria_a_level": criteria_a_level, "criteria_b_level": criteria_b_level,
        }

    st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 7 - DIAGNOSIS + RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
st.subheader("Conclusion & Recommendations")

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Primary Diagnosis (DSM-5)**")
    primary_dx = st.selectbox("Primary diagnosis", [
        "F84.0 - Autism Spectrum Disorder",
        "F84.0 - ASD (Level 1: Requiring Support)",
        "F84.0 - ASD (Level 2: Requiring Substantial Support)",
        "F84.0 - ASD (Level 3: Requiring Very Substantial Support)",
        "F70 - Mild Intellectual Disability",
        "F71 - Moderate Intellectual Disability",
        "F72 - Severe Intellectual Disability",
        "F73 - Profound Intellectual Disability",
        "F88 - Other Disorders of Psychological Development",
        "Rule Out ASD - Inconclusive",
    ])
    additional_dx = st.text_input("Additional diagnoses", placeholder="e.g. F70 Mild ID, F80.9 Language Disorder")

with col2:
    st.markdown("**Recommendations** *(check all that apply)*")
    rec_parent_meeting  = st.checkbox("Meet with parents to review results", value=True)
    rec_medical_followup= st.checkbox("Medical follow-up - Psychiatrist or Neurologist", value=True)
    rec_cpse            = st.checkbox("Contact NYC DOE / CPSE - submit report, request meeting", value=True)
    rec_classification  = st.text_input("Proposed CSE classification change", placeholder="e.g. Multiple Disabilities (MD)")
    rec_parent_training = st.checkbox("Parent training via school district IEP", value=True)
    rec_aba             = st.checkbox("ABA therapy", value=True)
    rec_aba_detail      = st.text_input("ABA focus areas", placeholder="e.g. Toilet training, aggression, SIB") if rec_aba else ""
    rec_feeding_therapy = st.checkbox("Add feeding therapy to IEP")
    rec_opwdd           = st.checkbox("Refer to OPWDD NYC", value=True)
    rec_other           = st.text_area("Other recommendations", height=60, placeholder="e.g. Sensory integration therapy, AAC evaluation")

st.divider()

# ══════════════════════════════════════════════════════════════════════
# BLOCK 8 - CHECKLIST
# ══════════════════════════════════════════════════════════════════════
st.subheader("Pre-Generation Checklist")
col1, col2 = st.columns(2)
with col1:
    chk1 = st.checkbox("Reason for referral documented")
    chk2 = st.checkbox("Background information collected (parent interview done)")
    chk3 = st.checkbox("Behavioral observations written")
    chk4 = st.checkbox("All administered test scores entered")
with col2:
    chk5 = st.checkbox("ADOS-2 DSM-5 criteria evaluated")
    chk6 = st.checkbox("Diagnosis confirmed")
    chk7 = st.checkbox("PHI de-identified (first name / initials only, no last name or full DOB)")
    chk8 = st.checkbox("Clinician will review and sign draft before use or distribution")

all_checked = all([chk1,chk2,chk3,chk4,chk5,chk6,chk7,chk8])
st.divider()


# ══════════════════════════════════════════════════════════════════════
# GENERATION
# ══════════════════════════════════════════════════════════════════════

def build_eval_materials_list():
    tests = []
    if use_wppsi:   tests.append("Cognitive Scores from WPPSI-IV (Composite Subtests)")
    if use_wisc:    tests.append("Cognitive Scores from WISC-V (Composite Subtests)")
    if use_ados:    tests.append(f"Autism Diagnostic Observation Schedule, Second Edition - {ados_data.get('module','Module 1')} (ADOS-2)")
    if use_basc:    tests.append(f"Behavior Assessment System for Children - 3rd Edition ({basc_data.get('form','PRS-P')})")
    if use_vineland:tests.append("Vineland Adaptive Behavior Scales - Parent")
    if use_case:    tests.append("Case Materials Reviewed")
    if extra_tests: tests += [t.strip() for t in extra_tests.split(",") if t.strip()]
    return tests

def build_background_text():
    lines = []
    if prior_evals_text:
        lines.append(f"Prior evaluations: {prior_evals_text}")
    if prior_diagnoses:
        lines.append(f"Prior diagnoses: {prior_diagnoses}")
    if services:
        lines.append(f"Prior/current services: {', '.join(services)}")
    lines.append(f"\nParent interview conducted on {interview_date}.")
    lines.append(f"Family: {family_composition or 'Not specified'}")
    if family_hx: lines.append(f"Family history: {family_hx}")
    lines.append(f"Communication: {comm_level}" + (f" - {echolalia_detail}" if echolalia and echolalia_detail else ""))
    if not toilet_trained: lines.append("Not toilet trained - wears diapers.")
    if feeding_difficulties: lines.append(f"Feeding: {feeding_difficulties}")
    if behavioral_concerns: lines.append(f"Behavioral concerns: {behavioral_concerns}")
    lines.append(f"School: {school_name or 'Not specified'} - Placement: {school_placement_type or grade_placement or 'Not specified'}")
    lines.append(f"IEP: {iep_status}" + (f" - Classification: {cse_classification}" if cse_classification else ""))
    if birth_complications: lines.append(f"Birth/prenatal: {birth_complications}")
    if medical_diagnoses:   lines.append(f"Medical: {medical_diagnoses}")
    if current_meds:        lines.append(f"Medications: {current_meds}")
    lines.append(f"Milestones - Walking: {milestone_walking}; Words: {milestone_words}; Phrases: {milestone_phrases}")
    if regression:          lines.append(f"Regression: {regression_detail or 'reported'}")
    return "\n".join(lines)

def build_recs_list():
    recs = []
    if rec_parent_meeting:  recs.append("Meet with parents/guardians to review the results and findings of the Psychological-ASD evaluation.")
    recs.append("It is common for those diagnosed with ASD to have comorbidities such as anxiety, depression, and ADHD. Follow-up with a medical provider (Psychiatrist or Neurologist) is recommended in addition to continuing with outside therapy.")
    if rec_cpse:
        r = f"Contact the office of Special Education in the NYC DOE, submit the report and request a CPSE meeting"
        if rec_classification: r += f" to change classification to {rec_classification}"
        recs.append(r + ".")
    if rec_parent_training: recs.append("Parent Training to be recommended by the school district on the IEP.")
    if rec_aba:
        r = "ABA therapy"
        if rec_aba_detail: r += f" to help with {rec_aba_detail}"
        recs.append(r + ".")
    if rec_feeding_therapy: recs.append("Add feeding therapy to IEP.")
    if rec_opwdd:           recs.append("Refer to the OPWDD NYC.")
    if rec_other:
        for line in rec_other.strip().split("\n"):
            if line.strip(): recs.append(line.strip())
    return recs

def build_prompt():
    bg   = build_background_text()
    obs  = st.session_state.get("observations","")
    recs = build_recs_list()
    tests_list = "\n".join(f"• {t}" for t in build_eval_materials_list())

    cog_section = ""
    if use_wppsi or use_wisc:
        cog_label = "WPPSI-IV" if use_wppsi else "WISC-V"
        if not cog_scores.get("obtained", True):
            cog_section = f"""
{cog_label}:
No composite score was obtained on the {cog_label}. {cog_scores.get('reason','')}"""
        else:
            score_lines = "\n".join(f"  {k}: SS={v} ({ss_to_pct(v)}th %ile - {('Low' if v<70 else 'Borderline' if v<80 else 'Low Average' if v<90 else 'Average' if v<110 else 'High Average')})"
                                     for k,v in cog_scores.items() if k not in ("obtained","reason"))
            cog_section = f"\n{cog_label} scores:\n{score_lines}"

    ados_section = ""
    if use_ados and ados_data:
        ados_section = generate_ados_narrative(
            ados_data["module"], ados_data["sa"], ados_data["rrb"],
            ados_data["comparison"], ados_data["classification"],
            ados_data["dsm5_a_met"], ados_data["dsm5_a_age"],
            ados_data["dsm5_b_desc"], ados_data["dsm5_c_desc"],
            ados_data["intellectual_impairment"], ados_data["language_impairment"],
            ados_data["criteria_a_level"], ados_data["criteria_b_level"],
            patient_name
        )

    vineland_section = ""
    if use_vineland and vineland_data:
        vineland_section = generate_vineland_narrative(
            vineland_data["abc"], vineland_data["comm"], vineland_data["daily"], vineland_data["social"],
            vineland_data["comm_pct"], vineland_data["daily_pct"], vineland_data["social_pct"],
            patient_name, vineland_data.get("respondent",""), vineland_data.get("date_completed","")
        )

    basc_section = ""
    if use_basc and basc_data:
        basc_section = generate_basc3_narrative(basc_data, patient_name, basc_data.get("respondent",""))

    recs_text = "\n".join(f"{i+1}. {r}" for i,r in enumerate(recs))

    return f"""You are a licensed clinical psychologist writing a formal Psychological Autism Spectrum Disorder Evaluation report.
Match this EXACT style: third person, formal clinical language, patient referred to by first name throughout.

Write ONLY the following three sections. The other sections have already been generated.
Use the name "{patient_name or '[PATIENT]'}" throughout.

---
SECTION TO WRITE 1 - BACKGROUND INFORMATION:
Write a flowing clinical narrative (2-3 paragraphs) from these structured notes.
First paragraph: prior evaluation history and results.
Second paragraph: parent interview findings - family, communication, behaviors, school, services, diagnoses.
Use formal language. Be specific. Include all details provided.

Background notes:
{bg}

---
SECTION TO WRITE 2 - BEHAVIORAL OBSERVATIONS:
Expand these clinician notes into a full clinical narrative (2-3 paragraphs).
Cover: arrival, appearance, motor, eye contact, communication, joint attention, play, attention, stereotypies, affect, effort.
End with validity statement.

Clinician notes:
{obs}

---
SECTION TO WRITE 3 - CONCLUSION AND STATEMENT OF DIAGNOSIS:
Based on all assessment data below, write the conclusion paragraph.
State that DSM-5 criteria are met (or not). Cite specific tests. Name the diagnosis: {primary_dx}.
{"Additional: " + additional_dx if additional_dx else ""}

Assessment summary:
{cog_section}

BASC-3 key findings:
  BSI composite T={basc_data.get('bsi_t','N/A')} (Clinically Significant if ≥70)
  Atypicality T={basc_data.get('atypicality_t','N/A')}
  Attention Problems T={basc_data.get('attention_problems_t','N/A')}
  Adaptive Skills composite T={basc_data.get('adp_t','N/A')}
  Functional Communication T={basc_data.get('functional_communication_t','N/A')}

Vineland-3: ABC={vineland_data.get('abc','N/A')}, Comm={vineland_data.get('comm','N/A')}, DLS={vineland_data.get('daily','N/A')}, Social={vineland_data.get('social','N/A')}

ADOS-2: SA={ados_data.get('sa','N/A')}, RRB={ados_data.get('rrb','N/A')}, Combined={ados_data.get('combined','N/A')}, Classification={ados_data.get('classification','N/A')}
Criteria A Level {ados_data.get('criteria_a_level','N/A')}, Criteria B Level {ados_data.get('criteria_b_level','N/A')}

---
Return ONLY the three sections with these exact headers:
BACKGROUND INFORMATION:
[text]

BEHAVIORAL OBSERVATIONS:
[text]

CONCLUSION AND STATEMENT OF DIAGNOSIS:
[text]
"""

def run_llm(prompt):
    resp = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role":"user","content":prompt}],
        temperature=0.25,
        max_tokens=4096,
    )
    return resp.choices[0].message.content

def assemble_full_report(llm_output):
    """Combine rule-based sections + LLM sections into final report."""

    tests_list = "\n".join(build_eval_materials_list())
    recs = build_recs_list()
    recs_text = "\n".join(f"     {i+1}. {r}" for i,r in enumerate(recs))

    cog_label = "WPPSI-IV" if use_wppsi else ("WISC-V" if use_wisc else "Cognitive Assessment")
    if use_wppsi or use_wisc:
        if not cog_scores.get("obtained", True):
            cog_text = f"""Wechsler Preschool Primary Scale of Intelligence - Fourth Edition ({cog_label})

No composite score was obtained on the {cog_label}. {cog_scores.get('reason','')}"""
        else:
            score_lines = "\n".join(f"  {k}: SS = {v}  ({ss_to_pct(v)}th percentile)"
                                     for k,v in cog_scores.items() if k not in ("obtained","reason"))
            cog_text = f"""{cog_label}\n\n{score_lines}"""
    else:
        cog_text = ""

    basc_text  = generate_basc3_narrative(basc_data, patient_name, basc_data.get("respondent","")) if use_basc and basc_data else ""
    vin_text   = generate_vineland_narrative(
        vineland_data["abc"],vineland_data["comm"],vineland_data["daily"],vineland_data["social"],
        vineland_data["comm_pct"],vineland_data["daily_pct"],vineland_data["social_pct"],
        patient_name,vineland_data.get("respondent",""),vineland_data.get("date_completed","")
    ) if use_vineland and vineland_data else ""
    ados_text  = generate_ados_narrative(
        ados_data["module"],ados_data["sa"],ados_data["rrb"],ados_data["comparison"],ados_data["classification"],
        ados_data["dsm5_a_met"],ados_data["dsm5_a_age"],ados_data["dsm5_b_desc"],ados_data["dsm5_c_desc"],
        ados_data["intellectual_impairment"],ados_data["language_impairment"],
        ados_data["criteria_a_level"],ados_data["criteria_b_level"],patient_name
    ) if use_ados and ados_data else ""

    report = f"""*** DRAFT - FOR CLINICIAN REVIEW AND SIGNATURE ONLY. NOT FOR DISTRIBUTION. ***

Psychological Autism Spectrum Disorder Evaluation
Privileged and Confidential Information

NAME: {patient_name or '[NAME]'}                          DATE OF EVALUATION: {eval_date.strftime('%m/%d/%Y')}
D.O.B: {dob_text or '[DOB]'}                              CHRONOLOGICAL AGE: {chron_age or '[AGE]'}
EXAMINER: {examiner_name or '[EXAMINER]'}                 LANGUAGE OF TESTING: {language}
SCHOOL: {school_name or '[SCHOOL]'}
GRADE: {grade_placement or '[GRADE]'}

─────────────────────────────────────────────────────────────

Reason for Referral:

{patient_name or 'The patient'} was referred to the examiner for a psychological evaluation by {referral_by.lower()} {referral_concern}

─────────────────────────────────────────────────────────────

Evaluation Materials:

{tests_list}

─────────────────────────────────────────────────────────────

{llm_output}

─────────────────────────────────────────────────────────────

Assessments:

{cog_text}

{"─" * 40 if cog_text else ""}

{"Behavior Assessment System for Children, (BASC-3)" if use_basc else ""}
{"Parent Rating Scales" if use_basc else ""}

{basc_text}

{"─" * 40 if vin_text else ""}

{"Vineland Adaptive Behavior Scales, Third Edition (Vineland™-3)" if use_vineland else ""}
{"Domain-Level Parent/Caregiver Form Report" if use_vineland else ""}

{vin_text}

{"─" * 40 if ados_text else ""}

{ados_text}

─────────────────────────────────────────────────────────────

Recommendations:

{recs_text}

─────────────────────────────────────────────────────────────

____________________________          ________________________
{examiner_name or '[Examiner Name]'}
Psychology Intern                     {supervisor_name or '[Supervisor Name]'}
                                      Supervising Psychologist
                                      NPI {supervisor_npi or '[NPI]'}
                                      License # {supervisor_lic or '[License]'}
                                      contact@rimonhealth.com
                                      347-746-6613
"""
    return report


# ── GENERATE BUTTON ────────────────────────────────────────────────────────────

if st.button("Generate Report", disabled=not all_checked, type="primary", use_container_width=True):
    missing = []
    if not patient_name.strip():      missing.append("Patient name")
    if not referral_concern.strip():  missing.append("Reason for referral")
    if not st.session_state.get("observations","").strip(): missing.append("Behavioral observations")
    if missing:
        st.error(f"Missing: {', '.join(missing)}")
    else:
        with st.spinner("Generating report... (~20 seconds)"):
            try:
                prompt     = build_prompt()
                llm_out    = run_llm(prompt)
                full_report = assemble_full_report(llm_out)
                st.session_state["report"]      = full_report
                st.session_state["report_ts"]   = datetime.now().strftime("%Y%m%d_%H%M")
                st.success("Draft generated. Review carefully before use.")
            except Exception as e:
                st.error(f"Generation failed: {e}")
elif not all_checked:
    st.warning("Complete all checklist items above to enable generation.")

# ══════════════════════════════════════════════════════════════════════
# OUTPUT
# ══════════════════════════════════════════════════════════════════════

def sanitize(text):
    for k,v in [("-","-"),("-","-"),("'","'"),("'","'"),(""",'"'),(""",'"'),("…","..."),("•","-"),("™","")]:
        text = text.replace(k,v)
    return text.encode("latin-1", errors="replace").decode("latin-1")

def make_pdf(report_text):
    txt = sanitize(report_text)
    pdf = FPDF(); pdf.set_auto_page_break(True, 20); pdf.set_margins(25,20,25); pdf.add_page()
    pdf.set_font("Helvetica","",10)
    for line in txt.split("\n"):
        line = sanitize(line)
        if not line.strip(): pdf.ln(3); continue
        clean = line.replace("**","").replace("*","").replace("#","").strip()
        if not clean: continue
        is_title = clean.startswith("Psychological Autism") or clean.startswith("Privileged")
        is_hdr   = (clean.isupper() and 6 < len(clean) < 100) or clean.startswith("─")
        is_draft = clean.startswith("***")
        if is_draft:
            pdf.set_font("Helvetica","BI",9); pdf.set_text_color(180,0,0)
            pdf.write(6, clean+"\n")
            pdf.set_font("Helvetica","",10); pdf.set_text_color(0,0,0)
        elif is_title:
            pdf.set_font("Helvetica","B",12); pdf.write(8, clean+"\n"); pdf.set_font("Helvetica","",10)
        elif is_hdr:
            pdf.set_font("Helvetica","B",10); pdf.write(7, clean+"\n"); pdf.set_font("Helvetica","",10)
        else:
            pdf.write(5.5, clean+"\n")
    pdf.ln(8); pdf.set_font("Helvetica","I",7); pdf.set_text_color(130,130,130)
    pdf.write(5,"DRAFT - FOR CLINICIAN REVIEW AND SIGNATURE ONLY. Not for distribution without authorized sign-off.")
    return pdf.output()

def make_docx(report_text):
    doc = Document()
    # Page margins
    for sec in doc.sections:
        sec.top_margin    = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin   = Inches(1.25)
        sec.right_margin  = Inches(1.25)

    for line in report_text.split("\n"):
        line = line.strip()
        if not line:
            doc.add_paragraph()
            continue
        clean = line.replace("**","").replace("*","").replace("#","").strip()
        if not clean: continue

        is_draft = clean.startswith("***")
        is_title = clean.startswith("Psychological Autism") or clean.startswith("Privileged")
        is_hdr   = (clean.isupper() and 6 < len(clean) < 100)

        if is_draft:
            p = doc.add_paragraph()
            r = p.add_run(clean)
            r.bold = True; r.italic = True
            r.font.color.rgb = RGBColor(0xCC,0x00,0x00)
            r.font.size = Pt(9)
        elif is_title:
            p = doc.add_heading(clean, level=1)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif is_hdr:
            p = doc.add_heading(clean, level=2)
        else:
            p = doc.add_paragraph(clean)
            if p.runs: p.runs[0].font.size = Pt(11)

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()

if "report" in st.session_state:
    st.divider()
    st.subheader("Draft Report")
    st.warning("AI-generated DRAFT. Clinician must review, edit, and sign before any use or distribution.")

    edited = st.text_area("Report (editable - changes saved on re-generate)", st.session_state["report"], height=700)

    st.divider()
    st.subheader("Export")
    fname = f"DRAFT_{(patient_name or 'patient').replace(' ','_')}_{st.session_state['report_ts']}"

    col1,col2,col3,col4 = st.columns(4)
    with col1:
        pdf_bytes = bytes(make_pdf(edited))
        st.download_button("Download PDF", pdf_bytes, fname+".pdf", "application/pdf", use_container_width=True)
    with col2:
        docx_bytes = bytes(make_docx(edited))
        st.download_button("Download Word", docx_bytes, fname+".docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with col3:
        with st.expander("Email"):
            to_email = st.text_input("To")
            to_name  = st.text_input("Recipient name")
            note     = st.text_area("Note", height=60)
            if st.button("Send"):
                sender = os.getenv("SENDER_EMAIL"); pw = os.getenv("SENDER_APP_PASSWORD")
                if not sender: st.error("Set SENDER_EMAIL in .env")
                elif not to_email: st.error("Enter recipient email")
                else:
                    try:
                        msg = MIMEMultipart()
                        msg["From"]=sender; msg["To"]=to_email
                        msg["Subject"]=f"[DRAFT] ASD Evaluation - {patient_name or 'patient'}"
                        body = f"Dear {to_name or 'Colleague'},\n\nPlease find attached the draft ASD evaluation report.\n\nThis is an AI-generated DRAFT requiring clinician review and sign-off before use.\n\n{note}\n\nWarm regards,\n{examiner_name or 'Rimon Health'}\nRimon Health\n\n---\nDRAFT only."
                        msg.attach(MIMEText(body,"plain"))
                        part = MIMEBase("application","octet-stream"); part.set_payload(pdf_bytes)
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition",f'attachment; filename="{fname}.pdf"')
                        msg.attach(part)
                        with smtplib.SMTP_SSL("smtp.gmail.com",465) as s:
                            s.login(sender,pw); s.sendmail(sender,to_email,msg.as_string())
                        st.success(f"Sent to {to_name or to_email}.")
                    except Exception as e: st.error(f"Failed: {e}")
    with col4:
        if st.button("Print", use_container_width=True):
            try:
                with tempfile.NamedTemporaryFile(suffix=".pdf",delete=False) as tmp:
                    tmp.write(pdf_bytes); path=tmp.name
                subprocess.Popen(["open","-a","Preview",path])
                st.success("Sent to Preview.")
            except Exception as e: st.error(f"Failed: {e}")
