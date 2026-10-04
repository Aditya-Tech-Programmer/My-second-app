
# ============================================================
# 🌍 LIFE SIMULATOR PRO
# Advanced Streamlit + OpenPyXL
# ============================================================

import os
import uuid
import random
import time
from datetime import datetime, date

import streamlit as st
from openpyxl import Workbook, load_workbook

try:
    from google import genai
except ImportError:
    genai = None

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Aditya^s Life Simulator",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded",
)

FILE_NAME = "Life_Simulator.xlsx"

# ============================================================
# GLOBAL STYLES
# ============================================================

st.markdown("""
<style>
/* Main */
.main {
    background: #f7f9fc;
}

.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 1400px;
}

/* Hide default menu/footer */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* Hero */
.hero {
    padding: 28px 22px;
    border-radius: 24px;
    background: linear-gradient(135deg, #111827 0%, #1d4ed8 52%, #7c3aed 100%);
    color: white;
    text-align: center;
    box-shadow: 0 12px 35px rgba(30, 64, 175, .20);
    margin-bottom: 18px;
}

.hero h1 {
    font-size: clamp(38px, 7vw, 72px);
    font-weight: 900;
    letter-spacing: 2px;
    margin: 0;
}

.hero p {
    margin: 8px 0 0 0;
    font-size: 18px;
    opacity: .92;
}

/* Feature strip */
.feature-strip {
    padding: 14px 18px;
    border-radius: 16px;
    background: white;
    border: 1px solid #e5e7eb;
    box-shadow: 0 5px 18px rgba(15,23,42,.05);
    margin-bottom: 18px;
    text-align: center;
    line-height: 2;
}

/* Cards */
.card {
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    padding: 18px;
    margin-bottom: 14px;
    box-shadow: 0 6px 20px rgba(15,23,42,.05);
}

.card-title {
    font-size: 20px;
    font-weight: 800;
    margin-bottom: 5px;
}

.card-subtitle {
    color: #64748b;
    font-size: 14px;
}

/* Stat cards */
.stat-card {
    background: white;
    border-radius: 18px;
    padding: 18px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 6px 20px rgba(15,23,42,.05);
    min-height: 110px;
}

.stat-label {
    color: #64748b;
    font-size: 14px;
    font-weight: 700;
}

.stat-value {
    font-size: 28px;
    font-weight: 900;
    margin-top: 6px;
}

/* Section */
.section-title {
    font-size: 27px;
    font-weight: 900;
    margin: 18px 0 12px;
}

/* Badges */
.badge {
    display: inline-block;
    padding: 6px 11px;
    border-radius: 999px;
    background: #eef2ff;
    color: #3730a3;
    font-weight: 700;
    font-size: 12px;
    margin: 2px;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #111827 100%);
}

[data-testid="stSidebar"] * {
    color: white !important;
}

[data-testid="stSidebar"] .stRadio label {
    border-radius: 10px;
}

/* Buttons */
.stButton > button {
    border-radius: 12px;
    min-height: 44px;
    font-weight: 700;
}

/* Progress */
.progress-wrap {
    height: 10px;
    background: #e5e7eb;
    border-radius: 999px;
    overflow: hidden;
    margin: 8px 0;
}

.progress-bar {
    height: 100%;
    background: linear-gradient(90deg, #2563eb, #7c3aed);
    border-radius: 999px;
}

/* Timeline */
.timeline-item {
    border-left: 4px solid #4f46e5;
    padding: 12px 16px;
    background: white;
    border-radius: 0 14px 14px 0;
    margin-bottom: 10px;
    box-shadow: 0 4px 14px rgba(15,23,42,.04);
}

/* Mobile */
@media (max-width: 700px) {
    .hero h1 {
        font-size: 42px;
    }

    .hero p {
        font-size: 15px;
    }
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# DATABASE SCHEMA
# ============================================================

HEADERS = {
    "Users": [
        "User_ID", "Name", "Age", "Education", "Career", "Hobby",
        "Health", "Happiness", "Energy", "Skills", "Discipline",
        "Social", "Money", "Level", "XP", "Streak", "Day",
        "Part_Time_Job", "Created_Date"
    ],
    "Goals": [
        "Goal_ID", "User_ID", "Goal", "Category", "Target",
        "Progress", "Deadline", "Status", "Created_Date"
    ],
    "Stats": [
        "Record_ID", "User_ID", "Day", "Health", "Happiness",
        "Energy", "Skills", "Discipline", "Social", "Money",
        "XP", "Level", "Date"
    ],
    "Timeline": [
        "Timeline_ID", "User_ID", "Day", "Time",
        "Activity", "Category", "Effect", "Date"
    ],
    "Achievements": [
        "Achievement_ID", "User_ID", "Achievement",
        "Description", "Unlocked_Day", "Date"
    ],
    "Money_History": [
        "Transaction_ID", "User_ID", "Day", "Type",
        "Category", "Description", "Amount", "Balance", "Date"
    ],
    "Timetable": [
        "Timetable_ID", "User_ID", "Day", "Time",
        "Activity", "Category", "Reason", "Priority", "Date"
    ],
}

# ============================================================
# DATABASE
# ============================================================

def ensure_database():
    if not os.path.exists(FILE_NAME):
        wb = Workbook()
        wb.active.title = "Users"

        for sheet_name, headers in HEADERS.items():
            if sheet_name == "Users":
                ws = wb["Users"]
            else:
                ws = wb.create_sheet(sheet_name)
            ws.append(headers)

        wb.save(FILE_NAME)
        return

    wb = load_workbook(FILE_NAME)

    for sheet_name, headers in HEADERS.items():
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            ws.append(headers)
            continue

        ws = wb[sheet_name]
        existing = [c.value for c in ws[1]]

        if not existing or all(v is None for v in existing):
            for col, header in enumerate(headers, 1):
                ws.cell(1, col, header)
        else:
            for header in headers:
                if header not in existing:
                    ws.cell(1, ws.max_column + 1, header)

    wb.save(FILE_NAME)


ensure_database()


def load_db():
    return load_workbook(FILE_NAME)


def save_db(wb):
    wb.save(FILE_NAME)


def append_row(sheet_name, data):
    wb = load_db()
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    ws.append([data.get(h, "") for h in headers])
    save_db(wb)


def get_rows(sheet_name):
    wb = load_db()
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    return [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)]


# ============================================================
# SAFE CONVERSION
# ============================================================

def safe_int(value, default=0):
    if value is None or value == "":
        return default
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default


def safe_float(value, default=0.0):
    if value is None or value == "":
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def normalize_user(user):
    defaults = {
        "Age": 0,
        "Health": 0,
        "Happiness": 0,
        "Energy": 0,
        "Skills": 0,
        "Discipline": 0,
        "Social": 0,
        "Money": 0.0,
        "Level": 0,
        "XP": 0,
        "Streak": 0,
        "Day": 0,
        "Name": "Player",
        "Education": "",
        "Career": "Write Career",
        "Hobby": "Write Hobby",
        "Part_Time_Job": "No",
    }

    for key, default in defaults.items():
        if user.get(key) is None or user.get(key) == "":
            user[key] = default

    for key in [
        "Age", "Health", "Happiness", "Energy", "Skills",
        "Discipline", "Social", "Level", "XP", "Streak", "Day"
    ]:
        user[key] = safe_int(user.get(key), defaults[key])

    user["Money"] = safe_float(user.get("Money"), 0.0)

    return user


def find_user(user_id):
    for user in get_rows("Users"):
        if str(user.get("User_ID")) == str(user_id):
            return normalize_user(user)
    return None


def all_users():
    return [
        normalize_user(u)
        for u in get_rows("Users")
        if u.get("User_ID")
    ]


def update_user(user):
    user = normalize_user(user)

    wb = load_db()
    ws = wb["Users"]
    headers = [c.value for c in ws[1]]

    for row in ws.iter_rows(min_row=2):
        if str(row[headers.index("User_ID")].value) == str(user["User_ID"]):
            row_number = row[0].row

            for key, value in user.items():
                if key in headers:
                    ws.cell(
                        row=row_number,
                        column=headers.index(key) + 1,
                        value=value
                    )
            break

    save_db(wb)


# ============================================================
# USER CREATION
# ============================================================

def create_user(name, age, education, career, hobby, part_time):
    user_id = "USR-" + str(random.randint(10000, 99999))

    while find_user(user_id):
        user_id = "USR-" + str(random.randint(10000, 99999))

    user = {
        "User_ID": user_id,
        "Name": name.strip(),
        "Age": int(age),
        "Education": education,
        "Career": career.strip() or "Write Career",
        "Hobby": hobby.strip() or "Write Hobby",
        "Health": 0,
        "Happiness": 0,
        "Energy": 0,
        "Skills": 0,
        "Discipline": 0,
        "Social": 0,
        "Money": 0.0,
        "Level": 0,
        "XP": 0,
        "Streak": 0,
        "Day": 0,
        "Part_Time_Job": "Yes" if part_time else "No",
        "Created_Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    append_row("Users", user)

    append_row("Money_History", {
        "Transaction_ID": "TXN-" + str(uuid.uuid4())[:8],
        "User_ID": user_id,
        "Day": 1,
        "Type": "Initial Balance",
        "Category": "Starting Money",
        "Description": "Initial Life Simulator balance",
        "Amount": 0.0,
        "Balance": 0.0,
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })

    return user


# ============================================================
# GAME LOGIC
# ============================================================

def clamp(value, low=0, high=100):
    return max(low, min(high, safe_int(value)))


def change_stat(user, stat, amount):
    if stat in user:
        user[stat] = clamp(safe_int(user.get(stat)) + amount)


def add_xp(user, amount):
    user = normalize_user(user)
    amount = safe_int(amount)

    user["XP"] += amount

    while user["XP"] >= user["Level"] * 100:
        required = user["Level"] * 100
        user["XP"] -= required
        user["Level"] += 1
        st.toast(f"🎉 LEVEL UP! Level {user['Level']}")

    update_user(user)


def add_timeline(user, activity, category, effect):
    append_row("Timeline", {
        "Timeline_ID": "TL-" + str(uuid.uuid4())[:8],
        "User_ID": user["User_ID"],
        "Day": user["Day"],
        "Time": datetime.now().strftime("%H:%M"),
        "Activity": activity,
        "Category": category,
        "Effect": effect,
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


def save_stats(user):
    user = normalize_user(user)

    append_row("Stats", {
        "Record_ID": "STAT-" + str(uuid.uuid4())[:8],
        "User_ID": user["User_ID"],
        "Day": user["Day"],
        "Health": user["Health"],
        "Happiness": user["Happiness"],
        "Energy": user["Energy"],
        "Skills": user["Skills"],
        "Discipline": user["Discipline"],
        "Social": user["Social"],
        "Money": user["Money"],
        "XP": user["XP"],
        "Level": user["Level"],
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


# ============================================================
# MONEY
# ============================================================

def money_transaction(user, transaction_type, category, description, amount):
    user = normalize_user(user)
    amount = safe_float(amount)

    if amount <= 0:
        return False, "Amount must be greater than 0."

    if transaction_type == "Expense":
        if amount > user["Money"]:
            return False, "Insufficient balance!"
        user["Money"] -= amount

    elif transaction_type == "Income":
        user["Money"] += amount

    append_row("Money_History", {
        "Transaction_ID": "TXN-" + str(uuid.uuid4())[:8],
        "User_ID": user["User_ID"],
        "Day": user["Day"],
        "Type": transaction_type,
        "Category": category,
        "Description": description or "No description",
        "Amount": amount,
        "Balance": user["Money"],
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })

    update_user(user)
    return True, "Transaction saved successfully!"


def transactions_for(user_id):
    return [
        r for r in get_rows("Money_History")
        if str(r.get("User_ID")) == str(user_id)
    ]


# ============================================================
# GOALS
# ============================================================

def create_goal(user, goal, category, target, deadline):
    append_row("Goals", {
        "Goal_ID": "GOAL-" + str(uuid.uuid4())[:8],
        "User_ID": user["User_ID"],
        "Goal": goal,
        "Category": category,
        "Target": target,
        "Progress": 0,
        "Deadline": str(deadline),
        "Status": "Active",
        "Created_Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


def goals_for(user_id):
    return [
        r for r in get_rows("Goals")
        if str(r.get("User_ID")) == str(user_id)
    ]


def update_goal(goal_id, progress):
    wb = load_db()
    ws = wb["Goals"]
    headers = [c.value for c in ws[1]]

    for row in ws.iter_rows(min_row=2):
        if str(row[headers.index("Goal_ID")].value) == str(goal_id):
            row_no = row[0].row

            ws.cell(
                row=row_no,
                column=headers.index("Progress") + 1,
                value=progress
            )
            ws.cell(
                row=row_no,
                column=headers.index("Status") + 1,
                value="Completed" if progress >= 100 else "Active"
            )
            break

    save_db(wb)


# ============================================================
# ACHIEVEMENTS
# ============================================================

def has_achievement(user_id, name):
    return any(
        str(r.get("User_ID")) == str(user_id)
        and r.get("Achievement") == name
        for r in get_rows("Achievements")
    )


def unlock(user, name, description):
    if has_achievement(user["User_ID"], name):
        return

    append_row("Achievements", {
        "Achievement_ID": "ACH-" + str(uuid.uuid4())[:8],
        "User_ID": user["User_ID"],
        "Achievement": name,
        "Description": description,
        "Unlocked_Day": user["Day"],
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })

    st.toast(f"🏆 Achievement unlocked: {name}")


def check_achievements(user):
    checks = [
        (user["Day"] >= 7, "🌱 One Week Survivor", "Completed 7 days."),
        (user["Level"] >= 5, "⚡ Level 5", "Reached Level 5."),
        (user["Skills"] >= 80, "🧠 Skill Master", "Skills reached 80."),
        (user["Money"] >= 10000, "💰 Money Maker", "Balance crossed ₹10,000."),
        (user["Discipline"] >= 90, "🔥 Discipline Master", "Discipline reached 90."),
        (len(transactions_for(user["User_ID"])) >= 10, "📒 Money Manager", "Recorded 10 financial transactions."),
    ]

    for condition, name, desc in checks:
        if condition:
            unlock(user, name, desc)


# ============================================================
# SMART TIMETABLE
# ============================================================

def generate_timetable(user):
    career = str(user["Career"]).lower()
    hobby = str(user["Hobby"])

    if "software" in career or "developer" in career:
        career_task = "💻 Coding + DSA Practice"
    elif "data" in career:
        career_task = "📊 Python + Data Science"
    elif "doctor" in career or "medical" in career:
        career_task = "🩺 Medical Studies"
    elif "entrepreneur" in career or "business" in career:
        career_task = "💼 Business Development"
    elif "civil" in career:
        career_task = "🏗️ Civil Engineering Practice"
    elif "mechanical" in career:
        career_task = "⚙️ Engineering Design"
    elif "teacher" in career:
        career_task = "📚 Teaching + Subject Practice"
    elif "design" in career:
        career_task = "🎨 Design Practice"
    else:
        career_task = "🚀 Career Skill Development"

    exercise = (
        "🚶 Light Walk / Stretching"
        if user["Health"] < 50
        else "🏃 Exercise / Fitness"
    )

    energy_task = (
        "😴 Power Rest"
        if user["Energy"] < 50
        else "🧠 Personal Skill Practice"
    )

    hobby_task = (
        f"😊 Relaxation — {hobby}"
        if user["Happiness"] >= 50
        else f"❤️ Happiness Activity — {hobby}"
    )

    timetable = [
        ("06:30", "🌅 Wake Up + Morning Routine", "Health", "Start the day", "High"),
        ("07:00", exercise, "Health", "Health-based adjustment", "High"),
        ("08:00", "🍳 Breakfast", "Energy", "Maintain energy", "Medium"),
        ("09:00", "🎓 College / Education", "Education", "Academic progress", "High"),
        ("13:00", "🍱 Lunch + Rest", "Energy", "Recover energy", "Medium"),
        ("15:00", career_task, "Career", "Career-based adjustment", "High"),
        ("17:00", energy_task, "Energy", "Energy-based adjustment", "Medium"),
    ]

    if user["Part_Time_Job"] == "Yes":
        timetable.append(
            ("18:00", "💼 Part-Time Job", "Money", "Extra income", "Medium")
        )

    timetable += [
        ("19:00", hobby_task, "Happiness", "Happiness-based adjustment", "Medium"),
        ("20:00", "📝 Revision / Personal Project", "Skills", "Long-term growth", "High"),
        ("21:00", "👨‍👩‍👦 Family / Social Time", "Social", "Maintain relationships", "Medium"),
        ("22:00", "📋 Plan Tomorrow", "Discipline", "Build consistency", "Medium"),
        ("22:30", "🌙 Sleep", "Health", "Recovery", "High"),
    ]

    return timetable


def save_timetable(user):
    timetable = generate_timetable(user)

    for time_, activity, category, reason, priority in timetable:
        append_row("Timetable", {
            "Timetable_ID": "TIME-" + str(uuid.uuid4())[:8],
            "User_ID": user["User_ID"],
            "Day": user["Day"],
            "Time": time_,
            "Activity": activity,
            "Category": category,
            "Reason": reason,
            "Priority": priority,
            "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })

    return timetable


# ============================================================
# RANDOM EVENT + RECOMMENDATIONS
# ============================================================

def random_life_event(user):
    events = [
        ("🎁 You received a small reward!", "Money", 500),
        ("📚 You learned something useful!", "Skills", 5),
        ("😊 You had a positive day!", "Happiness", 5),
        ("⚡ You feel more energetic!", "Energy", 5),
        ("🤝 You made a new connection!", "Social", 5),
        ("😴 You feel a little tired.", "Energy", -5),
    ]

    message, stat, amount = random.choice(events)

    if stat == "Money":
        user["Money"] += amount
    else:
        change_stat(user, stat, amount)

    add_timeline(
        user,
        message,
        "Random Event",
        f"{stat}: {amount:+d}"
    )

    update_user(user)
    return message


def recommendations(user):
    result = []

    if user["Health"] < 50:
        result.append("❤️ Give more attention to healthy routines.")
    if user["Happiness"] < 50:
        result.append("😊 Spend some time on hobbies and relaxation.")
    if user["Energy"] < 40:
        result.append("🔋 Add more rest to your timetable.")
    if user["Skills"] < 50:
        result.append("🧠 Practice career-related skills.")
    if user["Discipline"] < 50:
        result.append("🔥 Build a consistent routine.")
    if user["Social"] < 50:
        result.append("🤝 Spend quality time with people you care about.")
    if user["Money"] < 1000:
        result.append("💰 Review your expenses and financial plan.")

    if not result:
        result.append("🚀 Your stats are balanced. Keep progressing!")

    return result


def complete_day(user):
    user = normalize_user(user)

    user["Day"] += 1
    user["Streak"] += 1

    change_stat(user, "Energy", -10)
    change_stat(user, "Health", -2)

    add_xp(user, 25)

    event = random_life_event(user)

    save_stats(user)
    check_achievements(user)
    update_user(user)

    return event


# ============================================================
# GEN AI — AI LIFE COACH
# ============================================================

def get_gemini_key():
    """Read Gemini API key from Streamlit Secrets or environment variables."""
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
        if key:
            return str(key).strip()
    except Exception:
        pass

    return os.environ.get("GEMINI_API_KEY", "").strip()


def _is_temporary_gemini_error(error):
    """Return True for temporary Gemini service/rate-limit errors."""
    message = str(error).upper()
    return any(code in message for code in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "500", "INTERNAL"))


def ai_life_coach(user, question):
    """Ask Google Gemini using the user's current simulator state.

    Includes a small exponential-backoff retry and a fallback model so a
    temporary Gemini 503/429 does not immediately break the Streamlit UI.
    """
    if genai is None:
        return None, "Google GenAI SDK is not installed. Run: pip install -U google-genai"

    api_key = get_gemini_key()
    if not api_key:
        return None, "GEMINI_API_KEY is not configured. Add it as a Colab environment variable before starting Streamlit."

    user = normalize_user(user)
    goals = goals_for(user["User_ID"])
    txns = transactions_for(user["User_ID"])

    expense_total = sum(
        safe_float(t.get("Amount"))
        for t in txns
        if str(t.get("Type", "")).lower() == "expense"
    )
    income_total = sum(
        safe_float(t.get("Amount"))
        for t in txns
        if str(t.get("Type", "")).lower() == "income"
    )

    goal_text = "\n".join(
        f"- {g.get('Goal')} | {g.get('Category')} | progress {safe_int(g.get('Progress'))}% | deadline {g.get('Deadline')}"
        for g in goals[-10:]
    ) or "No goals created yet."

    context = f"""
You are the AI Life Coach inside a Python Streamlit app called LIFE SIMULATOR.
Give practical, encouraging, age-appropriate and concise suggestions.
Use the player's simulator data to personalize the answer.
Do not make medical diagnoses, financial guarantees, investment recommendations,
or dangerous recommendations.

PLAYER PROFILE
Name: {user['Name']}
Age: {user['Age']}
Education: {user['Education']}
Career: {user['Career']}
Hobby: {user['Hobby']}
Part-time job: {user['Part_Time_Job']}
Day: {user['Day']}
Level: {user['Level']}
XP: {user['XP']}
Streak: {user['Streak']}

CURRENT STATS
Health: {user['Health']}/100
Happiness: {user['Happiness']}/100
Energy: {user['Energy']}/100
Skills: {user['Skills']}/100
Discipline: {user['Discipline']}/100
Social: {user['Social']}/100

FINANCE
Balance: ₹{safe_float(user['Money']):.2f}
Total recorded income: ₹{income_total:.2f}
Total recorded expenses: ₹{expense_total:.2f}

GOALS
{goal_text}

USER QUESTION
{question}

Respond in a friendly Hinglish/English style. Use headings and bullets when useful.
"""

    client = genai.Client(api_key=api_key)
    last_error = None

    GEMINI_MODEL = ["gemini-3.8-flash"]

    # Try the preferred model, then a fallback model. Each temporary error
    # gets a short exponential-backoff retry before moving on.
    for model_name in GEMINI_MODEL:
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=context,
                )
                text = getattr(response, "text", None)
                if text and text.strip():
                    return text.strip(), None
                last_error = "Gemini returned an empty response."
                break

            except Exception as e:
                last_error = e
                if not _is_temporary_gemini_error(e):
                    return None, f"Gemini AI request failed: {e}"

                if attempt < 2:
                    time.sleep(2 ** attempt)

    return None, (
        "⚠️ Gemini is temporarily unavailable after multiple attempts. "
        "Please wait a little and click ASK AI LIFE COACH again.\n\n"
        f"Last error: {last_error}"
    )


def ai_prompt_from_mode(mode, user):
    prompts = {
        "📅 Improve My Timetable":
            "Analyze my current profile and create a better one-day timetable. Explain why each major block is useful.",
        "🎯 Improve My Goals":
            "Review my current goals and suggest a realistic priority order and next actions for each goal.",
        "💸 Analyze My Expenses":
            "Analyze my financial state and recorded expenses. Suggest simple budgeting habits and categories to watch. Do not make investment recommendations.",
        "📊 Analyze My Life Stats":
            "Analyze my Health, Happiness, Energy, Skills, Discipline and Social stats. Tell me which areas need attention and give practical next actions.",
        "🚀 Give Me A Day Plan":
            "Create a motivating plan for my next simulator day using my current stats, career, hobby, goals and energy.",
    }
    return prompts.get(mode, "Give me a useful overview of my current Life Simulator progress and three practical next actions.")

# ============================================================
# SESSION
# ============================================================

if "user" not in st.session_state:
    st.session_state.user = None

if "page" not in st.session_state:
    st.session_state.page = "🏠 Dashboard"


# ============================================================
# HERO
# ============================================================

st.markdown("""
<div class="hero">
    <h1>🌍 LIFE SIMULATOR</h1>
    <p>Your decisions • Your progress • Your virtual life</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="feature-strip">
    <span class="badge">👤 Profiles</span>
    <span class="badge">🎯 Goals</span>
    <span class="badge">📊 Stats</span>
    <span class="badge">📅 Smart Timetable</span>
    <span class="badge">💰 Finance</span>
    <span class="badge">💸 Expense Manager</span>
    <span class="badge">🏆 Achievements</span>
    <span class="badge">🕒 Timeline</span>
    <span class="badge">⚡ XP & Levels</span>
    <span class="badge">🔥 Streaks</span>
    <span class="badge">🎲 Life Events</span>
    <span class="badge">🤖 AI Life Coach</span>
    <span class="badge">🤖 Recommendations</span>
    <span class="badge">💾 Excel Persistence</span>
</div>
""", unsafe_allow_html=True)

# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.markdown("## 🌍 LIFE SIMULATOR")

users = all_users()

if users:
    labels = [
        f"{u['Name']} • {u['User_ID']}"
        for u in users
    ]

    current_label = None
    if st.session_state.user:
        current_label = next(
            (
                f"{u['Name']} • {u['User_ID']}"
                for u in users
                if str(u["User_ID"]) == str(st.session_state.user["User_ID"])
            ),
            None
        )

    default_index = labels.index(current_label) if current_label in labels else 0

    selected_label = st.sidebar.selectbox(
        "👤 Profile",
        ["-- Select Profile --"] + labels,
        index=default_index + 1,
        key="profile_loader"
    )

    if selected_label != "-- Select Profile --":
        selected_id = selected_label.split(" • ")[-1]

        if not st.session_state.user or str(st.session_state.user["User_ID"]) != selected_id:
            st.session_state.user = find_user(selected_id)

else:
    st.sidebar.info("No profiles yet.")

st.sidebar.markdown("---")

menu = [
    "🏠 Dashboard",
    "👤 Profile",
    "🎯 Goals",
    "📅 Smart Timetable",
    "🎮 Activities",
    "💸 Expense Manager",
    "💰 Income Manager",
    "📊 Finance Dashboard",
    "🤖 AI Life Coach",
    "🏆 Achievements",
    "🕒 Timeline",
]

page = st.sidebar.radio(
    "📌 Choose Section",
    menu,
    index=menu.index(st.session_state.page)
)

st.session_state.page = page

if st.session_state.user:
    current = normalize_user(find_user(st.session_state.user["User_ID"]))
    st.session_state.user = current
else:
    current = None

# ============================================================
# CREATE PROFILE
# ============================================================

if current is None:

    st.markdown('<div class="section-title">🚀 Create Your Life</div>', unsafe_allow_html=True)

    st.info("Create a profile once. Your progress will be stored permanently in Life_Simulator.xlsx.")

    c1, c2 = st.columns(2)

    with c1:
        name = st.text_input("👤 Name")
        age = st.number_input("🎂 Age", min_value=0, max_value=100, value=0)
        education = st.text_input("🎓 Education", value="0")

    with c2:
        career_options = [
            "Write Career",
            "Software Engineer",
            "Data Scientist",
            "Doctor",
            "Entrepreneur",
            "Civil Engineer",
            "Mechanical Engineer",
            "Teacher",
            "Designer",
            "Other",
        ]

        career = st.selectbox(
            "💼 Career Goal",
            career_options
        )

        hobby = st.text_input("🎮 Hobby", value="Write Hobby")
        part_time = st.checkbox("💼 Enable Part-Time Job")

    if st.button("🚀 START MY LIFE", use_container_width=True):
        if not name.strip():
            st.error("Please enter your name.")
        elif career == "Write Career":
            st.warning("Please select a career.")
        else:
            new_user = create_user(
                name,
                age,
                education,
                career,
                hobby,
                part_time
            )
            st.session_state.user = new_user
            st.success("🎉 Profile created successfully!")
            st.rerun()

    st.stop()

user = current

# ============================================================
# DASHBOARD
# ============================================================

if page == "🏠 Dashboard":

    st.markdown(
        f'<div class="section-title">👋 Welcome, {user["Name"]}!</div>',
        unsafe_allow_html=True
    )

    st.write(
        f"📅 Day **{user['Day']}**  •  "
        f"⭐ Level **{user['Level']}**  •  "
        f"🔥 Streak **{user['Streak']} days**"
    )

    st.markdown("### 📊 LIFE OVERVIEW")

    cols = st.columns(4)

    metrics = [
        ("❤️ Health", user["Health"]),
        ("😊 Happiness", user["Happiness"]),
        ("⚡ Energy", user["Energy"]),
        ("💰 Money", f"₹{user['Money']:,.2f}"),
    ]

    for col, (label, value) in zip(cols, metrics):
        with col:
            st.markdown(
                f"""
                <div class="stat-card">
                    <div class="stat-label">{label}</div>
                    <div class="stat-value">{value}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.write("")

    cols = st.columns(4)

    metrics2 = [
        ("🧠 Skills", user["Skills"]),
        ("🔥 Discipline", user["Discipline"]),
        ("🤝 Social", user["Social"]),
        ("⭐ XP", user["XP"]),
    ]

    for col, (label, value) in zip(cols, metrics2):
        with col:
            st.markdown(
                f"""
                <div class="stat-card">
                    <div class="stat-label">{label}</div>
                    <div class="stat-value">{value}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown("### 📈 XP PROGRESS")

    required_xp = max(1, user["Level"] * 100)
    xp_progress = min(user["XP"] / required_xp, 1.0)

    st.progress(xp_progress)
    st.caption(f"{user['XP']} / {required_xp} XP")

    st.markdown("### 🤖 PERSONALIZED RECOMMENDATIONS")

    for rec in recommendations(user):
        st.info(rec)

    st.markdown("### ⚡ QUICK ACTIONS")

    q1, q2, q3, q4 = st.columns(4)

    with q1:
        if st.button("📅 Smart Timetable", use_container_width=True):
            st.session_state.page = "📅 Smart Timetable"
            st.rerun()

    with q2:
        if st.button("💸 Add Expense", use_container_width=True):
            st.session_state.page = "💸 Expense Manager"
            st.rerun()

    with q3:
        if st.button("🎲 Life Event", use_container_width=True):
            message = random_life_event(user)
            st.success(message)
            st.rerun()

    with q4:
        if st.button("🌅 Complete Day", use_container_width=True):
            message = complete_day(user)
            st.success(f"Day completed! {message}")
            st.rerun()

# ============================================================
# PROFILE
# ============================================================

elif page == "👤 Profile":

    st.markdown('<div class="section-title">👤 USER PROFILE</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        st.markdown(f"""
        <div class="card">
        <div class="card-title">👤 Personal</div>
        <b>User ID:</b> {user['User_ID']}<br>
        <b>Name:</b> {user['Name']}<br>
        <b>Age:</b> {user['Age']}<br>
        <b>Education:</b> {user['Education']}
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="card">
        <div class="card-title">🚀 Life Plan</div>
        <b>Career:</b> {user['Career']}<br>
        <b>Hobby:</b> {user['Hobby']}<br>
        <b>Part-Time:</b> {user['Part_Time_Job']}<br>
        <b>Current Day:</b> {user['Day']}
        </div>
        """, unsafe_allow_html=True)

    st.markdown("### ✏️ UPDATE PROFILE")

    c1, c2 = st.columns(2)

    with c1:
        education = st.text_input("Education", value=str(user["Education"]))
        career = st.text_input("Career", value=str(user["Career"]))

    with c2:
        hobby = st.text_input("Hobby", value=str(user["Hobby"]))
        part_time = st.radio(
            "Part-Time Job",
            ["Yes", "No"],
            index=0 if user["Part_Time_Job"] == "Yes" else 1,
            horizontal=True
        )

    if st.button("💾 SAVE PROFILE", use_container_width=True):
        user["Education"] = education
        user["Career"] = career
        user["Hobby"] = hobby
        user["Part_Time_Job"] = part_time
        update_user(user)
        st.success("Profile updated!")
        st.rerun()

# ============================================================
# GOALS
# ============================================================

elif page == "🎯 Goals":

    st.markdown('<div class="section-title">🎯 GOAL MANAGEMENT</div>', unsafe_allow_html=True)

    with st.form("goal_form"):
        goal = st.text_input("🎯 Goal Name")

        category = st.radio(
            "📂 Category",
            ["Education", "Career", "Health", "Money", "Skills", "Personal"],
            horizontal=True
        )

        target = st.number_input("🎯 Target", min_value=1, value=100)
        deadline = st.date_input("📅 Deadline")

        if st.form_submit_button("➕ ADD GOAL", use_container_width=True):
            if goal.strip():
                create_goal(user, goal, category, target, deadline)
                st.success("Goal added!")
                st.rerun()
            else:
                st.warning("Enter a goal name.")

    st.markdown("### 📋 YOUR GOALS")

    goals = goals_for(user["User_ID"])

    if not goals:
        st.info("No goals yet. Create your first goal above.")
    else:
        for g in goals:
            progress = safe_int(g.get("Progress"), 0)

            st.markdown(
                f"""
                <div class="card">
                    <div class="card-title">🎯 {g['Goal']}</div>
                    <div class="card-subtitle">
                        {g['Category']} • Deadline: {g['Deadline']} • Status: {g['Status']}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            new_progress = st.slider(
                "Progress",
                0,
                100,
                progress,
                key=f"goal_{g['Goal_ID']}"
            )

            st.progress(new_progress / 100)

            if st.button("💾 UPDATE", key=f"goal_update_{g['Goal_ID']}"):
                update_goal(g["Goal_ID"], new_progress)
                st.success("Progress updated!")
                st.rerun()

# ============================================================
# TIMETABLE
# ============================================================

elif page == "📅 Smart Timetable":

    st.markdown('<div class="section-title">📅 SMART AUTO TIMETABLE</div>', unsafe_allow_html=True)

    st.info(
        "Your timetable automatically adapts to career, health, energy, happiness and hobby."
    )

    timetable = generate_timetable(user)

    if st.button("💾 SAVE TODAY'S TIMETABLE", use_container_width=True):
        save_timetable(user)
        st.success("Timetable saved to Excel!")

    for time_, activity, category, reason, priority in timetable:
        st.markdown(
            f"""
            <div class="card">
                <div class="card-title">⏰ {time_} — {activity}</div>
                <div class="card-subtitle">
                    <span class="badge">{category}</span>
                    <span class="badge">{priority}</span>
                    {reason}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

# ============================================================
# ACTIVITIES
# ============================================================

elif page == "🎮 Activities":

    st.markdown('<div class="section-title">🎮 DAILY ACTIVITIES</div>', unsafe_allow_html=True)

    activities = [
        "📚 Study",
        "💻 Coding",
        "🏃 Exercise",
        "🧘 Meditation",
        "🎨 Hobby",
        "👨‍👩‍👦 Family Time",
        "🤝 Socialize",
        "😴 Rest",
        "💼 Work",
        "🧠 Skill Practice",
    ]

    activity = st.radio(
        "Choose Activity",
        activities,
        horizontal=False
    )

    if st.button("▶️ COMPLETE ACTIVITY", use_container_width=True):

        effects = {
            "📚 Study": ("Skills", 5),
            "💻 Coding": ("Skills", 7),
            "🏃 Exercise": ("Health", 5),
            "🧘 Meditation": ("Happiness", 4),
            "🎨 Hobby": ("Happiness", 5),
            "👨‍👩‍👦 Family Time": ("Social", 5),
            "🤝 Socialize": ("Social", 7),
            "😴 Rest": ("Energy", 10),
            "💼 Work": ("Money", 500),
            "🧠 Skill Practice": ("Skills", 6),
        }

        stat, amount = effects[activity]

        if stat == "Money":
            success, message = money_transaction(
                user,
                "Income",
                "Work",
                activity,
                amount
            )
            if not success:
                st.error(message)
            else:
                add_xp(user, 10)
                add_timeline(user, activity, "Activity", f"Money +{amount}")
                check_achievements(user)
                st.success(f"✅ {activity} completed!")
                st.rerun()
        else:
            change_stat(user, stat, amount)
            add_xp(user, 10)
            add_timeline(user, activity, "Activity", f"{stat} +{amount}")
            check_achievements(user)
            update_user(user)
            st.success(f"✅ {activity} completed!")
            st.rerun()

# ============================================================
# EXPENSE MANAGER
# ============================================================

elif page == "💸 Expense Manager":

    st.markdown('<div class="section-title">💸 EXPENSE MANAGER</div>', unsafe_allow_html=True)

    txns = transactions_for(user["User_ID"])

    expenses = [
        t for t in txns
        if t.get("Type") == "Expense"
    ]

    total_expense = sum(
        safe_float(t.get("Amount"))
        for t in expenses
    )

    c1, c2, c3 = st.columns(3)

    c1.metric("💰 Current Balance", f"₹{user['Money']:,.2f}")
    c2.metric("💸 Total Expenses", f"₹{total_expense:,.2f}")
    c3.metric("🧾 Expense Count", len(expenses))

    if user["Money"] < 1000:
        st.warning("⚠️ Low balance. Review your expenses.")

    st.markdown("### ➕ ADD EXPENSE")

    with st.form("expense_form"):

        categories = [
            "-- Select Category --",
            "Food",
            "Travel",
            "Education",
            "Shopping",
            "Bills",
            "Health",
            "Entertainment",
            "Personal",
            "Other",
        ]

        category = st.selectbox(
            "📂 Expense Category",
            categories
        )

        amount = st.number_input(
            "💵 Amount",
            min_value=1.0,
            value=100.0,
            step=50.0
        )

        description = st.text_input(
            "📝 Description"
        )

        expense_date = st.date_input(
            "📅 Date",
            value=date.today()
        )

        submit = st.form_submit_button(
            "💸 ADD EXPENSE",
            use_container_width=True
        )

        if submit:

            if category == "-- Select Category --":
                st.warning("Please select an expense category.")

            elif amount > user["Money"]:
                st.error(
                    f"❌ Insufficient balance! Available: ₹{user['Money']:,.2f}"
                )

            else:
                success, message = money_transaction(
                    user,
                    "Expense",
                    category,
                    f"{description} | {expense_date}",
                    amount
                )

                if success:
                    add_timeline(
                        user,
                        f"Expense: {description or category}",
                        "Finance",
                        f"-₹{amount:.2f}"
                    )
                    check_achievements(user)
                    st.success(message)
                    st.rerun()

    st.markdown("### 📊 CATEGORY SUMMARY")

    category_totals = {}

    for e in expenses:
        cat = e.get("Category") or "Other"
        category_totals[cat] = (
            category_totals.get(cat, 0)
            + safe_float(e.get("Amount"))
        )

    if category_totals:
        for cat, value in sorted(
            category_totals.items(),
            key=lambda x: x[1],
            reverse=True
        ):
            percentage = (
                value / total_expense
                if total_expense > 0
                else 0
            )

            st.write(f"**{cat}** — ₹{value:,.2f}")
            st.progress(min(percentage, 1.0))

    st.markdown("### 📜 EXPENSE HISTORY")

    if expenses:
        for e in reversed(expenses):
            st.markdown(
                f"""
                <div class="card">
                    <div class="card-title">
                        💸 ₹{safe_float(e.get('Amount')):,.2f}
                    </div>
                    <div class="card-subtitle">
                        {e.get('Category')} • {e.get('Description')} •
                        Balance: ₹{safe_float(e.get('Balance')):,.2f}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
    else:
        st.info("No expenses recorded yet.")

# ============================================================
# INCOME
# ============================================================

elif page == "💰 Income Manager":

    st.markdown('<div class="section-title">💰 INCOME MANAGER</div>', unsafe_allow_html=True)

    with st.form("income_form"):

        categories = [
            "-- Select Category --",
            "Salary",
            "Freelancing",
            "Part-Time Job",
            "Gift",
            "Reward",
            "Other",
        ]

        category = st.selectbox(
            "📂 Income Category",
            categories
        )

        amount = st.number_input(
            "💵 Amount",
            min_value=1.0,
            value=500.0,
            step=100.0
        )

        description = st.text_input("📝 Description")

        submit = st.form_submit_button(
            "➕ ADD INCOME",
            use_container_width=True
        )

        if submit:

            if category == "-- Select Category --":
                st.warning("Please select an income category.")
            else:
                success, message = money_transaction(
                    user,
                    "Income",
                    category,
                    description,
                    amount
                )

                if success:
                    add_timeline(
                        user,
                        f"Income: {description or category}",
                        "Finance",
                        f"+₹{amount:.2f}"
                    )
                    check_achievements(user)
                    st.success(message)
                    st.rerun()

# ============================================================
# FINANCE DASHBOARD
# ============================================================

elif page == "📊 Finance Dashboard":

    st.markdown('<div class="section-title">📊 FINANCE DASHBOARD</div>', unsafe_allow_html=True)

    txns = transactions_for(user["User_ID"])

    income = sum(
        safe_float(t.get("Amount"))
        for t in txns
        if t.get("Type") == "Income"
    )

    expense = sum(
        safe_float(t.get("Amount"))
        for t in txns
        if t.get("Type") == "Expense"
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("💰 Balance", f"₹{user['Money']:,.2f}")
    c2.metric("📈 Income", f"₹{income:,.2f}")
    c3.metric("📉 Expenses", f"₹{expense:,.2f}")
    c4.metric("📊 Net Flow", f"₹{income - expense:,.2f}")

    st.markdown("### 📜 MONEY HISTORY")

    for t in reversed(txns):
        icon = "🟢" if t.get("Type") in ["Income", "Initial Balance"] else "🔴"

        st.markdown(
            f"""
            <div class="card">
                <div class="card-title">
                    {icon} {t.get('Type')} — ₹{safe_float(t.get('Amount')):,.2f}
                </div>
                <div class="card-subtitle">
                    {t.get('Category')} • {t.get('Description')}
                    • Balance ₹{safe_float(t.get('Balance')):,.2f}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

# ============================================================
# AI LIFE COACH PAGE
# ============================================================

elif page == "🤖 AI Life Coach":

    st.markdown('<div class="section-title">🤖 AI LIFE COACH</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="card">
        <div class="card-title">🧠 Your Personal Simulator AI</div>
        <div class="card-subtitle">
            GenAI reads your current simulator stats, goals and finance summary and turns them into personalized suggestions.
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not get_gemini_key():
        st.warning("🔑 Gemini AI is ready, but no Gemini API key is configured.")
        st.code('import os\n# Set GEMINI_API_KEY in Streamlit Secrets (or a private environment)', language="python")
        st.caption("Never paste your real API key into public GitHub code. Use a private environment variable or Streamlit secret.")

    mode = st.selectbox(
        "✨ Quick AI Action",
        [
            "🚀 Give Me A Day Plan",
            "📅 Improve My Timetable",
            "🎯 Improve My Goals",
            "💸 Analyze My Expenses",
            "📊 Analyze My Life Stats",
            "💬 Ask My Own Question",
        ]
    )

    if mode == "💬 Ask My Own Question":
        question = st.text_area(
            "💬 Ask your AI Life Coach",
            placeholder="Example: How can I improve my skills and discipline in the next 7 simulator days?",
            height=130,
        )
    else:
        question = ai_prompt_from_mode(mode, user)
        st.info(question)

    if st.button("✨ ASK AI LIFE COACH", use_container_width=True):
        if not question.strip():
            st.warning("Please enter a question.")
        else:
            with st.spinner("🤖 AI is analyzing your virtual life..."):
                answer, error = ai_life_coach(user, question)

            if error:
                st.error(error)
            else:
                st.markdown("### 🧠 AI RESPONSE")
                st.markdown(answer)

    st.markdown("### 📌 CURRENT AI CONTEXT")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("❤️ Health", user["Health"])
    c2.metric("⚡ Energy", user["Energy"])
    c3.metric("🧠 Skills", user["Skills"])
    c4.metric("💰 Balance", f"₹{user['Money']:,.0f}")

# ============================================================
# ACHIEVEMENTS
# ============================================================

elif page == "🏆 Achievements":

    st.markdown('<div class="section-title">🏆 ACHIEVEMENTS</div>', unsafe_allow_html=True)

    achievements = [
        r for r in get_rows("Achievements")
        if str(r.get("User_ID")) == str(user["User_ID"])
    ]

    if achievements:
        for a in reversed(achievements):
            st.markdown(
                f"""
                <div class="card">
                    <div class="card-title">🏆 {a.get('Achievement')}</div>
                    <div class="card-subtitle">
                        {a.get('Description')} • Unlocked on Day {a.get('Unlocked_Day')}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
    else:
        st.info("No achievements unlocked yet. Keep progressing!")

# ============================================================
# TIMELINE
# ============================================================

elif page == "🕒 Timeline":

    st.markdown('<div class="section-title">🕒 LIFE TIMELINE</div>', unsafe_allow_html=True)

    timeline = [
        r for r in get_rows("Timeline")
        if str(r.get("User_ID")) == str(user["User_ID"])
    ]

    if timeline:
        for item in reversed(timeline):
            st.markdown(
                f"""
                <div class="timeline-item">
                    <b>Day {item.get('Day')} • {item.get('Time')}</b><br>
                    <b>{item.get('Activity')}</b><br>
                    <small>{item.get('Category')} • {item.get('Effect')}</small>
                </div>
                """,
                unsafe_allow_html=True
            )
    else:
        st.info("Your timeline is empty.")

# ============================================================
# FOOTER
# ============================================================

st.markdown("""
<div style="text-align:center; padding:30px 10px 10px;">
    <h3>🌍 LIFE SIMULATOR PRO</h3>
    <p>Your decisions • Your progress • Your virtual life</p>
    <small>💾 Persistent Excel Database • Built with Streamlit + OpenPyXL</small>
</div>
""", unsafe_allow_html=True)
